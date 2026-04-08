"""
land_cover_yangju 테이블을 개발DB → 운영DB로 복제하는 스크립트.
- 테이블 생성 (DDL) + 데이터 복사 (5000행 배치) + 인덱스 생성
- 운영DB에 이미 테이블이 있으면 확인 후 재생성
"""
import psycopg2
import sys
import time

DEV = {
    "host": "geo-spatial-hub.postgres.database.azure.com",
    "port": 6432,
    "dbname": "dde-water",
    "user": "waterviewer",
    "password": "dohwaviewer",
}

PROD = {
    "host": "geo-spatial-hub-prod.postgres.database.azure.com",
    "port": 6432,
    "dbname": "dde-water",
    "user": "postgres",
    "password": "dusrnxla123!@#",
}

TABLE = "land_cover_yangju"
BATCH_SIZE = 5000

DDL = """
CREATE TABLE IF NOT EXISTS public.land_cover_yangju (
    gid          INTEGER NOT NULL PRIMARY KEY,
    l1_code      VARCHAR(10),
    l1_name      VARCHAR(100),
    l2_code      VARCHAR(10),
    l2_name      VARCHAR(100),
    l3_code      VARCHAR(10),
    l3_name      VARCHAR(100),
    img_name     VARCHAR(100),
    img_date     VARCHAR(20),
    lu_info      TEXT,
    etc_info     TEXT,
    env_info     TEXT,
    for_info     TEXT,
    ud_info      TEXT,
    inx_num      VARCHAR(20),
    geom         geometry(MULTIPOLYGON, 5186) NOT NULL,
    created_at   TIMESTAMP
);
"""

INDEXES = [
    "CREATE INDEX IF NOT EXISTS idx_land_cover_yangju_geom ON public.land_cover_yangju USING gist (geom);",
    "CREATE INDEX IF NOT EXISTS idx_land_cover_yangju_l1_code ON public.land_cover_yangju USING btree (l1_code);",
    "CREATE INDEX IF NOT EXISTS idx_land_cover_yangju_l3_code ON public.land_cover_yangju USING btree (l3_code);",
    "CREATE INDEX IF NOT EXISTS idx_land_cover_yangju_inx_num ON public.land_cover_yangju USING btree (inx_num);",
]

COLUMNS = [
    "gid", "l1_code", "l1_name", "l2_code", "l2_name",
    "l3_code", "l3_name", "img_name", "img_date",
    "lu_info", "etc_info", "env_info", "for_info", "ud_info",
    "inx_num", "ST_AsEWKB(geom)", "created_at",
]

INSERT_COLS = [
    "gid", "l1_code", "l1_name", "l2_code", "l2_name",
    "l3_code", "l3_name", "img_name", "img_date",
    "lu_info", "etc_info", "env_info", "for_info", "ud_info",
    "inx_num", "geom", "created_at",
]


def main():
    t0 = time.time()

    print(f"[1/4] DEV DB 연결...")
    dev_conn = psycopg2.connect(**DEV, connect_timeout=15)
    dev_cur = dev_conn.cursor()

    print(f"[2/4] PROD DB 연결...")
    prod_conn = psycopg2.connect(**PROD, connect_timeout=15)
    prod_conn.autocommit = False
    prod_cur = prod_conn.cursor()

    # 기존 테이블 확인
    prod_cur.execute(
        "SELECT EXISTS(SELECT 1 FROM information_schema.tables "
        "WHERE table_schema='public' AND table_name=%s)", (TABLE,)
    )
    if prod_cur.fetchone()[0]:
        print(f"  [!] PROD에 {TABLE} 테이블이 이미 존재합니다. DROP 후 재생성합니다.")
        prod_cur.execute(f"DROP TABLE IF EXISTS public.{TABLE} CASCADE;")
        prod_conn.commit()

    # DDL
    print(f"[3/4] 테이블 생성...")
    prod_cur.execute(DDL)
    prod_conn.commit()

    # 데이터 복사
    print(f"[4/4] 데이터 복사 (배치 {BATCH_SIZE}행)...")
    select_sql = f"SELECT {', '.join(COLUMNS)} FROM public.{TABLE} ORDER BY gid"
    dev_cur.execute(select_sql)

    # geom은 index 15 — 해당 위치만 ST_GeomFromEWKB로 감싸기
    ph = []
    for i, col in enumerate(INSERT_COLS):
        ph.append("ST_GeomFromEWKB(%s)" if col == "geom" else "%s")
    insert_sql = f"INSERT INTO public.{TABLE} ({', '.join(INSERT_COLS)}) VALUES ({', '.join(ph)})"

    total = 0
    while True:
        rows = dev_cur.fetchmany(BATCH_SIZE)
        if not rows:
            break

        # geom(EWKB bytes)을 psycopg2 Binary로 감싸기
        batch = []
        for row in rows:
            row_list = list(row)
            if row_list[15] is not None:
                row_list[15] = psycopg2.Binary(row_list[15])
            batch.append(tuple(row_list))

        prod_cur.executemany(insert_sql, batch)
        total += len(rows)
        elapsed = time.time() - t0
        print(f"  {total:,} rows 복사됨 ({elapsed:.1f}s)")

    prod_conn.commit()

    # 인덱스 생성
    print(f"인덱스 생성 중...")
    for idx_sql in INDEXES:
        prod_cur.execute(idx_sql)
        print(f"  {idx_sql.split('idx_')[1].split(' ON')[0]} 완료")
    prod_conn.commit()

    # 검증
    prod_cur.execute(f"SELECT COUNT(*) FROM public.{TABLE}")
    prod_count = prod_cur.fetchone()[0]

    dev_cur.close()
    dev_conn.close()
    prod_cur.close()
    prod_conn.close()

    elapsed = time.time() - t0
    print(f"\n완료! PROD {TABLE}: {prod_count:,} rows ({elapsed:.1f}s)")
    if prod_count == 297048:
        print("[OK] DEV와 row count 일치 (297,048)")
    else:
        print(f"[!] DEV 297,048 vs PROD {prod_count:,} -- 확인 필요")


if __name__ == "__main__":
    main()
