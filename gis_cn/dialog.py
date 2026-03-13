import os
import logging

import pandas as pd
from qgis.PyQt import uic
from qgis.PyQt.QtWidgets import (
    QDialog, QFileDialog, QMessageBox,
    QInputDialog, QTableWidgetItem, QHeaderView,
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
    QLabel, QTableWidget, QAbstractItemView, QFrame,
    QStyledItemDelegate, QRadioButton, QLineEdit,
    QComboBox, QSplitter,
)
from qgis.PyQt.QtCore import Qt, QThread, pyqtSignal, QEvent
from qgis.core import (
    QgsProject, QgsVectorLayer, QgsWkbTypes, QgsMapLayerType
)

from .core.db_manager import get_soil_layer, get_land_cover_layer, get_soil_lc_intersection
from .core.local_data_handler import (
    load_local_soil, load_local_land_cover,
    get_local_soil_lc_intersection, ValidationError
)
from .core.spatial_ops import (
    intersect_layers,
    _build_result_layer, LEVEL_COLUMNS
)
from .core.cn_matcher import load_cn_table, apply_cn_to_layer, XLSX_PATH
from .core.result_calculator import calculate_results, calculate_grouped_results, export_results
from .core.watershed_group import load_groups, save_groups

logger = logging.getLogger(__name__)

FORM_CLASS, _ = uic.loadUiType(
    os.path.join(os.path.dirname(__file__), 'reference', 'cn_calculator.ui')
)

TAB_CALC    = 0   # 레이어 불러오기
TAB_MAPPING = 1   # 토지이용 재분류
TAB_CN_EDIT = 2   # CN값 편집
TAB_RECALC  = 3   # CN값 계산

# ── 카드 기반 다이얼로그 스타일 (reference/region_selector_dialog.py 기준) ──
DIALOG_STYLESHEET = """
* {
    font-family: 'Pretendard', 'Pretendard Variable', 'Malgun Gothic', 'Apple SD Gothic Neo', sans-serif;
    font-weight: 500;
    font-size: 13px;
}

/* ── 다이얼로그 ── */
QDialog { background-color: #f9fafb; }

/* ── 레이블 ── */
QLabel { color: #374151; font-weight: 500; }

/* ── 탭 위젯 ── */
QTabWidget::pane {
    border: none;
    background-color: #f9fafb;
}
QTabWidget > QWidget { background-color: #f9fafb; }
QTabBar {
    background-color: white;
    border-bottom: 1px solid #e5e7eb;
}
QTabBar::tab {
    background-color: transparent;
    color: #9ca3af;
    padding: 10px 22px;
    border: none;
    border-bottom: 2px solid transparent;
    font-size: 13px;
    font-weight: 500;
}
QTabBar::tab:selected {
    color: #1f2937;
    border-bottom: 2px solid #1f2937;
    font-weight: bold;
}
QTabBar::tab:hover:!selected { color: #6b7280; }

/* ── 입력 위젯 ── */
QComboBox {
    border: 1px solid #d1d5db;
    border-radius: 4px;
    padding: 6px 10px;
    background-color: #f9fafb;
    color: #374151;
}
QComboBox:hover { border-color: #9ca3af; background-color: white; }
QComboBox::drop-down { border: none; width: 20px; }
QComboBox QAbstractItemView {
    border: 1px solid #d1d5db;
    background-color: white;
    selection-background-color: #e5e7eb;
}

QLineEdit {
    border: 1px solid #d1d5db;
    border-radius: 4px;
    padding: 6px 10px;
    background-color: #f9fafb;
    color: #374151;
}
QLineEdit:hover  { border-color: #9ca3af; }
QLineEdit:focus  { border-color: #6b7280; background-color: white; }

/* ── 버튼 (기본: 아웃라인) ── */
QPushButton {
    border: 1px solid #d1d5db;
    border-radius: 6px;
    padding: 6px 14px;
    color: #374151;
    background-color: white;
    font-size: 13px;
}
QPushButton:hover {
    background-color: #f3f4f6;
    border-color: #9ca3af;
}
QPushButton:pressed { background-color: #e5e7eb; }
QPushButton:disabled {
    background-color: #f3f4f6;
    color: #9ca3af;
    border-color: #e5e7eb;
}

/* ── 주요 액션 버튼 (어두운 배경) ── */
QPushButton#btnRun,
QPushButton#btnSaveCn,
QPushButton#btnExportResult1 {
    background-color: #1f2937;
    color: white;
    border: none;
    font-weight: 600;
}
QPushButton#btnRun:hover,
QPushButton#btnSaveCn:hover,
QPushButton#btnExportResult1:hover {
    background-color: #374151;
}
QPushButton#btnRun:disabled { background-color: #9ca3af; }

/* ── 진행바 ── */
QProgressBar {
    background-color: #e5e7eb;
    border: none;
    border-radius: 3px;
    text-align: center;
    font-size: 10px;
    color: #6b7280;
}
QProgressBar::chunk {
    background-color: #1f2937;
    border-radius: 3px;
}

/* ── 텍스트 에디터 (로그) ── */
QTextEdit {
    border: 1px solid #e5e7eb;
    border-radius: 6px;
    background-color: white;
    padding: 6px;
    color: #374151;
}

/* ── 테이블 ── */
QTableWidget {
    border: 1px solid #e5e7eb;
    border-radius: 6px;
    background-color: white;
    gridline-color: #f3f4f6;
    selection-background-color: #f3f4f6;
    selection-color: #1f2937;
}
QTableWidget::item { padding: 3px 6px; color: #374151; }
QTableWidget::item:selected {
    background-color: #f3f4f6;
    color: #1f2937;
}
QHeaderView::section {
    background-color: #f9fafb;
    border: none;
    border-bottom: 1px solid #e5e7eb;
    border-right: 1px solid #e5e7eb;
    padding: 5px 8px;
    font-size: 12px;
    font-weight: 600;
    color: #6b7280;
}

/* ── 라디오/체크 버튼 ── */
QRadioButton, QCheckBox {
    color: #374151;
    spacing: 7px;
}
QRadioButton::indicator {
    width: 15px;
    height: 15px;
    border-radius: 8px;
}
QRadioButton::indicator:checked {
    background-color: #1f2937;
    border: 2px solid #1f2937;
}
QRadioButton::indicator:unchecked {
    background-color: white;
    border: 2px solid #d1d5db;
}
QCheckBox::indicator {
    width: 14px;
    height: 14px;
    border-radius: 3px;
}
QCheckBox::indicator:checked {
    background-color: #1f2937;
    border: 2px solid #1f2937;
}
QCheckBox::indicator:unchecked {
    background-color: white;
    border: 2px solid #d1d5db;
}

/* ── 스크롤바 ── */
QScrollArea { border: none; background-color: transparent; }
QScrollBar:vertical {
    background-color: #f3f4f6;
    width: 8px;
    border-radius: 4px;
    margin: 2px;
}
QScrollBar::handle:vertical {
    background-color: #9ca3af;
    border-radius: 4px;
    min-height: 24px;
}
QScrollBar::handle:vertical:hover { background-color: #6b7280; }
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0px; }
QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical { background: none; }
QScrollBar:horizontal {
    background-color: #f3f4f6;
    height: 8px;
    border-radius: 4px;
    margin: 2px;
}
QScrollBar::handle:horizontal {
    background-color: #9ca3af;
    border-radius: 4px;
    min-width: 24px;
}
QScrollBar::handle:horizontal:hover { background-color: #6b7280; }
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal { width: 0px; }
QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal { background: none; }
"""


# ──────────────────────────────────────────────────────────────────────────────
# 테이블 Enter키 → 다음 셀 이동 delegate
# ──────────────────────────────────────────────────────────────────────────────

class _NextCellDelegate(QStyledItemDelegate):
    """Enter/Return 키로 편집 완료 후 다음 셀로 자동 이동."""

    def eventFilter(self, editor, event):
        if event.type() == QEvent.KeyPress and event.key() in (Qt.Key_Return, Qt.Key_Enter):
            tbl = self.parent()
            self.commitData.emit(editor)
            self.closeEditor.emit(editor, QStyledItemDelegate.NoHint)
            row = tbl.currentRow()
            col = tbl.currentColumn() + 1
            if col >= tbl.columnCount():
                col = 0
                row += 1
            if row < tbl.rowCount():
                tbl.setCurrentCell(row, col)
                tbl.edit(tbl.model().index(row, col))
            return True
        return super().eventFilter(editor, event)


class _MappingComboDelegate(QStyledItemDelegate):
    """재분류 이름 열 전용 — CN표 토지이용분류 드롭다운 + Enter키 이동."""

    def createEditor(self, parent, option, index):
        combo = QComboBox(parent)
        combo.setEditable(True)
        # 매 호출 시 최신 CN표 목록 로드
        dialog = self.parent()
        while dialog and not isinstance(dialog, CnCalculatorDialog):
            dialog = dialog.parent()
        if dialog:
            names = dialog._get_cn_land_use_names()
            combo.addItem('')
            combo.addItems(names)
        combo.setStyleSheet(
            "QComboBox { border: 1px solid #d1d5db; border-radius: 3px; padding: 1px 4px; }"
        )
        return combo

    def setEditorData(self, editor, index):
        val = index.data(Qt.EditRole) or ''
        idx = editor.findText(str(val))
        if idx >= 0:
            editor.setCurrentIndex(idx)
        else:
            editor.setCurrentText(str(val))

    def setModelData(self, editor, model, index):
        model.setData(index, editor.currentText(), Qt.EditRole)

    def updateEditorGeometry(self, editor, option, index):
        editor.setGeometry(option.rect)

    def eventFilter(self, obj, event):
        if event.type() == QEvent.KeyPress and event.key() in (Qt.Key_Return, Qt.Key_Enter):
            self.commitData.emit(obj)
            self.closeEditor.emit(obj)
            # Move to next row, same column
            table = self.parent()
            if table:
                cur = table.currentIndex()
                next_row = cur.row() + 1
                if next_row < table.model().rowCount():
                    table.setCurrentIndex(table.model().index(next_row, cur.column()))
            return True
        return super().eventFilter(obj, event)


# ──────────────────────────────────────────────────────────────────────────────
# 백그라운드 워커 스레드
# ──────────────────────────────────────────────────────────────────────────────

class CnWorker(QThread):
    progress    = pyqtSignal(int, str)      # (퍼센트, 메시지)
    layer_ready = pyqtSignal(object, str)   # (QgsVectorLayer, 레이어명) — 단계별 즉시 표시
    finished    = pyqtSignal(object, list)  # (cn_input_layer, fail_list)
    error       = pyqtSignal(str)

    def __init__(self, input_layer, name_field, level,
                 data_source='db', soil_path=None, lc_path=None):
        super().__init__()
        self.input_layer = input_layer
        self.name_field  = name_field
        self.level       = level
        self.data_source = data_source
        self.soil_path   = soil_path
        self.lc_path     = lc_path

    def run(self):
        try:
            import processing as proc

            # ── 소유역 전체 영역 Union ─────────────────────────────────────
            self.progress.emit(3, "소유역 영역 계산 중...")
            dissolved = proc.run("native:dissolve", {
                'INPUT': self.input_layer,
                'FIELD': [],
                'OUTPUT': 'memory:'
            })['OUTPUT']

            if self.data_source == 'local':
                # ── ① 로컬 토양군 Clip ──────────────────────────────────
                self.progress.emit(8, "① 로컬 토양군 Clip 중...")
                soil_clipped = load_local_soil(self.soil_path, dissolved)
                self.progress.emit(22, f"   → {soil_clipped.featureCount()}개 피처")
                self.layer_ready.emit(soil_clipped, "토양군_clip")

                # ── ② 로컬 토지피복도 Clip + Dissolve ───────────────────
                self.progress.emit(26, f"② 로컬 토지피복도 Clip 중 (level={self.level})...")
                lc_clipped = load_local_land_cover(self.lc_path, dissolved, self.level)
                self.progress.emit(40, f"   → {lc_clipped.featureCount()}개 피처")
                self.layer_ready.emit(lc_clipped, "토지피복도_clip")

                # ── ③ 로컬 Intersection (토양군 × 토지피복도) ───────────
                self.progress.emit(44, "③ Intersection(토양군 × 토지피복도) 중...")
                soil_lc = get_local_soil_lc_intersection(soil_clipped, lc_clipped, self.level)
                self.progress.emit(68, f"   → {soil_lc.featureCount()} features")
                self.layer_ready.emit(soil_lc, "토양군_토지피복_교차")

                # ── ④ Intersection × 소유역계 ──────────────────────────
                self.progress.emit(72, "④ Intersection × 소유역계 중...")
                final_intersect = intersect_layers(soil_lc, self.input_layer, "최종교차_raw")
                self.progress.emit(95, f"   → {final_intersect.featureCount()} features")
                self.finished.emit(final_intersect, [])

            else:
                # ── 기존 DB 모드 (변경 없음) ────────────────────────────
                union_geom = None
                for feat in dissolved.getFeatures():
                    union_geom = feat.geometry()
                    break
                if not union_geom:
                    raise ValueError("소유역 레이어에서 영역을 읽을 수 없습니다.")

                srid = self.input_layer.crs().postgisSrid()
                if srid <= 0:
                    srid = 5186
                polygon_wkt = union_geom.asWkt()

                # ── ① DB: 토양군 Clip (ST_Intersection) ──────────────────────
                self.progress.emit(8, "① PostGIS 토양군 추출 (Clip 포함)...")
                soil_clipped = get_soil_layer(polygon_wkt, srid)
                self.progress.emit(22, f"   → {soil_clipped.featureCount()}개 피처")
                self.layer_ready.emit(soil_clipped, "토양군_clip")

                # ── ② DB: 토지피복도 Clip (ST_Intersection) ───────────────────
                self.progress.emit(26, f"② PostGIS 토지피복도 추출 (Clip 포함, level={self.level})...")
                lc_clipped = get_land_cover_layer(polygon_wkt, srid, self.level)
                self.progress.emit(40, f"   → {lc_clipped.featureCount()}개 피처")
                self.layer_ready.emit(lc_clipped, "토지피복도_clip")

                # ── ③ DB: 토양군 × 토지피복도 Intersection (PostGIS) ──────────
                self.progress.emit(44, "③ PostGIS Intersection(토양군 × 토지피복도) 중...")
                soil_lc = get_soil_lc_intersection(polygon_wkt, srid, self.level)
                self.progress.emit(68, f"   → {soil_lc.featureCount()} features")
                self.layer_ready.emit(soil_lc, "토양군_토지피복_교차")

                # ── ④ QGIS Intersection × 소유역계 ───────────────────────────
                self.progress.emit(72, "④ Intersection × 소유역계 중...")
                final_intersect = intersect_layers(soil_lc, self.input_layer, "최종교차_raw")
                self.progress.emit(95, f"   → {final_intersect.featureCount()} features")
                self.finished.emit(final_intersect, [])

        except Exception as e:
            logger.exception("워커 오류")
            self.error.emit(str(e))


# ──────────────────────────────────────────────────────────────────────────────
# 메인 다이얼로그
# ──────────────────────────────────────────────────────────────────────────────

class CnCalculatorDialog(QDialog, FORM_CLASS):

    def __init__(self, iface, parent=None):
        super().__init__(parent)
        self.iface  = iface
        self.worker = None
        self.input_layer = None
        self._cn_table_loaded    = False
        self._mapping_tab_loaded = False
        self._final_intersect_layer = None
        self._last_level        = 'l1'
        self._last_name_field   = ''
        self._cn_ref_loaded     = False
        self._cn_ref_dirty      = False
        self.setupUi(self)
        self._init_ui()
        self._connect_signals()

    # ── 초기화 ────────────────────────────────────────────────────────────────

    def _init_ui(self):
        self.setStyleSheet(DIALOG_STYLESHEET)
        self.progressBar.setValue(0)
        self._refresh_layer_list()
        self._toggle_input_mode()
        self._setup_data_source_card()
        self._init_cn_table_widget()
        self._setup_mapping_tab()
        self._enhance_recalc_tab()

    def _refresh_layer_list(self):
        self.cmbLayer.clear()
        for layer in QgsProject.instance().mapLayers().values():
            if (layer.type() == QgsMapLayerType.VectorLayer and
                    layer.geometryType() == QgsWkbTypes.PolygonGeometry):
                self.cmbLayer.addItem(layer.name(), layer.id())

    def _refresh_recalc_layer_list(self):
        self.cmbRecalcLayer.clear()
        for layer in QgsProject.instance().mapLayers().values():
            if (layer.type() == QgsMapLayerType.VectorLayer and
                    layer.name() == "CN값_input"):
                self.cmbRecalcLayer.addItem(layer.name(), layer.id())

    def _init_cn_table_widget(self):
        tbl = self.tblCnValues
        tbl.horizontalHeader().setSectionResizeMode(0, QHeaderView.Interactive)
        tbl.horizontalHeader().setMinimumSectionSize(40)
        for col in range(1, tbl.columnCount()):
            tbl.horizontalHeader().setSectionResizeMode(col, QHeaderView.ResizeToContents)
        tbl.verticalHeader().setDefaultSectionSize(32)
        tbl.setItemDelegate(_NextCellDelegate(tbl))

    # ── 시그널 연결 ───────────────────────────────────────────────────────────

    def _connect_signals(self):
        # Tab 1
        self.rbFile.toggled.connect(self._toggle_input_mode)
        self.btnBrowse.clicked.connect(self._browse_file)
        self.leFilePath.editingFinished.connect(self._on_file_path_changed)
        self.cmbLayer.currentIndexChanged.connect(self._on_layer_changed)
        self.btnRun.clicked.connect(self._run)
        self.btnClose.clicked.connect(self.close)
        # Tab 2
        self.btnAddRow.clicked.connect(self._cn_add_row)
        self.btnDeleteRow.clicked.connect(self._cn_delete_row)
        self.btnAddColumn.clicked.connect(self._cn_add_column)
        self.btnReloadCn.clicked.connect(self._cn_reload)
        self.btnImportCn.clicked.connect(self._cn_import)
        self.btnSaveCn.clicked.connect(self._cn_export)
        # Tab 3 (CN값 계산)
        self.btnApplyCn.clicked.connect(self._apply_cn_calc)
        self.btnOutputDir.clicked.connect(self._browse_output_dir)
        self.btnExportResult1.clicked.connect(self._export_results)
        self.btnExportResult2.setVisible(False)
        # Tab 2: 토지이용 재분류
        self.btnMappingLoadLayer.clicked.connect(self._mapping_load_from_layer)
        self.btnMappingAddRow.clicked.connect(self._mapping_add_row)
        self.btnMappingDeleteRow.clicked.connect(self._mapping_delete_row)
        self.btnMappingSave.clicked.connect(self._mapping_save)
        self.btnMappingClear.clicked.connect(self._mapping_clear)
        # Tab 1: CN표 참조 패널
        self.tblCnRef.doubleClicked.connect(self._on_cn_ref_double_click)
        self.leCnRefSearch.textChanged.connect(self._filter_cn_ref_table)
        # Tab 1: 매핑 유효성 검증
        self.tblMapping.cellChanged.connect(self._validate_mapping_cell)
        # Tab 2: CN표 편집 시 참조 패널 dirty 플래그
        self.tblCnValues.cellChanged.connect(lambda: setattr(self, '_cn_ref_dirty', True))
        # 다음 단계 버튼
        self.btnNextStep0.clicked.connect(lambda: self.tabWidget.setCurrentIndex(TAB_MAPPING))
        self.btnNextStep1.clicked.connect(lambda: self.tabWidget.setCurrentIndex(TAB_CN_EDIT))
        self.btnNextStep2.clicked.connect(lambda: self.tabWidget.setCurrentIndex(TAB_RECALC))
        # 탭 전환
        self.tabWidget.currentChanged.connect(self._on_tab_changed)

    # ── 입력 처리 ─────────────────────────────────────────────────────────────

    def _toggle_input_mode(self):
        file_mode = self.rbFile.isChecked()
        self.leFilePath.setEnabled(file_mode)
        self.btnBrowse.setEnabled(file_mode)
        self.cmbLayer.setEnabled(not file_mode)
        if not file_mode:
            self._refresh_layer_list()
            self._populate_name_fields_from_layer()

    def _browse_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "소유역 레이어 선택", "",
            "벡터 파일 (*.shp *.gpkg *.geojson);;모든 파일 (*)"
        )
        if path:
            self.leFilePath.setText(path)
            self._on_file_path_changed()

    def _on_file_path_changed(self):
        path = self.leFilePath.text().strip()
        if path and os.path.exists(path):
            layer = QgsVectorLayer(path, "_check", "ogr")
            if layer.isValid():
                self._populate_name_fields(layer)

    def _on_layer_changed(self):
        if self.rbLayer.isChecked():
            self._populate_name_fields_from_layer()

    def _on_tab_changed(self, index):
        if index == TAB_CN_EDIT and not self._cn_table_loaded:
            self._load_cn_to_table()
        elif index == TAB_RECALC:
            self._refresh_recalc_layer_list()
        elif index == TAB_MAPPING:
            if not self._mapping_tab_loaded:
                self._mapping_load_saved()
            if not self._cn_ref_loaded:
                self._load_cn_ref_table()
            elif self._cn_ref_dirty:
                self._sync_cn_ref_from_edit()

    def _populate_name_fields_from_layer(self):
        layer_id = self.cmbLayer.currentData()
        if layer_id:
            layer = QgsProject.instance().mapLayer(layer_id)
            if layer:
                self._populate_name_fields(layer)

    def _populate_name_fields(self, layer: QgsVectorLayer):
        self.cmbNameField.clear()
        for field in layer.fields():
            self.cmbNameField.addItem(field.name())

    # ── 데이터 소스 선택 ─────────────────────────────────────────────────────

    def _setup_data_source_card(self):
        """Tab 0에 데이터 소스 선택 카드를 동적으로 추가."""
        tab0 = self.tabWidget.widget(TAB_CALC)
        layout = tab0.layout()
        if layout is None:
            return

        card = QFrame()
        card.setObjectName("dataSourceCard")
        card.setStyleSheet(
            "QFrame#dataSourceCard {"
            "  background-color: white; border: 1px solid #e5e7eb; border-radius: 8px;"
            "}"
        )
        card_layout = QVBoxLayout(card)
        card_layout.setSpacing(8)
        card_layout.setContentsMargins(16, 14, 16, 14)

        # 카드 헤더
        title = QLabel("데이터 소스")
        title.setStyleSheet(
            "font-size: 14px; font-weight: bold; color: #374151; border: none;"
        )
        card_layout.addWidget(title)

        # 라디오 버튼 행
        radio_row = QHBoxLayout()
        from qgis.PyQt.QtWidgets import QButtonGroup
        self.rbSourceDB = QRadioButton("DB에서 가져오기")
        self.rbSourceLocal = QRadioButton("로컬 파일 사용")
        self.rbSourceDB.setChecked(True)
        self._source_group = QButtonGroup(self)
        self._source_group.addButton(self.rbSourceDB)
        self._source_group.addButton(self.rbSourceLocal)
        radio_row.addWidget(self.rbSourceDB)
        radio_row.addWidget(self.rbSourceLocal)
        radio_row.addStretch()
        card_layout.addLayout(radio_row)

        # 로컬 파일 입력 위젯 (초기 숨김)
        self._local_file_widget = QWidget()
        local_layout = QVBoxLayout(self._local_file_widget)
        local_layout.setContentsMargins(0, 4, 0, 0)
        local_layout.setSpacing(6)

        # 토양군 파일
        soil_row = QHBoxLayout()
        soil_lbl = QLabel("토양군 파일:")
        soil_lbl.setFixedWidth(80)
        soil_lbl.setStyleSheet("border: none;")
        self.leSoilPath = QLineEdit()
        self.leSoilPath.setPlaceholderText("토양군 SHP/GPKG 파일 선택...")
        self.btnBrowseSoil = QPushButton("찾아보기")
        soil_row.addWidget(soil_lbl)
        soil_row.addWidget(self.leSoilPath, 1)
        soil_row.addWidget(self.btnBrowseSoil)
        local_layout.addLayout(soil_row)

        # 토지피복도 파일
        lc_row = QHBoxLayout()
        lc_lbl = QLabel("토지피복도:")
        lc_lbl.setFixedWidth(80)
        lc_lbl.setStyleSheet("border: none;")
        self.leLcPath = QLineEdit()
        self.leLcPath.setPlaceholderText("토지피복도 SHP/GPKG 파일 선택...")
        self.btnBrowseLc = QPushButton("찾아보기")
        lc_row.addWidget(lc_lbl)
        lc_row.addWidget(self.leLcPath, 1)
        lc_row.addWidget(self.btnBrowseLc)
        local_layout.addLayout(lc_row)

        card_layout.addWidget(self._local_file_widget)
        self._local_file_widget.setVisible(False)

        # Insert card into Tab 0 layout - insert near top (index 0 or after header)
        layout.insertWidget(0, card)

        # Connect signals
        self.rbSourceLocal.toggled.connect(self._toggle_data_source)
        self.btnBrowseSoil.clicked.connect(lambda: self._browse_local_file(self.leSoilPath, "토양군"))
        self.btnBrowseLc.clicked.connect(lambda: self._browse_local_file(self.leLcPath, "토지피복도"))

    def _toggle_data_source(self, local_checked):
        """데이터 소스 전환 (DB ↔ 로컬)."""
        self._local_file_widget.setVisible(local_checked)

    def _browse_local_file(self, line_edit, label):
        """로컬 파일 찾아보기 다이얼로그."""
        path, _ = QFileDialog.getOpenFileName(
            self, f"{label} 파일 선택", "",
            "벡터 파일 (*.shp *.gpkg);;모든 파일 (*)"
        )
        if path:
            line_edit.setText(path)

    # ── 실행 ──────────────────────────────────────────────────────────────────

    def _get_level(self) -> str:
        if self.rbL2.isChecked(): return 'l2'
        if self.rbL3.isChecked(): return 'l3'
        return 'l1'

    def _get_input_layer(self) -> QgsVectorLayer:
        if self.rbFile.isChecked():
            path = self.leFilePath.text().strip()
            if not path:
                raise ValueError("파일 경로를 입력하세요.")
            layer = QgsVectorLayer(path, "소유역계", "ogr")
            if not layer.isValid():
                raise ValueError(f"레이어 로드 실패: {path}")
            return layer
        else:
            layer_id = self.cmbLayer.currentData()
            if not layer_id:
                raise ValueError("레이어를 선택하세요.")
            layer = QgsProject.instance().mapLayer(layer_id)
            if not layer:
                raise ValueError("선택한 레이어를 찾을 수 없습니다.")
            return layer

    def _run(self):
        if not self.cmbNameField.currentText():
            QMessageBox.warning(self, "입력 오류", "소유역명 컬럼을 선택하세요.")
            return
        try:
            input_layer = self._get_input_layer()
        except ValueError as e:
            QMessageBox.warning(self, "입력 오류", str(e))
            return

        # 입력 레이어 저장 (유역합성 UI에서 소유역명 조회용)
        self.input_layer = input_layer

        # 데이터 소스 확인
        data_source = 'local' if self.rbSourceLocal.isChecked() else 'db'
        soil_path = None
        lc_path = None
        if data_source == 'local':
            soil_path = self.leSoilPath.text().strip()
            lc_path = self.leLcPath.text().strip()
            if not soil_path:
                QMessageBox.warning(self, "입력 오류", "토양군 파일 경로를 입력하세요.")
                return
            if not lc_path:
                QMessageBox.warning(self, "입력 오류", "토지피복도 파일 경로를 입력하세요.")
                return

        self.btnRun.setEnabled(False)
        self.progressBar.setValue(0)
        self.txtLog.clear()
        self._log("▶ 레이어 불러오기 시작")
        if data_source == 'local':
            self._log("  모드: 로컬 파일")
        else:
            self._log("  모드: PostGIS DB")
        self._log("  단계별 중간 레이어가 캔버스에 즉시 표시됩니다.")
        self._log("─" * 44)

        self.worker = CnWorker(
            input_layer,
            self.cmbNameField.currentText(),
            self._get_level(),
            data_source=data_source,
            soil_path=soil_path,
            lc_path=lc_path,
        )
        self.worker.progress.connect(self._on_progress)
        self.worker.layer_ready.connect(self._on_layer_ready)
        self.worker.finished.connect(self._on_finished)
        self.worker.error.connect(self._on_error)
        self.worker.start()

    # ── 워커 콜백 ─────────────────────────────────────────────────────────────

    def _on_progress(self, pct: int, msg: str):
        self.progressBar.setValue(pct)
        self._log(msg)

    def _on_layer_ready(self, layer, name: str):
        """단계별 중간 레이어를 QGIS 캔버스에 즉시 추가."""
        layer.setName(name)
        QgsProject.instance().addMapLayer(layer)
        self._log(f"   ✦ 레이어 추가됨: [{name}]")

    def _on_finished(self, final_intersect_layer, _fail_list):
        # 교차 레이어 저장 (CN값 계산 탭에서 사용)
        self._final_intersect_layer = final_intersect_layer
        self._last_level      = self.worker.level
        self._last_name_field = self.worker.name_field

        self.progressBar.setValue(100)
        self._log("━" * 44)
        self._log("✔ 완료! 중간 레이어가 캔버스에 추가되었습니다.")
        self._log("  → [CN값 계산] 탭으로 이동하여 CN값을 계산하세요.")
        self.btnRun.setEnabled(True)

    def _on_error(self, msg: str):
        self._log(f"[오류] {msg}")
        QMessageBox.critical(self, "오류 발생", msg)
        self.btnRun.setEnabled(True)
        self.progressBar.setValue(0)

    # ── CN값 편집 탭 ──────────────────────────────────────────────────────────

    def _load_cn_to_table(self):
        try:
            df = load_cn_table(XLSX_PATH)
        except FileNotFoundError as e:
            QMessageBox.warning(self, "파일 없음", str(e))
            return
        except Exception as e:
            QMessageBox.critical(self, "로드 오류", str(e))
            return
        self._df_to_table(df)
        self._cn_table_loaded = True

    def _df_to_table(self, df: pd.DataFrame):
        tbl = self.tblCnValues
        tbl.blockSignals(True)
        cols = list(df.columns)
        tbl.setColumnCount(len(cols))
        tbl.setHorizontalHeaderLabels(cols)
        tbl.setRowCount(len(df))
        for r, (_, row) in enumerate(df.iterrows()):
            for c, val in enumerate(row):
                text = "" if pd.isna(val) else str(val)
                if c > 0 and text:
                    try:
                        text = str(int(float(text)))
                    except (ValueError, OverflowError):
                        pass
                item = QTableWidgetItem(text)
                if c > 0:
                    item.setTextAlignment(Qt.AlignCenter)
                tbl.setItem(r, c, item)
        tbl.horizontalHeader().setSectionResizeMode(0, QHeaderView.Interactive)
        for c in range(1, tbl.columnCount()):
            tbl.horizontalHeader().setSectionResizeMode(c, QHeaderView.ResizeToContents)
        tbl.resizeColumnToContents(0)
        tbl.setColumnWidth(0, max(tbl.columnWidth(0), 160))
        tbl.blockSignals(False)

    def _get_cn_table_from_widget(self) -> pd.DataFrame:
        """현재 CN값 편집 테이블 위젯 내용을 DataFrame으로 반환."""
        tbl = self.tblCnValues
        if tbl.rowCount() == 0:
            return pd.DataFrame()
        headers = []
        for c in range(tbl.columnCount()):
            h = tbl.horizontalHeaderItem(c)
            headers.append(h.text() if h else f"col_{c}")
        data = []
        for r in range(tbl.rowCount()):
            row_data = []
            for c in range(tbl.columnCount()):
                item = tbl.item(r, c)
                row_data.append(item.text() if item else "")
            data.append(row_data)
        df = pd.DataFrame(data, columns=headers)
        for col in headers[1:]:
            df[col] = pd.to_numeric(df[col], errors='coerce')
        return df

    def _get_cn_land_use_names(self) -> list:
        """CN값 테이블에서 토지이용분류 목록 반환."""
        # 위젯 테이블에 데이터가 있으면 거기서 로드
        tbl = self.tblCnValues
        if tbl.rowCount() > 0:
            names = []
            for r in range(tbl.rowCount()):
                item = tbl.item(r, 0)
                if item and item.text().strip():
                    names.append(item.text().strip())
            if names:
                return names
        # 파일에서 폴백 로드
        try:
            from .core.cn_matcher import load_cn_table, XLSX_PATH
            df = load_cn_table(XLSX_PATH)
            return [str(v) for v in df.iloc[:, 0].dropna().tolist() if str(v).strip()]
        except Exception:
            return []

    def _cn_export(self):
        """현재 CN값 테이블을 사용자가 지정한 파일로 내보내기."""
        tbl = self.tblCnValues
        if tbl.rowCount() == 0:
            QMessageBox.warning(self, "내보내기 불가", "테이블에 데이터가 없습니다.")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "CN값 테이블 내보내기", "cn_value_custom.xlsx",
            "Excel 파일 (*.xlsx)"
        )
        if not path:
            return
        if not path.lower().endswith('.xlsx'):
            path += '.xlsx'
        df = self._get_cn_table_from_widget()
        try:
            df.to_excel(path, index=False)
            QMessageBox.information(self, "내보내기 완료", f"저장 완료:\n{path}")
        except Exception as e:
            QMessageBox.critical(self, "내보내기 오류", str(e))

    def _cn_import(self):
        """사용자가 지정한 xlsx 파일에서 CN값 테이블 불러오기."""
        path, _ = QFileDialog.getOpenFileName(
            self, "CN값 테이블 불러오기", "",
            "Excel 파일 (*.xlsx)"
        )
        if not path:
            return
        try:
            df = load_cn_table(path)
            self._df_to_table(df)
            self._cn_table_loaded = True
            QMessageBox.information(self, "불러오기 완료", f"불러오기 완료:\n{path}")
        except Exception as e:
            QMessageBox.critical(self, "불러오기 오류", str(e))

    def _cn_add_row(self):
        tbl = self.tblCnValues
        row = tbl.rowCount()
        tbl.insertRow(row)
        for c in range(tbl.columnCount()):
            item = QTableWidgetItem("")
            if c > 0:
                item.setTextAlignment(Qt.AlignCenter)
            tbl.setItem(row, c, item)
        tbl.scrollToBottom()
        tbl.setCurrentCell(row, 0)

    def _cn_delete_row(self):
        tbl = self.tblCnValues
        if not tbl.selectedItems():
            QMessageBox.information(self, "행 삭제", "삭제할 행을 먼저 선택하세요.")
            return
        row = tbl.currentRow()
        name_item = tbl.item(row, 0)
        name = name_item.text() if name_item else f"행 {row + 1}"
        reply = QMessageBox.question(
            self, "행 삭제 확인",
            f"'{name}' 항목을 삭제하시겠습니까?",
            QMessageBox.Yes | QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            tbl.removeRow(row)

    def _cn_add_column(self):
        col_name, ok = QInputDialog.getText(
            self, "열 추가", "새 열 이름을 입력하세요\n(예: E, 기타 등):"
        )
        if not ok or not col_name.strip():
            return
        col_name = col_name.strip()
        tbl = self.tblCnValues
        for c in range(tbl.columnCount()):
            h = tbl.horizontalHeaderItem(c)
            if h and h.text() == col_name:
                QMessageBox.warning(self, "중복 열", f"'{col_name}' 열이 이미 존재합니다.")
                return
        col = tbl.columnCount()
        tbl.insertColumn(col)
        tbl.setHorizontalHeaderItem(col, QTableWidgetItem(col_name))
        tbl.horizontalHeader().setSectionResizeMode(col, QHeaderView.ResizeToContents)
        for r in range(tbl.rowCount()):
            item = QTableWidgetItem("")
            item.setTextAlignment(Qt.AlignCenter)
            tbl.setItem(r, col, item)

    def _cn_reload(self):
        reply = QMessageBox.question(
            self, "기본값 불러오기",
            "현재 편집 내용이 사라지고 기본 CN값(cn_value.xlsx)으로 초기화됩니다.\n계속하시겠습니까?",
            QMessageBox.Yes | QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            self._cn_table_loaded = False
            self._load_cn_to_table()

    # ── 재계산 탭 ─────────────────────────────────────────────────────────────

    def _browse_output_dir(self):
        folder = QFileDialog.getExistingDirectory(
            self, "결과 저장 폴더 선택", self.leOutputDir.text() or ""
        )
        if folder:
            self.leOutputDir.setText(folder)

    def _get_recalc_layer(self):
        """cmbRecalcLayer에서 선택된 CN값_input 레이어 반환."""
        layer_id = self.cmbRecalcLayer.currentData()
        if not layer_id:
            raise ValueError("CN값_input 레이어를 선택하세요.")
        layer = QgsProject.instance().mapLayer(layer_id)
        if not layer:
            raise ValueError("선택한 레이어를 찾을 수 없습니다. 레이어 목록을 새로고침하세요.")
        return layer

    def _get_output_dir(self):
        """출력 폴더 경로 반환 (없으면 ValueError)."""
        folder = self.leOutputDir.text().strip()
        if not folder:
            raise ValueError("결과 저장 폴더를 선택하세요.")
        if not os.path.isdir(folder):
            raise ValueError(f"존재하지 않는 폴더입니다: {folder}")
        return folder

    def _export_results(self):
        try:
            layer = self._get_recalc_layer()
            folder = self._get_output_dir()
        except ValueError as e:
            QMessageBox.warning(self, "입력 오류", str(e))
            return

        path = os.path.join(folder, "results.xlsx")
        self._recalc_log("▶ results.xlsx 계산 중...")
        try:
            result1_data, result2_data, null_cn_rows = calculate_results(layer)
            self._warn_null_cn(null_cn_rows)

            # 유역합성 결과
            groups = self._get_watershed_groups()
            grouped_r1, grouped_r2 = None, None
            if groups:
                try:
                    self._recalc_log(f"  유역합성 {len(groups)}개 그룹 계산 중...")
                    grouped_r1, grouped_r2 = calculate_grouped_results(layer, groups)
                    self._recalc_log(f"  유역합성 계산 완료: {len(grouped_r1)}개 그룹")
                except Exception as e:
                    logger.exception("유역합성 계산 오류")
                    self._recalc_log(f"  [경고] 유역합성 계산 실패: {e}")
                    self._recalc_log(f"  → 개별 소유역 결과만 내보냅니다.")
                    grouped_r1, grouped_r2 = None, None

            export_results(result1_data, result2_data, path,
                           grouped_result1=grouped_r1, grouped_result2=grouped_r2)
            self._recalc_log(f"✔ 저장 완료: {path}")
            self._load_xlsx_as_layer(path, "result1", sheet="result1")
            self._load_xlsx_as_layer(path, "result2", sheet="result2")
            QMessageBox.information(self, "저장 완료", f"results.xlsx 저장:\n{path}")
        except Exception as e:
            logger.exception("결과 내보내기 오류")
            self._recalc_log(f"[오류] {e}")
            QMessageBox.critical(self, "오류", str(e))

    # ── 토지이용 재분류 탭 ────────────────────────────────────────────────────

    def _setup_mapping_tab(self):
        """토지이용 재분류 탭을 프로그래매틱으로 생성하여 tabWidget에 삽입."""
        tab = QWidget()
        tab.setStyleSheet("background-color: #f9fafb;")
        outer = QVBoxLayout(tab)
        outer.setSpacing(12)
        outer.setContentsMargins(16, 16, 16, 16)

        # ── 설명 카드 ──────────────────────────────────────────────────
        desc_card = QFrame()
        desc_card.setObjectName("mappingDescCard")
        desc_card.setStyleSheet(
            "QFrame#mappingDescCard {"
            "  background-color: #fef9c3; border: 1px solid #fde047;"
            "  border-radius: 8px;"
            "}"
        )
        desc_layout = QHBoxLayout(desc_card)
        desc_layout.setContentsMargins(14, 10, 14, 10)
        desc_layout.setSpacing(10)

        icon_lbl = QLabel("ℹ")
        icon_lbl.setStyleSheet("font-size: 14px; color: #854d0e; border: none;")
        icon_lbl.setFixedWidth(18)
        desc_layout.addWidget(icon_lbl)

        info_lbl = QLabel(
            "토지피복도 원본 이름 → 재분류 이름으로 매핑합니다. "
            "재분류 이름은 드롭다운에서 선택하거나 직접 입력할 수 있습니다. "
            "오른쪽 CN값 참조 테이블에서 항목을 더블클릭하면 자동 입력됩니다."
        )
        info_lbl.setWordWrap(True)
        info_lbl.setStyleSheet("color: #854d0e; font-size: 12px; border: none;")
        desc_layout.addWidget(info_lbl, 1)
        outer.addWidget(desc_card)

        # ── 매핑 테이블 카드 ───────────────────────────────────────────
        table_card = QFrame()
        table_card.setObjectName("mappingTableCard")
        table_card.setStyleSheet(
            "QFrame#mappingTableCard {"
            "  background-color: white; border: 1px solid #e5e7eb; border-radius: 8px;"
            "}"
        )
        table_layout = QVBoxLayout(table_card)
        table_layout.setSpacing(10)
        table_layout.setContentsMargins(16, 14, 16, 14)

        # 카드 헤더 + 버튼 행
        header_row = QHBoxLayout()
        header_row.setSpacing(8)
        card_title = QLabel("토지이용 재분류 매핑")
        card_title.setStyleSheet(
            "font-size: 14px; font-weight: bold; color: #374151; border: none;"
        )
        header_row.addWidget(card_title)
        header_row.addStretch()

        self.btnMappingLoadLayer = QPushButton("레이어에서 불러오기")
        self.btnMappingAddRow    = QPushButton("+ 행 추가")
        self.btnMappingDeleteRow = QPushButton("- 행 삭제")
        for btn in (self.btnMappingLoadLayer, self.btnMappingAddRow, self.btnMappingDeleteRow):
            header_row.addWidget(btn)
        table_layout.addLayout(header_row)

        # 테이블
        self.tblMapping = QTableWidget(0, 2)
        self.tblMapping.setHorizontalHeaderLabels(['원본 이름', '재분류 이름'])
        self.tblMapping.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.tblMapping.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.tblMapping.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.tblMapping.verticalHeader().setDefaultSectionSize(32)
        self.tblMapping.verticalHeader().setVisible(False)
        self.tblMapping.setItemDelegateForColumn(0, _NextCellDelegate(self.tblMapping))
        self.tblMapping.setItemDelegateForColumn(1, _MappingComboDelegate(self.tblMapping))
        table_layout.addWidget(self.tblMapping)

        # 저장/초기화 버튼 행
        btn_row = QHBoxLayout()
        btn_row.setSpacing(8)
        btn_row.addStretch()
        self.btnMappingClear = QPushButton("초기화")
        self.btnMappingSave  = QPushButton("저장")
        self.btnMappingSave.setStyleSheet(
            "QPushButton { background-color: #1f2937; color: white;"
            "  border: none; border-radius: 6px; padding: 6px 20px; font-weight: 600; }"
            "QPushButton:hover { background-color: #374151; }"
        )
        btn_row.addWidget(self.btnMappingClear)
        btn_row.addWidget(self.btnMappingSave)
        table_layout.addLayout(btn_row)

        # ── CN값 참조 패널 (우측) ────────────────────────────────────────
        ref_card = QFrame()
        ref_card.setObjectName("cnRefCard")
        ref_card.setStyleSheet(
            "QFrame#cnRefCard {"
            "  background-color: white; border: 1px solid #e5e7eb; border-radius: 8px;"
            "}"
        )
        ref_layout = QVBoxLayout(ref_card)
        ref_layout.setSpacing(8)
        ref_layout.setContentsMargins(16, 14, 16, 14)

        ref_title = QLabel("CN값 참조 테이블")
        ref_title.setStyleSheet(
            "font-size: 14px; font-weight: bold; color: #374151; border: none;"
        )
        ref_layout.addWidget(ref_title)

        self.leCnRefSearch = QLineEdit()
        self.leCnRefSearch.setPlaceholderText("검색...")
        ref_layout.addWidget(self.leCnRefSearch)

        self.tblCnRef = QTableWidget(0, 5)
        self.tblCnRef.setHorizontalHeaderLabels(['토지이용분류', 'A', 'B', 'C', 'D'])
        self.tblCnRef.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.tblCnRef.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.tblCnRef.verticalHeader().setDefaultSectionSize(32)
        self.tblCnRef.verticalHeader().setVisible(False)
        ref_layout.addWidget(self.tblCnRef, 1)

        # ── QSplitter (좌: 매핑 테이블, 우: CN참조 패널) ─────────────
        splitter = QSplitter(Qt.Horizontal)
        splitter.addWidget(table_card)
        splitter.addWidget(ref_card)
        splitter.setSizes([550, 450])
        splitter.setStyleSheet("QSplitter::handle { background-color: #e5e7eb; width: 2px; }")
        outer.addWidget(splitter, 1)

        # 다음 단계 버튼
        next_row = QHBoxLayout()
        next_row.addStretch()
        self.btnNextStep1 = QPushButton("다음 단계 →")
        self.btnNextStep1.setMinimumHeight(34)
        self.btnNextStep1.setMinimumWidth(110)
        self.btnNextStep1.setStyleSheet(
            "QPushButton { border: 1px solid #374151; color: #374151;"
            "  border-radius: 6px; padding: 6px 16px; font-weight: 600; background: white; }"
            "QPushButton:hover { background-color: #f3f4f6; }"
        )
        next_row.addWidget(self.btnNextStep1)
        outer.addLayout(next_row)

        self.tabWidget.insertTab(1, tab, "토지이용 재분류")
        self.tabWidget.setTabText(3, "CN값 계산")

    def _mapping_load_saved(self):
        """저장된 land_use_mapping.json을 테이블에 로드."""
        from .core.land_use_mapper import load_mapping
        mapping = load_mapping()
        self.tblMapping.setRowCount(0)
        for original, remapped in mapping.items():
            r = self.tblMapping.rowCount()
            self.tblMapping.insertRow(r)
            self.tblMapping.setItem(r, 0, QTableWidgetItem(original))
            self.tblMapping.setItem(r, 1, QTableWidgetItem(remapped))
        self._mapping_tab_loaded = True

    def _mapping_load_from_layer(self):
        """캔버스의 토양군_토지피복_교차 레이어에서 선택된 분류 수준의 고유 토지이용 값을 채운다."""
        intersect_lyrs = [
            l for l in QgsProject.instance().mapLayers().values()
            if (l.type() == QgsMapLayerType.VectorLayer and l.name() == "토양군_토지피복_교차")
        ]
        if not intersect_lyrs:
            QMessageBox.information(
                self, "레이어 없음",
                "캔버스에 토양군_토지피복_교차 레이어가 없습니다.\n"
                "먼저 [레이어 불러오기] 탭에서 실행하세요."
            )
            return

        level = self._get_level()
        col_name = {'l1': 'l1_name', 'l2': 'l2_name', 'l3': 'l3_name'}[level]

        layer = intersect_lyrs[0]
        field_names = [f.name() for f in layer.fields()]
        if col_name not in field_names:
            QMessageBox.warning(
                self, "칼럼 없음",
                f"레이어에 '{col_name}' 칼럼이 없습니다.\n"
                f"분류 수준({level})에 맞는 레이어인지 확인하세요."
            )
            return

        # 기존 재분류 이름 보존
        existing = self._mapping_table_to_dict()

        unique_values = sorted({
            str(f[col_name]) for f in layer.getFeatures()
            if f[col_name] and str(f[col_name]) not in ('NULL', 'None', '')
        })

        self.tblMapping.setRowCount(0)
        for val in unique_values:
            r = self.tblMapping.rowCount()
            self.tblMapping.insertRow(r)
            self.tblMapping.setItem(r, 0, QTableWidgetItem(val))
            self.tblMapping.setItem(r, 1, QTableWidgetItem(existing.get(val, '')))

    def _mapping_table_to_dict(self) -> dict:
        """현재 테이블을 {원본: 재분류} dict로 변환. 재분류 이름이 빈 행은 제외."""
        mapping = {}
        for r in range(self.tblMapping.rowCount()):
            orig  = (self.tblMapping.item(r, 0) or QTableWidgetItem()).text().strip()
            remap = (self.tblMapping.item(r, 1) or QTableWidgetItem()).text().strip()
            if orig and remap:
                mapping[orig] = remap
        return mapping

    def _mapping_add_row(self):
        r = self.tblMapping.rowCount()
        self.tblMapping.insertRow(r)
        self.tblMapping.setItem(r, 0, QTableWidgetItem(''))
        self.tblMapping.setItem(r, 1, QTableWidgetItem(''))
        self.tblMapping.setCurrentCell(r, 0)

    def _mapping_delete_row(self):
        r = self.tblMapping.currentRow()
        if r < 0:
            QMessageBox.information(self, "행 삭제", "삭제할 행을 먼저 선택하세요.")
            return
        self.tblMapping.removeRow(r)

    def _mapping_save(self):
        from .core.land_use_mapper import save_mapping
        mapping = self._mapping_table_to_dict()
        try:
            save_mapping(mapping)
            QMessageBox.information(
                self, "저장 완료",
                f"토지이용 재분류 매핑이 저장되었습니다.\n적용 항목: {len(mapping)}개"
            )
        except Exception as e:
            QMessageBox.critical(self, "저장 오류", str(e))

    def _load_cn_ref_table(self):
        """CN값 참조 테이블 로드 (Tab 1 우측 패널)."""
        try:
            from .core.cn_matcher import load_cn_table, XLSX_PATH
            df = load_cn_table(XLSX_PATH)
        except Exception:
            # 위젯 테이블에서 폴백
            df = self._get_cn_table_from_widget()
            if df is None or df.empty:
                return
        tbl = self.tblCnRef
        tbl.blockSignals(True)
        cols = list(df.columns)
        tbl.setColumnCount(len(cols))
        tbl.setHorizontalHeaderLabels(cols)
        tbl.setRowCount(len(df))
        for r, (_, row) in enumerate(df.iterrows()):
            for c, val in enumerate(row):
                text = "" if pd.isna(val) else str(val)
                item = QTableWidgetItem(text)
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                tbl.setItem(r, c, item)
        tbl.blockSignals(False)
        # Auto-resize columns
        tbl.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        for c in range(1, tbl.columnCount()):
            tbl.horizontalHeader().setSectionResizeMode(c, QHeaderView.ResizeToContents)
        self._cn_ref_loaded = True

    def _filter_cn_ref_table(self, text):
        """검색어로 CN표 참조 테이블 행 필터링."""
        tbl = self.tblCnRef
        search = text.strip().lower()
        for r in range(tbl.rowCount()):
            item = tbl.item(r, 0)
            name = item.text().lower() if item else ''
            tbl.setRowHidden(r, search != '' and search not in name)

    def _sync_cn_ref_from_edit(self):
        """Tab 2에서 편집된 CN표를 참조 패널에 동기화."""
        df = self._get_cn_table_from_widget()
        if df is not None and not df.empty:
            tbl = self.tblCnRef
            tbl.blockSignals(True)
            cols = list(df.columns)
            tbl.setColumnCount(len(cols))
            tbl.setHorizontalHeaderLabels(cols)
            tbl.setRowCount(len(df))
            for r, (_, row) in enumerate(df.iterrows()):
                for c, val in enumerate(row):
                    text = "" if pd.isna(val) else str(val)
                    item = QTableWidgetItem(text)
                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                    tbl.setItem(r, c, item)
            tbl.blockSignals(False)
        self._cn_ref_dirty = False
        # Re-apply search filter if active
        if hasattr(self, 'leCnRefSearch'):
            self._filter_cn_ref_table(self.leCnRefSearch.text())

    def _mapping_clear(self):
        reply = QMessageBox.question(
            self, "초기화 확인",
            "모든 매핑을 삭제하고 저장하시겠습니까?",
            QMessageBox.Yes | QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            self.tblMapping.setRowCount(0)
            from .core.land_use_mapper import save_mapping
            save_mapping({})

    def _validate_mapping_cell(self, row, col):
        """재분류명 유효성 → 셀 배경색 표시."""
        if col != 1:
            return
        item = self.tblMapping.item(row, col)
        if not item:
            return
        text = item.text().strip()
        if not text:
            item.setBackground(Qt.transparent)
            item.setToolTip('')
            return
        cn_names = self._get_cn_land_use_names()
        if text in cn_names:
            from qgis.PyQt.QtGui import QColor
            item.setBackground(QColor('#f0fdf4'))
            item.setToolTip('')
        else:
            from qgis.PyQt.QtGui import QColor
            item.setBackground(QColor('#fef2f2'))
            item.setToolTip(f'CN값 테이블에 "{text}" 항목이 없습니다')

    def _on_cn_ref_double_click(self, index):
        """CN표 참조 패널 더블클릭 → 매핑 테이블 재분류명 자동 입력."""
        # 0번 열(토지이용분류)의 텍스트 가져오기
        item = self.tblCnRef.item(index.row(), 0)
        if not item:
            return
        lu_name = item.text().strip()
        if not lu_name:
            return
        # 매핑 테이블의 현재 선택 행
        cur_row = self.tblMapping.currentRow()
        if cur_row < 0:
            return
        # 재분류 이름 열(col 1)에 설정
        target_item = self.tblMapping.item(cur_row, 1)
        if target_item is None:
            target_item = QTableWidgetItem('')
            self.tblMapping.setItem(cur_row, 1, target_item)
        target_item.setText(lu_name)
        # 다음 행으로 자동 이동
        next_row = cur_row + 1
        if next_row < self.tblMapping.rowCount():
            self.tblMapping.setCurrentCell(next_row, 1)

    def _enhance_recalc_tab(self):
        """CN값 계산 탭 상단에 'CN값 계산' 카드를 동적으로 추가."""
        recalc_tab = self.tabWidget.widget(TAB_RECALC)
        outer_layout = recalc_tab.layout()
        if outer_layout is None:
            return

        card = QFrame()
        card.setObjectName("applyCard")
        card.setStyleSheet(
            "QFrame#applyCard {"
            "  background-color: white; border: 1px solid #e5e7eb; border-radius: 8px;"
            "}"
        )
        card_layout = QVBoxLayout(card)
        card_layout.setSpacing(10)
        card_layout.setContentsMargins(16, 14, 16, 14)

        # 카드 헤더
        card_title = QLabel("CN값 계산")
        card_title.setStyleSheet(
            "font-size: 14px; font-weight: bold; color: #374151; border: none;"
        )
        card_layout.addWidget(card_title)

        # 설명
        lbl = QLabel(
            "[레이어 불러오기] 탭에서 불러온 교차 레이어를 기반으로 토지이용 재분류 매핑과 "
            "CN값을 적용하여 CN값_input 레이어를 생성합니다."
        )
        lbl.setWordWrap(True)
        lbl.setStyleSheet("color: #6b7280; font-size: 12px; border: none;")
        card_layout.addWidget(lbl)

        # 버튼 행
        btn_row = QHBoxLayout()
        btn_row.addStretch()
        self.btnApplyCn = QPushButton("CN값 계산 실행")
        self.btnApplyCn.setMinimumHeight(34)
        self.btnApplyCn.setStyleSheet(
            "QPushButton {"
            "  background-color: #1f2937; color: white;"
            "  border: none; border-radius: 6px;"
            "  padding: 7px 24px; font-weight: 600;"
            "}"
            "QPushButton:hover { background-color: #374151; }"
            "QPushButton:disabled { background-color: #9ca3af; }"
        )
        btn_row.addWidget(self.btnApplyCn)
        card_layout.addLayout(btn_row)

        outer_layout.insertWidget(0, card)

        self._setup_watershed_group_card(outer_layout)

    def _setup_watershed_group_card(self, recalc_layout):
        """Tab 3에 유역합성 설정 카드를 동적으로 추가."""
        card = QFrame()
        card.setObjectName("watershedGroupCard")
        card.setStyleSheet(
            "QFrame#watershedGroupCard {"
            "  background-color: white; border: 1px solid #e5e7eb; border-radius: 8px;"
            "}"
        )
        card_layout = QVBoxLayout(card)
        card_layout.setSpacing(10)
        card_layout.setContentsMargins(16, 14, 16, 14)

        # 카드 헤더
        header_row = QHBoxLayout()
        card_title = QLabel("유역합성 설정")
        card_title.setStyleSheet(
            "font-size: 14px; font-weight: bold; color: #374151; border: none;"
        )
        header_row.addWidget(card_title)
        header_row.addStretch()

        self.btnWsGroupAddRow = QPushButton("+ 행 추가")
        self.btnWsGroupDeleteRow = QPushButton("- 행 삭제")
        header_row.addWidget(self.btnWsGroupAddRow)
        header_row.addWidget(self.btnWsGroupDeleteRow)
        card_layout.addLayout(header_row)

        # 설명
        desc = QLabel(
            "소유역을 그룹화하여 합성 CN값을 계산합니다. "
            "그룹명을 입력하고, 해당 그룹에 포함될 소유역을 선택하세요. (최대 20개)"
        )
        desc.setWordWrap(True)
        desc.setStyleSheet("color: #6b7280; font-size: 12px; border: none;")
        card_layout.addWidget(desc)

        # 테이블 (그룹명 + 소유역1~20)
        self._ws_max_members = 20
        self.tblWsGroups = QTableWidget(0, 1 + self._ws_max_members)
        headers = ['그룹명'] + [f'소유역{i}' for i in range(1, self._ws_max_members + 1)]
        self.tblWsGroups.setHorizontalHeaderLabels(headers)
        self.tblWsGroups.horizontalHeader().setSectionResizeMode(0, QHeaderView.Interactive)
        self.tblWsGroups.setColumnWidth(0, 100)
        for c in range(1, 1 + self._ws_max_members):
            self.tblWsGroups.horizontalHeader().setSectionResizeMode(c, QHeaderView.Interactive)
            self.tblWsGroups.setColumnWidth(c, 90)
        self.tblWsGroups.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.tblWsGroups.verticalHeader().setDefaultSectionSize(32)
        self.tblWsGroups.verticalHeader().setVisible(False)
        card_layout.addWidget(self.tblWsGroups)

        # 버튼 행
        btn_row = QHBoxLayout()
        btn_row.setSpacing(8)
        btn_row.addStretch()

        self.btnWsGroupLoad = QPushButton("불러오기")
        self.btnWsGroupSave = QPushButton("저장")
        self.btnWsGroupSave.setStyleSheet(
            "QPushButton { background-color: #1f2937; color: white;"
            "  border: none; border-radius: 6px; padding: 6px 20px; font-weight: 600; }"
            "QPushButton:hover { background-color: #374151; }"
        )
        btn_row.addWidget(self.btnWsGroupLoad)
        btn_row.addWidget(self.btnWsGroupSave)
        card_layout.addLayout(btn_row)

        # Insert after CN calculation card (position 1)
        recalc_layout.insertWidget(1, card)

        # Connect signals
        self.btnWsGroupAddRow.clicked.connect(self._ws_group_add_row)
        self.btnWsGroupDeleteRow.clicked.connect(self._ws_group_delete_row)
        self.btnWsGroupSave.clicked.connect(self._ws_group_save)
        self.btnWsGroupLoad.clicked.connect(self._ws_group_load)

    def _get_watershed_names(self) -> list:
        """소유역명 고유값 목록을 반환. 입력 레이어 → CN값_input 폴백."""
        _INVALID = {'', 'null', 'none', 'nan'}

        def _valid_name(val):
            if val is None:
                return None
            try:
                from qgis.PyQt.QtCore import QVariant
                if isinstance(val, QVariant) and val.isNull():
                    return None
            except Exception:
                pass
            s = str(val).strip()
            return s if s and s.lower() not in _INVALID else None

        # 1) 입력 레이어에서 조회
        if hasattr(self, 'input_layer') and self.input_layer is not None and self._last_name_field:
            try:
                names = sorted({
                    n for f in self.input_layer.getFeatures()
                    if (n := _valid_name(f[self._last_name_field]))
                })
                if names:
                    return names
            except Exception as e:
                logger.debug(f"입력 레이어에서 소유역명 조회 실패: {e}")

        # 2) CN값_input 레이어에서 폴백 조회
        for lyr in QgsProject.instance().mapLayers().values():
            try:
                if lyr.type() == QgsMapLayerType.VectorLayer and lyr.name() == "CN값_input":
                    names = sorted({
                        n for f in lyr.getFeatures()
                        if (n := _valid_name(f['소유역명']))
                    })
                    if names:
                        return names
            except Exception as e:
                logger.debug(f"CN값_input 레이어에서 소유역명 조회 실패: {e}")

        return []

    def _ws_create_combo(self, ws_names: list, selected: str = '') -> 'QComboBox':
        """소유역 선택용 콤보박스 생성."""
        from qgis.PyQt.QtWidgets import QComboBox
        cmb = QComboBox()
        cmb.setEditable(True)
        cmb.addItem('')  # 빈 항목 (선택 안 함)
        cmb.addItems(ws_names)
        if selected:
            idx = cmb.findText(selected)
            if idx >= 0:
                cmb.setCurrentIndex(idx)
            else:
                cmb.setCurrentText(selected)
        cmb.setStyleSheet(
            "QComboBox { border: none; padding: 2px 4px; background: transparent; }"
        )
        return cmb

    def _ws_group_add_row(self):
        try:
            ws_names = self._get_watershed_names()
        except Exception as e:
            logger.exception("소유역명 목록 조회 오류")
            ws_names = []

        if not ws_names:
            reply = QMessageBox.question(
                self, "소유역 데이터 없음",
                "소유역 목록을 불러올 수 없습니다.\n"
                "[레이어 불러오기] 탭에서 먼저 실행하거나,\n"
                "CN값 계산을 먼저 수행하세요.\n\n"
                "소유역명을 직접 입력하시겠습니까?",
                QMessageBox.Yes | QMessageBox.No
            )
            if reply != QMessageBox.Yes:
                return
        tbl = self.tblWsGroups
        row = tbl.rowCount()
        tbl.insertRow(row)
        # 그룹명은 텍스트 입력
        tbl.setItem(row, 0, QTableWidgetItem(''))
        # 소유역 열은 콤보박스
        for c in range(1, tbl.columnCount()):
            cmb = self._ws_create_combo(ws_names)
            tbl.setCellWidget(row, c, cmb)
        tbl.scrollToBottom()
        tbl.setCurrentCell(row, 0)

    def _ws_group_delete_row(self):
        tbl = self.tblWsGroups
        row = tbl.currentRow()
        if row < 0:
            QMessageBox.information(self, "행 삭제", "삭제할 행을 먼저 선택하세요.")
            return
        tbl.removeRow(row)

    def _ws_group_save(self):
        """유역합성 테이블 → JSON 저장."""
        groups = self._get_watershed_groups()
        try:
            save_groups(groups)
            QMessageBox.information(
                self, "저장 완료",
                f"유역합성 그룹이 저장되었습니다.\n그룹 수: {len(groups)}개"
            )
        except Exception as e:
            QMessageBox.critical(self, "저장 오류", str(e))

    def _ws_group_load(self):
        """저장된 유역합성 그룹을 테이블에 로드."""
        groups = load_groups()
        if not groups:
            QMessageBox.information(self, "불러오기", "저장된 유역합성 그룹이 없습니다.")
            return
        try:
            ws_names = self._get_watershed_names()
        except Exception as e:
            logger.exception("소유역명 목록 조회 오류 (불러오기)")
            ws_names = []
        tbl = self.tblWsGroups
        tbl.setRowCount(0)
        for group_name, members in groups.items():
            row = tbl.rowCount()
            tbl.insertRow(row)
            tbl.setItem(row, 0, QTableWidgetItem(group_name))
            for i in range(self._ws_max_members):
                selected = members[i] if i < len(members) else ''
                cmb = self._ws_create_combo(ws_names, selected)
                tbl.setCellWidget(row, i + 1, cmb)
        QMessageBox.information(
            self, "불러오기 완료",
            f"유역합성 그룹 {len(groups)}개를 불러왔습니다."
        )

    def _get_watershed_groups(self) -> dict:
        """유역합성 테이블에서 그룹 데이터 추출 (콤보박스에서 값 읽기)."""
        tbl = self.tblWsGroups
        groups = {}
        for r in range(tbl.rowCount()):
            name_item = tbl.item(r, 0)
            group_name = name_item.text().strip() if name_item else ''
            if not group_name:
                continue
            members = []
            for c in range(1, tbl.columnCount()):
                widget = tbl.cellWidget(r, c)
                if widget:
                    val = widget.currentText().strip()
                else:
                    item = tbl.item(r, c)
                    val = item.text().strip() if item else ''
                if val:
                    members.append(val)
            if members:
                groups[group_name] = members
        return groups

    def _apply_cn_calc(self):
        """교차 레이어에 토지이용 매핑·CN값을 적용하여 CN값_input 레이어를 최초 생성."""
        if self._final_intersect_layer is None:
            QMessageBox.warning(
                self, "레이어 없음",
                "먼저 [레이어 불러오기] 탭에서 실행하세요."
            )
            return

        self._recalc_log("▶ CN값 계산 중...")
        try:
            # ⑤ CN값_input 레이어 구성
            self._recalc_log("⑤ CN값_input 레이어 구성 중...")
            _, name_col = LEVEL_COLUMNS[self._last_level]
            cn_input_layer = _build_result_layer(
                self._final_intersect_layer, self._last_name_field, name_col
            )
            self._recalc_log(f"   → {cn_input_layer.featureCount()} features 생성")

            # ⑥ 토지이용 재분류 매핑 적용
            from .core.land_use_mapper import load_mapping, apply_mapping_to_layer
            mapping = load_mapping()
            if mapping:
                changed = apply_mapping_to_layer(cn_input_layer, mapping)
                self._recalc_log(f"⑥ 토지이용 재분류 적용: {changed}건")
            else:
                self._recalc_log("⑥ (저장된 재분류 매핑 없음)")

            # ⑦ CN값 매칭 (편집 탭에 로드된 테이블 우선, 없으면 기본 파일)
            self._recalc_log("⑦ CN값 매칭 중...")
            if not self._cn_table_loaded:
                self._load_cn_to_table()
            cn_table = self._get_cn_table_from_widget()
            if cn_table.empty:
                cn_table = load_cn_table()
            fail_list = apply_cn_to_layer(cn_input_layer, cn_table, self._last_level)
            if fail_list:
                self._recalc_log(f"   ⚠ CN 매칭 실패: {len(fail_list)}건")
                for wname, luse, htype in fail_list[:10]:
                    self._recalc_log(f"      - 소유역={wname}, 토지이용={luse}, 토양군={htype}")
                if len(fail_list) > 10:
                    self._recalc_log(f"      ... 외 {len(fail_list) - 10}건")
            else:
                self._recalc_log("   ✔ 모든 피처 CN값 매칭 완료")

            cn_input_layer.setName("CN값_input")
            QgsProject.instance().addMapLayer(cn_input_layer)
            self._recalc_log("   ✦ 레이어 추가됨: [CN값_input]")
            self._refresh_recalc_layer_list()

            self._recalc_log("━" * 44)
            self._recalc_log("✔ CN값 계산 완료!")
        except Exception as e:
            logger.exception("CN값 계산 오류")
            self._recalc_log(f"[오류] {e}")
            QMessageBox.critical(self, "오류", str(e))

    def _load_xlsx_as_layer(self, path: str, name: str, sheet: str = "Sheet1"):
        """xlsx 파일의 특정 시트를 QGIS 테이블 레이어로 프로젝트에 추가."""
        uri = f"{path}|layername={sheet}"
        layer = QgsVectorLayer(uri, name, "ogr")
        if layer.isValid():
            QgsProject.instance().addMapLayer(layer)
            self._recalc_log(f"   ✦ 레이어로 표시됨: [{name}] (시트: {sheet})")
        else:
            self._recalc_log(f"   ⚠ 레이어 로드 실패 (파일은 저장됨): {path}")

    def _warn_null_cn(self, null_cn_rows: list):
        """CN값 NULL 피처가 있을 경우 경고 다이얼로그 표시."""
        if not null_cn_rows:
            return
        lines = [f"CN값 NULL로 계산에서 제외된 피처 {len(null_cn_rows)}건:\n"]
        for ws, lu, ht in null_cn_rows[:20]:
            lines.append(f"  • 소유역={ws}, 토지이용={lu}, 토양군={ht}")
        if len(null_cn_rows) > 20:
            lines.append(f"  ... 외 {len(null_cn_rows) - 20}건")
        lines.append("\n해당 피처의 CN값을 [CN값 편집] 탭에서 확인·수정하세요.")
        QMessageBox.warning(self, "CN값 NULL 피처 있음", "\n".join(lines))

    def _recalc_log(self, msg: str):
        self.txtRecalcLog.append(msg)
        sb = self.txtRecalcLog.verticalScrollBar()
        sb.setValue(sb.maximum())

    # ── 유틸리티 ──────────────────────────────────────────────────────────────

    def _log(self, msg: str):
        self.txtLog.append(msg)
        sb = self.txtLog.verticalScrollBar()
        sb.setValue(sb.maximum())
