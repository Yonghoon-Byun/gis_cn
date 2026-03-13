def classFactory(iface):
    from .plugin import CnCalculatorPlugin
    return CnCalculatorPlugin(iface)
