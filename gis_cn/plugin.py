import os
from qgis.PyQt.QtWidgets import QAction
from qgis.PyQt.QtGui import QIcon


class CnCalculatorPlugin:
    def __init__(self, iface):
        self.iface = iface
        self.action = None
        self.dialog = None

    def initGui(self):
        icon = QIcon(os.path.join(os.path.dirname(__file__), 'cn.png'))
        self.action = QAction(icon, "CN값 계산기", self.iface.mainWindow())
        self.action.triggered.connect(self.run)
        self.iface.addToolBarIcon(self.action)
        self.iface.addPluginToMenu("CN값 계산기", self.action)

    def unload(self):
        self.iface.removePluginMenu("CN값 계산기", self.action)
        self.iface.removeToolBarIcon(self.action)
        del self.action

    def run(self):
        from .dialog import CnCalculatorDialog
        if self.dialog is None:
            self.dialog = CnCalculatorDialog(self.iface)
        self.dialog._refresh_layer_list()
        self.dialog.show()
        self.dialog.raise_()
        self.dialog.activateWindow()
