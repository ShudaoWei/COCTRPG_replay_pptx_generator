import sys
from GUI import main as ui
from PyQt5.QtWidgets import QApplication
from PyQt5 import QtWidgets

app = QApplication(sys.argv)
mainwindow = ui.MainWindow()
mainwindow.show()
sys.exit(app.exec_())
