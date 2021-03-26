import sys
from GUI import main as ui
from PyQt5.QtWidgets import QApplication
from PyQt5 import QtWidgets

app = QApplication(sys.argv)
mainwindow = ui.MainWindow()
widget = QtWidgets.QStackedWidget()
widget.addWidget(mainwindow)
widget.setFixedSize(510,930)
widget.show()
sys.exit(app.exec_())
