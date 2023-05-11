from PyQt5 import QtCore, QtGui, QtWidgets
import sys, os, pptx_maker

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.setFixedSize(230, 100)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.horizontalLayout = QtWidgets.QHBoxLayout(self.centralwidget)
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.MainLayout = QtWidgets.QVBoxLayout()
        self.MainLayout.setObjectName("MainLayout")
        self.SaveButton = QtWidgets.QPushButton(self.centralwidget)
        self.SaveButton.setObjectName("SaveButton")
        self.SaveButton.clicked.connect(self.open)
        self.MainLayout.addWidget(self.SaveButton)
        self.StartButton = QtWidgets.QPushButton(self.centralwidget)
        self.StartButton.setObjectName("StartButton")
        self.StartButton.setEnabled(False)
        self.StartButton.clicked.connect(self.start)
        self.MainLayout.addWidget(self.StartButton)
        self.horizontalLayout.addLayout(self.MainLayout)
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 240, 21))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Pptx Maker"))
        self.SaveButton.setText(_translate("MainWindow", "Emplacement de sauvegarde"))
        self.StartButton.setText(_translate("MainWindow", "Lancer la création de pptx"))

    def open(self):
        self.file, ok = QtWidgets. QFileDialog.getSaveFileName(MainWindow, 'Sauvegarder', os.getenv('HOME'), 'PPTX(*.pptx)')
        if ok:
            self.StartButton.setEnabled(True)
        else:
            self.StartButton.setEnabled(False)
    
    def start(self):
        valid = pptx_maker.aqui(self.file)

        if not valid:
            msg = QtWidgets.QMessageBox()
            
            msg.setIcon(QtWidgets.QMessageBox.Critical)
            
            msg.setText('Echec')

            msg.setWindowTitle("")
            msg.exec_()

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)

    QtCore.QDir.addSearchPath('Assets', os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Assets'))
    app.setWindowIcon(QtGui.QIcon("Assets:Logo_small.png"))
    file = QtCore.QFile('Assets:Style.qss')
    file.open(QtCore.QFile.ReadOnly | QtCore.QFile.Text)
    app.setStyleSheet(str(file.readAll(), 'utf-8'))

    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())
