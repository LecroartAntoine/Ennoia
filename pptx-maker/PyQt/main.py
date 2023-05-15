from PyQt5 import QtCore, QtGui, QtWidgets
import sys, os, pptx_maker

class MyBar(QtWidgets.QWidget):

    def __init__(self, parent):
        super(MyBar, self).__init__()
        self.parent = parent
        self.Barlayout = QtWidgets.QHBoxLayout()
        self.Barlayout.setContentsMargins(0,0,0,0)
        
        self.title = QtWidgets.QLabel("CMF PPTX")
        
        btn_size = 40

        self.btn_close = QtWidgets.QPushButton("")
        self.btn_close.setIcon(QtGui.QIcon('Assets:close.png'))
        self.btn_close.setIconSize(QtCore.QSize(25, 25))
        self.btn_close.clicked.connect(self.btn_close_clicked)
        self.btn_close.setFixedSize(btn_size,btn_size)
        self.btn_close.setStyleSheet("""
            QPushButton 
            {
                border-width: 5px; 
                border-color: #242526;
            }
            QPushButton::hover
            {
                background-color: #8399ff;
                border-width: 3px;
                border-color: #242526;
            }
            QPushButton::pressed
            {
                background-color: #4969ff;
                border-color: #242526;
            }
        """)

        self.btn_min = QtWidgets.QPushButton("")
        self.btn_min.setIcon(QtGui.QIcon('Assets:min.png'))
        self.btn_min.setIconSize(QtCore.QSize(25, 25))
        self.btn_min.clicked.connect(self.btn_min_clicked)
        self.btn_min.setFixedSize(btn_size, btn_size)
        self.btn_min.setStyleSheet("""
            QPushButton 
            {
                border-width: 5px; 
                border-color: #242526;
            }
            QPushButton::hover
            {
                background-color: #8399ff;
                border-width: 3px;
                border-color: #242526;
            }
            QPushButton::pressed
            {
                background-color: #4969ff;
                border-color: #242526;
            }
        """)

        self.title.setFixedHeight(40)
        self.title.setAlignment(QtCore.Qt.AlignCenter)
        self.Barlayout.addWidget(self.title)
        self.Barlayout.addWidget(self.btn_min)
        self.Barlayout.addWidget(self.btn_close)

        self.title.setStyleSheet("""
            background-color: #242526;
            color: white;
            font-weight: bold;
            font-family : Trebuchet MS;
            font-size : 28px;
        """)

        self.setLayout(self.Barlayout)
        self.start = QtCore.QPoint(0, 0)
        self.pressing = False
        self.title.mouseMoveEvent = self.mouseMoveEvent
        self.title.mousePressEvent = self.mousePressEvent
        self.title.mouseReleaseEvent = self.mouseReleaseEvent

        self.icon = QtWidgets.QLabel(self.parent)
        self.icon.setPixmap(QtGui.QPixmap("Assets:Logo2.png").scaledToHeight(30, QtCore.Qt.SmoothTransformation))
        self.icon.setContentsMargins(2, 5, 0 , 2)
        self.icon.setGeometry(QtCore.QRect(0, 0, 102, 37))
        self.icon.setStyleSheet("background-color: #242526;")
        self.icon.mouseMoveEvent = self.mouseMoveEvent
        self.icon.mousePressEvent = self.mousePressEvent
        self.icon.mouseReleaseEvent = self.mouseReleaseEvent

    def resizeEvent(self, QResizeEvent):
        super(MyBar, self).resizeEvent(QResizeEvent)
        self.title.resize(self.parent.width(), self.parent.height())

    def mousePressEvent(self, event):
        if event.buttons() == QtCore.Qt.LeftButton:
            self.start = self.mapToGlobal(event.pos())
            self.pressing = True
            
    def mouseMoveEvent(self, event):
        if self.pressing:
            self.end = self.mapToGlobal(event.pos())
            self.movement = self.end-self.start
            self.parent.setGeometry(self.mapToGlobal(self.movement).x(),
                                self.mapToGlobal(self.movement).y(),
                                self.parent.width(),
                                self.parent.height())
            self.start = self.end

    def mouseReleaseEvent(self, QMouseEvent):
        self.pressing = False
        if self.parent.pos().y() < 0:
            self.parent.move(self.parent.pos().x(), 0)
        
    def btn_min_clicked(self):
        self.parent.showMinimized()

    def btn_close_clicked(self):
        self.parent.close()

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")

        self.MainLayout = QtWidgets.QVBoxLayout(self.centralwidget)
        self.MainLayout.setObjectName("MainLayout")

        self.MainLayout.addWidget(MyBar(MainWindow))

        self.MainLayout.setContentsMargins(0,0,0,0)
        MainWindow.setFixedSize(400, 135)
        MainWindow.setWindowFlags(QtCore.Qt.CustomizeWindowHint | QtCore.Qt.FramelessWindowHint)

        self.SaveButton = QtWidgets.QPushButton(self.centralwidget)
        self.SaveButton.setObjectName("SaveButton")
        self.SaveButton.clicked.connect(self.open)
        self.MainLayout.addWidget(self.SaveButton)
        self.StartButton = QtWidgets.QPushButton(self.centralwidget)
        self.StartButton.setObjectName("StartButton")
        self.StartButton.setEnabled(False)
        self.StartButton.clicked.connect(self.start)
        self.MainLayout.addWidget(self.StartButton)
       
        MainWindow.setCentralWidget(self.centralwidget)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "CMF PPTX"))
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
            
            msg.setIcon(QtWidgets.QMessageBox.Warning)

            msg.setWindowTitle(" ")
            
            msg.setText('Échec')

            msg.exec_()

        else :
            msg = QtWidgets.QMessageBox()
            
            msg.setWindowIcon(QtWidgets.QMessageBox.Information)

            msg.setWindowTitle(" ")

            msg.setText('Succès')

            msg.exec_()

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)

    QtCore.QDir.addSearchPath('Assets', os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Assets'))
    app.setWindowIcon(QtGui.QIcon("Assets:Logo.ico"))
    file = QtCore.QFile('Assets:Style.qss')
    file.open(QtCore.QFile.ReadOnly | QtCore.QFile.Text)
    app.setStyleSheet(str(file.readAll(), 'utf-8'))

    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())
