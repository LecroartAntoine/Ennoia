from PyQt5 import QtCore, QtGui, QtWidgets
import sys, os
import pandas as pd
   
class PandasModel(QtCore.QAbstractTableModel):
    def __init__(self, dataframe: pd.DataFrame, parent=None):
        QtCore.QAbstractTableModel.__init__(self, parent)
        self._dataframe = dataframe

    def rowCount(self, parent=QtCore.QModelIndex()) -> int:
        if parent == QtCore.QModelIndex():
            return len(self._dataframe)

        return 0

    def columnCount(self, parent=QtCore.QModelIndex()) -> int:
        if parent == QtCore.QModelIndex():
            return len(self._dataframe.columns)
        return 0

    def data(self, index: QtCore.QModelIndex, role=QtCore.Qt.ItemDataRole):
        if not index.isValid():
            return None

        if role == QtCore.Qt.DisplayRole:
            try :
                x = round(self._dataframe.iloc[index.row()][index.column()], 2)
                return str(x)
            except :
                return str(self._dataframe.iloc[index.row()][index.column()])
            
        if role == QtCore.Qt.BackgroundRole and self._dataframe.iloc[index.row()][index.column()] == 'x':
            return QtGui.QColor('red')

        return None

    def headerData(
        self, section: int, orientation: QtCore.Qt.Orientation, role: QtCore.Qt.ItemDataRole
    ):

        if role == QtCore.Qt.DisplayRole:
            if orientation == QtCore.Qt.Horizontal:
                return str(self._dataframe.columns[section])

            if orientation == QtCore.Qt.Vertical:
                return str(self._dataframe.index[section])
        return None
    
class MyBar(QtWidgets.QWidget):

    def __init__(self, parent):
        super(MyBar, self).__init__()
        self.parent = parent
        self.layout = QtWidgets.QHBoxLayout()
        self.layout.setContentsMargins(0,20,0,5)
        self.pressing = False

        self.title = QtWidgets.QLabel("""
                                        <div>
                                        <img style="vertical-align:middle;text-align:center;float:left;" src="Assets:Logo_small.png" width="40" height="40">
                                        <span>Déplacement céphalométrique</span>
                                        </div>
                                    """)

        btn_size = 50

        self.btn_close = QtWidgets.QPushButton("")
        self.btn_close.setIcon(QtGui.QIcon('Assets:close.png'))
        self.btn_close.setIconSize(QtCore.QSize(35, 35))
        self.btn_close.clicked.connect(self.btn_close_clicked)
        self.btn_close.setFixedSize(btn_size,btn_size)
        self.btn_close.setStyleSheet("""
            QPushButton 
            {
                border-width: 5px; 
                border-color: #404040;
            }
            QPushButton::hover
            {
                background-color: #8399ff;
                border-width: 3px;
                border-color: #404040;
            }
            QPushButton::pressed
            {
                background-color: #4969ff;
                border-color: #404040;
            }
        """)

        self.btn_min = QtWidgets.QPushButton("")
        self.btn_min.setIcon(QtGui.QIcon('Assets:min.png'))
        self.btn_min.setIconSize(QtCore.QSize(35, 35))
        self.btn_min.clicked.connect(self.btn_min_clicked)
        self.btn_min.setFixedSize(btn_size, btn_size)
        self.btn_min.setStyleSheet("""
            QPushButton 
            {
                border-width: 5px; 
                border-color: #404040;
            }
            QPushButton::hover
            {
                background-color: #8399ff;
                border-width: 3px;
                border-color: #404040;
            }
            QPushButton::pressed
            {
                background-color: #4969ff;
                border-color: #404040;
            }
        """)
        

        self.btn_max = QtWidgets.QPushButton("")
        self.btn_max.setIcon(QtGui.QIcon('Assets:max.png'))
        self.btn_max.setIconSize(QtCore.QSize(35, 35))
        self.btn_max.clicked.connect(self.btn_max_clicked)
        self.btn_max.setFixedSize(btn_size, btn_size)
        self.btn_max.setStyleSheet("""
            QPushButton 
            {
                border-width: 5px; 
                border-color: #404040;
            }
            QPushButton::hover
            {
                background-color: #8399ff;
                border-width: 3px;
                border-color: #404040;
            }
            QPushButton::pressed
            {
                background-color: #4969ff;
                border-color: #404040;
            }
        """)

        self.title.setFixedHeight(50)
        self.title.setAlignment(QtCore.Qt.AlignCenter)
        self.layout.addWidget(self.title)
        self.layout.addWidget(self.btn_min)
        self.layout.addWidget(self.btn_max)
        self.layout.addWidget(self.btn_close)

        self.title.setStyleSheet("""
            background-color: #404040;
            color: white;
            font-weight: bold;
            font-family : Trebuchet MS;
            font-size : 30px;
        """)
        self.setLayout(self.layout)

        self.start = QtCore.QPoint(0, 0)
        self.pressing = False

    def resizeEvent(self, QResizeEvent):
        super(MyBar, self).resizeEvent(QResizeEvent)
        self.title.resize(self.parent.width(), self.parent.height())

    def mousePressEvent(self, event):
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

    def btn_close_clicked(self):
        self.parent.close()

    def btn_max_clicked(self):
        if int(self.parent.windowState()) == 0:
            self.parent.showMaximized()
            self.btn_max.setIcon(QtGui.QIcon('Assets:max_inv.png'))
        elif int(self.parent.windowState()) == 2:
            self.parent.showNormal()
            self.btn_max.setIcon(QtGui.QIcon('Assets:max.png'))
        
    def btn_min_clicked(self):
        self.parent.showMinimized()

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        self.fenetre = QtWidgets.QWidget(MainWindow)
        self.fenetre.setObjectName("fenetre")
        
        self.main_layout = QtWidgets.QVBoxLayout(self.fenetre)
        self.main_layout.setObjectName("page_layout")

        self.main_layout.addWidget(MyBar(MainWindow))
        self.main_layout.setContentsMargins(0,0,0,0)
        MainWindow.resize(1400,800)
        MainWindow.setWindowFlags(QtCore.Qt.FramelessWindowHint)

        self.page_layout = QtWidgets.QVBoxLayout()
        self.page_layout.setObjectName("page_layout")
        self.page_layout.setContentsMargins(5,5,5,5)

        self.tables_titles_layout = QtWidgets.QHBoxLayout()
        self.tables_titles_layout.setObjectName("table_layout")

        self.before_label = QtWidgets.QLabel(self.fenetre)
        self.before_label.setObjectName("before_label")
        self.before_label.setAlignment(QtCore.Qt.AlignCenter)
        self.before_label.setStyleSheet("QLabel {background-color: #242526; border: 1px solid #32414B;}")
        self.tables_titles_layout.addWidget(self.before_label, 2)

        self.after_label = QtWidgets.QLabel(self.fenetre)
        self.after_label.setObjectName("after_label")
        self.after_label.setAlignment(QtCore.Qt.AlignCenter)
        self.after_label.setStyleSheet("QLabel {background-color: #242526; border: 1px solid #32414B;}")
        self.tables_titles_layout.addWidget(self.after_label, 2)

        self.result_label = QtWidgets.QLabel(self.fenetre)
        self.result_label.setObjectName("result_label")
        self.result_label.setAlignment(QtCore.Qt.AlignCenter)
        self.result_label.setStyleSheet("QLabel {background-color: #242526; border: 1px solid #32414B;}")
        self.tables_titles_layout.addWidget(self.result_label, 3)

        self.page_layout.addLayout(self.tables_titles_layout)

        self.tables_layout = QtWidgets.QHBoxLayout()
        self.tables_layout.setObjectName("tables_layout")
        self.data_table_before = QtWidgets.QTableView(self.fenetre)
        self.data_table_before.setObjectName("data_table_before")
        self.tables_layout.addWidget(self.data_table_before, 2)

        self.data_table_after = QtWidgets.QTableView(self.fenetre)
        self.data_table_after.setObjectName("data_table_after")
        self.tables_layout.addWidget(self.data_table_after, 2)

        self.result_table = QtWidgets.QTableView(self.fenetre)
        self.result_table.setObjectName("result_table")
        self.tables_layout.addWidget(self.result_table, 3)

        self.page_layout.addLayout(self.tables_layout)

        self.buttons_layout = QtWidgets.QHBoxLayout()
        self.buttons_layout.setObjectName("buttons_layout")
        self.open_button = QtWidgets.QPushButton(self.fenetre)
        self.open_button.setObjectName("open_button")
        self.open_button.setIcon(QtGui.QIcon('Assets:open.png'))
        self.open_button.setIconSize(QtCore.QSize(35, 35))
        self.open_button.clicked.connect(lambda : self.open(MainWindow))
        self.buttons_layout.addWidget(self.open_button)
        self.save_button = QtWidgets.QPushButton(self.fenetre)
        self.save_button.setObjectName("save_button")
        self.save_button.setIcon(QtGui.QIcon('Assets:save.png'))
        self.save_button.setIconSize(QtCore.QSize(35, 35))
        self.save_button.setEnabled(False)
        self.save_button.clicked.connect(lambda : self.save(MainWindow))
        self.buttons_layout.addWidget(self.save_button)
        
        self.page_layout.addLayout(self.buttons_layout)

        self.main_layout.addLayout(self.page_layout)

        MainWindow.setCentralWidget(self.fenetre)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)     

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Déplacement céphalométrique"))
        self.before_label.setText(_translate("MainWindow", "Avant planification"))
        self.after_label.setText(_translate("MainWindow", "Après planification"))
        self.result_label.setText(_translate("MainWindow", "Déplacement planifié"))
        self.open_button.setText(_translate("MainWindow", "Ouvrir"))
        self.save_button.setText(_translate("MainWindow", "Sauvegarder"))

    def open(self, MainWindow):
        self.save_button.setEnabled(False)
        self.data_table_before.setModel(None)
        self.data_table_after.setModel(None)
        self.result_table.setModel(None)

        self.file, ok = QtWidgets. QFileDialog.getOpenFileName(MainWindow, 'Ouvrir', os.getenv('HOME'), 'XML(*.xml)')
        if ok:
            self.df = self.create_dataframe()

            self.check_mistake()

            self.break_point_inf = [value for value in self.df.index.tolist() if value[:2] == '00']
            self.break_point_sup = [value for value in self.df.index.tolist() if value[:2] != '00']
            self.df_before = self.df.loc[:self.break_point_inf[-1]]
            self.df_after = self.df.loc[self.break_point_sup[0]:]

            self.df_before = self.df.loc[:self.break_point_inf[-1]].reset_index()
            self.df_after = self.df.loc[self.break_point_sup[0]:].reset_index()
            self.df_before['Repère'] = self.df_before['Repère'].apply(lambda x : x[2:])
            self.df_after['Repère'] = self.df_after['Repère'].apply(lambda x : x[2:])
            self.df_before.set_index('Repère', inplace=True)
            self.df_after.set_index('Repère', inplace=True)

            model = PandasModel(self.df_before)
            self.data_table_before.setModel(model)
            self.data_table_before.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
            self.data_table_before.verticalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
            
            model = PandasModel(self.df_after)
            self.data_table_after.setModel(model)
            self.data_table_after.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
            self.data_table_after.verticalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        
            self.calc_dif()

            model = PandasModel(self.df_dif)
            self.result_table.setModel(model)
            self.result_table.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
            self.result_table.verticalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
   

    def create_dataframe(self):
        df = pd.read_xml(self.file, encoding='utf-16')
        df = df.sort_values(by='Name', key=lambda col: col.str.lower()).reset_index(drop=True)
        df[['X\n\n', 'Y\n\n', 'Z\n\n']] = df['Coordinate'].str.split('  ', expand = True)
        df.drop('Coordinate', axis = 1, inplace = True)
        for col in ['X\n\n', 'Y\n\n', 'Z\n\n']:
            df[col] = df[col].apply(lambda x : float(x))
        df.rename(columns={'Name':'Repère'}, inplace = True)
        df.set_index('Repère', inplace = True)
        return (df)
    
    def check_mistake(self):
        mistakes = []
        values = self.df.index.tolist()
        for value in values:
            occ = [x for x in values if x[2:] == value[2:]]

            if len(occ) != 2:
                mistakes.append(value)
        if mistakes:

            msg = QtWidgets.QMessageBox()
            msg.setIcon(QtWidgets.QMessageBox.Critical)
            if len(mistakes) > 1:
                msg.setText(f"Les points {', '.join(mistakes)} sont sans équivalences")
            else :
                msg.setText(f"Le point {mistakes[0]} est sans équivalence")
            msg.setWindowTitle("Erreur")
            msg.exec_()

            for mistake in mistakes:
                point = mistake[:1] + str((int(mistake[1]) - 1) * -1) + mistake[2:]

                self.df.loc[point] = ('x', 'x', 'x')
                self.df.sort_index(key=lambda x: x.str.lower(), axis = 0, inplace = True)


    def calc_dif(self):
        self.df_dif = pd.DataFrame(columns=['Repère', 'X\n+ : Décalage à gauche\n- : Décalage à droite', 'Y\n+ : recul\n- : avancée', 'Z\n+ : impaction\n- : épaction'])
        self.df_dif

        try:
            for value in self.df.index[:len(self.break_point_inf)]:
                row = [value[2:]]
                row.append(self.calc_dif_value('X\n\n', value))
                row.append(self.calc_dif_value('Y\n\n', value))
                row.append(self.calc_dif_value('Z\n\n', value))
                self.df_dif.loc[len(self.df_dif)] = row

            self.df_dif.set_index('Repère', inplace = True)
            self.save_button.setEnabled(True)
        except:
            pass

    def calc_dif_value(self, key, value):
        i = self.df[self.df.index == (value[:1] + '1' + value[2:])][key].values
        j = self.df[self.df.index == value][key].values

        try :
            return (float(i) - float(j))
        except:
            return ('x')


    def save(self, MainWindow):
        save_path, ok = QtWidgets.QFileDialog.getSaveFileName(MainWindow, 'Sauvegarder', self.file[:-4] + '.xlsx', 'EXCEL(*.xlsx)')

        if ok:
            with pd.ExcelWriter(save_path, engine='xlsxwriter') as writer:
                self.df_before.to_excel(writer, sheet_name='Déplacement', startrow=1 , startcol=0) 

                self.df_after.to_excel(writer, sheet_name='Déplacement', startrow=1 , startcol=self.df.shape[1]+3)

                self.df_dif.to_excel(writer, sheet_name='Déplacement', startrow=1, startcol=self.df.shape[1]*2+6) 

                format = writer.book.add_format({'num_format': '0.0'})
                format_neg = writer.book.add_format({'num_format': '0.0', 'color' : 'red'})
                format_pos = writer.book.add_format({'num_format': '0.0', 'color' : 'green'})
                format_miss = writer.book.add_format({'bg_color' : '#fa0000'})

                merge_format = writer.book.add_format({"bold": 1, "border": 1, "align": "center", "valign": "vcenter", "fg_color": "#008db9"})

                writer.sheets['Déplacement'].conditional_format(f'{chr(65 + self.df.shape[1]*2+7)}2:{chr(65 + self.df.shape[1]*2+6 + self.df_dif.shape[1])}{self.df_dif.shape[0]+2}', {'type' : 'cell', 'criteria' : '<', 'value' : 0, 'format' : format_neg})
                writer.sheets['Déplacement'].conditional_format(f'{chr(65 + self.df.shape[1]*2+7)}2:{chr(65 + self.df.shape[1]*2+6 + self.df_dif.shape[1])}{self.df_dif.shape[0]+2}', {'type' : 'cell', 'criteria' : '>=', 'value' : 0, 'format' : format_pos})
                writer.sheets['Déplacement'].conditional_format(f'A2:{chr(65 + self.df.shape[1]*2+6 + self.df_dif.shape[1])}{self.df_dif.shape[0]+2}',{'type' : 'cell', 'criteria' : '=', 'value' : 'x', 'format' : format_miss})

                writer.sheets['Déplacement'].add_table(f'A2:{chr(65 + self.df.shape[1])}{len(self.break_point_inf)+2}', {'style': 'Table Style Medium 2','header_row': False})
                writer.sheets['Déplacement'].add_table(f'{chr(65 + self.df.shape[1]+3)}2:{chr(65 + self.df.shape[1]*2+3)}{len(self.break_point_sup)+2}', {'style': 'Table Style Medium 2','header_row': False})
                writer.sheets['Déplacement'].add_table(f'{chr(65 + self.df.shape[1]*2+6)}2:{chr(65 + self.df.shape[1]*2+6 + self.df_dif.shape[1])}{self.df_dif.shape[0]+2}', {'style': 'Table Style Light 9','header_row': False})

                writer.sheets['Déplacement'].merge_range(f"A1:{chr(65 + self.df.shape[1])}1", "Avant planification", merge_format)
                writer.sheets['Déplacement'].merge_range(f"{chr(65 + self.df.shape[1]+3)}1:{chr(65 + self.df.shape[1]*2+3)}1", "Après planification", merge_format)
                writer.sheets['Déplacement'].merge_range(f"{chr(65 + self.df.shape[1]*2+6)}1:{chr(65 + self.df.shape[1]*3+6)}1", "Déplacement planifié", merge_format)

                writer.sheets['Déplacement'].set_column(f'B:{chr(65 + self.df.shape[1])}', None, format)
                writer.sheets['Déplacement'].set_column(f'{chr(65 + self.df.shape[1]+3)}:{chr(65 + self.df.shape[1]+3 + self.df_dif.shape[1])}', None, format)

                writer.sheets['Déplacement'].set_row(1, 40)

                writer.sheets['Déplacement'].set_column(f'{chr(65 + self.df.shape[1]*2+7)}:{chr(65 + self.df.shape[1]*2+6 + self.df_dif.shape[1])}', 22)

def start():

    app = QtWidgets.QApplication(sys.argv)

    QtCore.QDir.addSearchPath('Assets', os.path.dirname(os.path.abspath(__file__)))
    app.setWindowIcon(QtGui.QIcon("Assets:Logo_small.png"))
    file = QtCore.QFile('Assets:Style.qss')
    file.open(QtCore.QFile.ReadOnly | QtCore.QFile.Text)
    app.setStyleSheet(str(file.readAll(), 'utf-8'))

    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())

start()