import io
import sys
import datetime
import openpyxl
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog
from PyQt5 import uic

template = '''<?xml version="1.0" encoding="UTF-8"?>
<ui version="4.0">
 <class>MainWindow</class>
 <widget class="QMainWindow" name="MainWindow">
  <property name="geometry">
   <rect>
    <x>0</x>
    <y>0</y>
    <width>735</width>
    <height>231</height>
   </rect>
  </property>
  <property name="minimumSize">
   <size>
    <width>735</width>
    <height>231</height>
   </size>
  </property>
  <property name="maximumSize">
   <size>
    <width>735</width>
    <height>231</height>
   </size>
  </property>
  <property name="windowTitle">
   <string>MainWindow</string>
  </property>
  <widget class="QWidget" name="centralwidget">
   <widget class="QPushButton" name="save_button">
    <property name="geometry">
     <rect>
      <x>230</x>
      <y>30</y>
      <width>281</width>
      <height>31</height>
     </rect>
    </property>
    <property name="text">
     <string>Сохранить изменения</string>
    </property>
   </widget>
   <widget class="QWidget" name="horizontalLayoutWidget">
    <property name="geometry">
     <rect>
      <x>50</x>
      <y>70</y>
      <width>659</width>
      <height>78</height>
     </rect>
    </property>
    <layout class="QHBoxLayout" name="horizontalLayout">
     <item>
      <widget class="QPushButton" name="file_button">
       <property name="text">
        <string>Открыть путь к изменяемому файлу</string>
       </property>
      </widget>
     </item>
     <item>
      <widget class="QPushButton" name="db_button">
       <property name="text">
        <string>Открыть путь к базе данных</string>
       </property>
      </widget>
     </item>
    </layout>
   </widget>
  </widget>
  <widget class="QMenuBar" name="menubar">
   <property name="geometry">
    <rect>
     <x>0</x>
     <y>0</y>
     <width>735</width>
     <height>21</height>
    </rect>
   </property>
  </widget>
  <widget class="QStatusBar" name="statusbar"/>
 </widget>
 <resources/>
 <connections/>
</ui>
'''
SCREEN_SIZE = [900, 900]


class Redactor(QMainWindow):
    def __init__(self):
        super().__init__()
        f = io.StringIO(template)
        uic.loadUi(f, self)
        self.db_button.clicked.connect(self.open_db)
        self.file_button.clicked.connect(self.open_file)
        self.save_button.clicked.connect(self.redactor)

    def open_file(self):
        self.dialog = QFileDialog.getOpenFileName(self, 'Открыть путь к файлу', '', filter='Лист Excel (*.xlsx)')[0]
        self.setGeometry(400, 400, *SCREEN_SIZE)
        try:
            self.data = openpyxl.load_workbook(self.dialog)
            self.active_data = self.data.active
        except Exception:
            self.statusBar().showMessage('Не указан файл')
            return

    def save(self, i):
        self.data.save(('/'.join(self.dialog.split('/')[0:-1]) + '/' + i + '.xlsx'))

    def open_db(self):
        self.dialog_db = QFileDialog.getOpenFileName(self, 'Открыть путь', '', filter='Блокнот (*.txt)')[0]
        self.setGeometry(400, 400, *SCREEN_SIZE)
        try:
            with open(self.dialog_db, 'r', encoding='utf-8') as file:
                self.db_data = file.read().split(';')
        except Exception:
            self.statusBar().showMessage('Не указан файл')
            return

    def redactor(self):
        try:
            if not (self.db_data[0]):
                self.statusBar().showMessage('База данных пуста')
                return
            for i in self.db_data:
                self.active_data['B1'] = i
                self.active_data['J1'] = datetime.date.today()
                self.save(i)


        except Exception as file:
            print(file)
            self.statusBar().showMessage('Не указан файл')
            return


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = Redactor()
    ex.show()
    sys.exit(app.exec())
