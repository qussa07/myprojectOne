import io
import sys
import os
from datetime import date
import openpyxl
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QMessageBox
from PyQt5 import uic, QtCore

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
      <x>140</x>
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
      <y>100</y>
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
   <widget class="QDateEdit" name="select_date">
    <property name="geometry">
     <rect>
      <x>500</x>
      <y>40</y>
      <width>110</width>
      <height>22</height>
     </rect>
    </property>
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
        self.setWindowTitle('Массовое изменение xlsx файлов')
        self.db_data = None
        self.return_real_date()
        self.db_button.clicked.connect(self.open_db)
        self.db_button.setEnabled(False)
        self.file_button.clicked.connect(self.open_file)
        self.save_button.clicked.connect(self.redactor)
        self.save_button.setEnabled(False)

    def open_file(self):
        self.dialog = QFileDialog.getOpenFileName(self, 'Открыть путь к файлу', '', filter='Лист Excel (*.xlsx)')[0]
        self.setGeometry(400, 400, *SCREEN_SIZE)
        try:
            self.data = openpyxl.load_workbook(self.dialog)
            self.active_data = self.data.active
        except Exception:
            self.statusBar().showMessage('Не указан файл')
            self.error_message('Ошибка!', 'Не указан нужный файл')
            return
        self.check()

    def save(self, i):
        directory = os.path.abspath(i)
        if not(directory):
            os.mkdir(directory)
        date = self.select_date.dateTime().toString('dd-MM-yyyy')
        self.data.save(directory + '/' + f'{date}-sm.xlsx')
        self.statusBar().showMessage('Успешно!')
        self.information_message('Успешно!', 'Создание и изменение файлов прошло успешно.')

    def open_db(self):
        self.dialog_db = QFileDialog.getOpenFileName(self, 'Открыть путь', '', filter='Блокнот (*.txt)')[0]
        self.setGeometry(400, 400, *SCREEN_SIZE)
        try:
            with open(self.dialog_db, 'r', encoding='utf-8') as file:
                self.db_data = file.read().split(';')
        except Exception:
            self.statusBar().showMessage('Не указан файл')
            self.error_message('Ошибка!', 'Не указан нужный файл')
            return
        self.check()

    def redactor(self):
        try:
            if not (self.db_data[0]):
                self.statusBar().showMessage('База данных пуста')
                return
            for i in self.db_data:
                self.active_data['B1'] = i
                date_from_QDateEdit = self.select_date.dateTime().toString('dd.MM.yyyy')
                self.active_data['J1'] = date_from_QDateEdit
                self.save(i)


        except Exception as file:
            self.error_message('Ошибка!', 'Не указан файл')
            self.statusBar().showMessage('Не указан файл')
            return

    def return_real_date(self):
        year = int(str(date.today()).split('-')[0])
        day = int(str(date.today()).split('-')[1])
        mouth = int(str(date.today()).split('-')[2])
        self.select_date.setDate(QtCore.QDate(year, day, mouth))
        self.select_date.setDisplayFormat("dd.MM.yyyy")

    def check(self):
        print(1)
        if self.data:
            self.db_button.setEnabled(True)
            if self.db_data:
                self.save_button.setEnabled(True)

    def error_message(self, title, message):
        error_dialog = QMessageBox(self)
        error_dialog.setIcon(QMessageBox.Critical)
        error_dialog.setWindowTitle(title)
        error_dialog.setText(message)
        error_dialog.exec_()

    def information_message(self, title, message):
        error_dialog = QMessageBox(self)
        error_dialog.setIcon(QMessageBox.Information)
        error_dialog.setWindowTitle(title)
        error_dialog.setText(message)
        error_dialog.exec_()



if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = Redactor()
    ex.show()
    sys.exit(app.exec())
