import io
import sys
import os
from datetime import date, datetime, timedelta
import openpyxl
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QMessageBox
from PyQt5 import uic, QtCore

template = '''<?xml version="1.0" encoding="UTF-8"?>
<ui version="4.0">
 <class>MainWindow</class>
 <widget class="QMainWindow" name="MainWindow">
  <property name="enabled">
   <bool>true</bool>
  </property>
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
  <property name="dockOptions">
   <set>QMainWindow::AllowTabbedDocks|QMainWindow::AnimatedDocks</set>
  </property>
  <property name="unifiedTitleAndToolBarOnMac">
   <bool>false</bool>
  </property>
  <widget class="QWidget" name="centralwidget">
   <widget class="QTabWidget" name="tabWidget">
    <property name="geometry">
     <rect>
      <x>0</x>
      <y>0</y>
      <width>731</width>
      <height>211</height>
     </rect>
    </property>
    <property name="whatsThis">
     <string>&lt;html&gt;&lt;head/&gt;&lt;body&gt;&lt;p&gt;dfsdafasdf&lt;/p&gt;&lt;/body&gt;&lt;/html&gt;</string>
    </property>
    <property name="currentIndex">
     <number>0</number>
    </property>
    <widget class="QWidget" name="tab">
     <attribute name="title">
      <string>Меню программы</string>
     </attribute>
     <widget class="QWidget" name="horizontalLayoutWidget">
      <property name="geometry">
       <rect>
        <x>20</x>
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
     <widget class="QPushButton" name="save_button">
      <property name="geometry">
       <rect>
        <x>40</x>
        <y>40</y>
        <width>281</width>
        <height>31</height>
       </rect>
      </property>
      <property name="text">
       <string>Сохранить изменения</string>
      </property>
     </widget>
     <widget class="QDateEdit" name="select_date">
      <property name="geometry">
       <rect>
        <x>350</x>
        <y>50</y>
        <width>110</width>
        <height>22</height>
       </rect>
      </property>
     </widget>
     <widget class="QSpinBox" name="days">
      <property name="geometry">
       <rect>
        <x>530</x>
        <y>50</y>
        <width>42</width>
        <height>22</height>
       </rect>
      </property>
     </widget>
     <widget class="QLabel" name="label">
      <property name="geometry">
       <rect>
        <x>500</x>
        <y>10</y>
        <width>131</width>
        <height>16</height>
       </rect>
      </property>
      <property name="font">
       <font>
        <pointsize>10</pointsize>
       </font>
      </property>
      <property name="text">
       <string>Количество дней:</string>
      </property>
     </widget>
    </widget>
    <widget class="QWidget" name="tab_2">
     <attribute name="title">
      <string>О создателе</string>
     </attribute>
    </widget>
   </widget>
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
        self.count = 0
        self.setWindowTitle('Массовое изменение xlsx файлов')
        self.db_data = None
        self.return_real_date()
        self.db_button.clicked.connect(self.open_db)
        self.db_button.setEnabled(False)
        self.file_button.clicked.connect(self.open_file)
        self.save_button.clicked.connect(self.redactor)
        self.save_button.setEnabled(False)

        self.file_button.setStyleSheet('QPushButton {'
                                       'border: 2px solid rgba(255, 0, 17, 239)'
                                       '}'
                                       'QPushButton:hover {'
                                       'background-color: rgba(199, 250, 90, 100);'
                                       'border-image: url(Без имени.png);'
                                       '}')
        # self.setStyleSheet('border-image: url(shutterstock_478784707.jpg)')

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

    def save(self, name, date):
        directory = os.path.abspath(name)
        if not (os.path.exists(directory)):
            os.mkdir(directory)
        self.data.save(directory + '/' + f'{date}-sm.xlsx')
        self.count += 1
        if self.count == len(self.db_data):
            self.сongratulations()
            self.count = 0

    def open_db(self):
        self.dialog_db = QFileDialog.getOpenFileName(self, 'Открыть путь', '', filter='Блокнот (*.txt)')[0]
        self.setGeometry(400, 400, *SCREEN_SIZE)
        try:
            with open(self.dialog_db, 'r', encoding='utf-8') as file:
                self.db_data = file.read().split(';')
                print(self.db_data)
        except Exception:
            self.statusBar().showMessage('Не указан файл')
            self.error_message('Ошибка!', 'Не указан нужный файл')
            return
        self.check()

    def redactor(self):
        if not (self.days.value()):
            self.error_message('Ошибка!', 'Не указано количество дней')
            self.statusBar().showMessage('Не указан файл')
            return
        try:
            if not (self.db_data[0]):
                self.statusBar().showMessage('База данных пуста')
                return
            for name in self.db_data:
                date = self.select_date.dateTime().toString('dd.MM.yyyy')
                date_from_QDateEdit = datetime.strptime(date, '%d.%m.%Y').date()
                for i in range(self.days.value()):
                    self.active_data['B1'] = name
                    self.active_data['J1'] = date_from_QDateEdit.strftime('%d.%m.%Y')
                    self.save(name, date_from_QDateEdit.strftime('%d-%m-%Y'))
                    date_from_QDateEdit = date_from_QDateEdit + timedelta(days=1)
        except Exception:
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
        if self.data:
            self.db_button.setEnabled(True)
            self.db_button.setStyleSheet('QPushButton {'
                                         'border: 2px solid rgba(255, 0, 17, 239)'
                                         '}'
                                         'QPushButton:hover {'
                                         'background-color: rgba(199, 250, 90, 100);'
                                         'border-image: url(Без имени.png);'
                                         '}')
            self.file_button.setStyleSheet('QPushButton {'
                                           'border: 2px solid rgba(0, 255, 9, 239)'
                                           '}'
                                           'QPushButton:hover {'
                                           'background-color: rgba(199, 250, 90, 100);'
                                           'border-image: url(263a-fe0f.png);'
                                           '}')
            if self.db_data:
                self.db_button.setStyleSheet('QPushButton {'
                                             'border: 2px solid rgba(0, 255, 9, 239)'
                                             '}'
                                             'QPushButton:hover {'
                                             'background-color: rgba(199, 250, 90, 100);'
                                             'border-image: url(shutterstock_478784707.jpg);'
                                             '}')
                self.save_button.setEnabled(True)
                self.save_button.setStyleSheet('QPushButton {'
                                               'border: 2px solid rgba(255, 0, 17, 239)'
                                               '}'
                                               'QPushButton:hover {'
                                               'background-color: rgba(199, 250, 90, 100)'
                                               '}')

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

    def сongratulations(self):
        self.statusBar().showMessage('Успешно!')
        self.information_message('Успешно!', 'Создание и изменение файлов прошло успешно.')


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = Redactor()
    ex.show()
    sys.exit(app.exec())
