import io
import sys
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
    <width>1114</width>
    <height>523</height>
   </rect>
  </property>
  <property name="minimumSize">
   <size>
    <width>1114</width>
    <height>523</height>
   </size>
  </property>
  <property name="maximumSize">
   <size>
    <width>1114</width>
    <height>523</height>
   </size>
  </property>
  <property name="windowTitle">
   <string>MainWindow</string>
  </property>
  <widget class="QWidget" name="centralwidget">
   <widget class="QLineEdit" name="line_box">
    <property name="geometry">
     <rect>
      <x>970</x>
      <y>40</y>
      <width>113</width>
      <height>20</height>
     </rect>
    </property>
   </widget>
   <widget class="QPushButton" name="file_button">
    <property name="geometry">
     <rect>
      <x>640</x>
      <y>460</y>
      <width>231</width>
      <height>21</height>
     </rect>
    </property>
    <property name="text">
     <string>Открыть путь к изменяемому файлу</string>
    </property>
   </widget>
   <widget class="QPushButton" name="db_button">
    <property name="geometry">
     <rect>
      <x>900</x>
      <y>460</y>
      <width>191</width>
      <height>21</height>
     </rect>
    </property>
    <property name="text">
     <string>Открыть путь к базе данных</string>
    </property>
   </widget>
   <widget class="QPushButton" name="save_button">
    <property name="geometry">
     <rect>
      <x>960</x>
      <y>80</y>
      <width>141</width>
      <height>21</height>
     </rect>
    </property>
    <property name="text">
     <string>Сохранить изменения</string>
    </property>
   </widget>
   <widget class="QTableWidget" name="tableWidget">
    <property name="geometry">
     <rect>
      <x>20</x>
      <y>10</y>
      <width>891</width>
      <height>241</height>
     </rect>
    </property>
   </widget>
  </widget>
  <widget class="QMenuBar" name="menubar">
   <property name="geometry">
    <rect>
     <x>0</x>
     <y>0</y>
     <width>1114</width>
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

    def text_to_PlainText(self, data):
        for index, elem in enumerate(data):
            data[index] = f'A{index + 1}: '+ elem
        self.text_db.setPlainText(''.join(data))

    def save(self):
        self.data.save(self.dialog)

    def open_db(self):
        self.dialog_db = QFileDialog.getOpenFileName(self, 'Открыть путь', '', filter='Блокнот (*.txt)')[0]
        self.setGeometry(400, 400, *SCREEN_SIZE)
        try:
            with open(self.dialog_db, 'r', encoding='utf-8') as file:
                data = file.readlines()
        except Exception:
            self.statusBar().showMessage('Не указан файл')
            return
        print(1)
        self.text_to_PlainText(data)

    def redactor(self):
        item = self.line_box.text()
        if not (item):
            self.statusBar().showMessage('Отсутсвуют клетки')
            return
        item = item.split()
        try:
            for box in item:
                print(box)
                self.active_data[box] = 3
            self.save()
        except Exception:
            self.statusBar().showMessage('Не указан файл')
            return


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = Redactor()
    ex.show()
    sys.exit(app.exec())
