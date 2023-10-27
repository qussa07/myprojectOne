import io
import sys
from PIL import Image

from PyQt5.QtGui import QPixmap, QImage, QColor, QTransform
from PyQt5.QtWidgets import QApplication, QLabel, QMainWindow, QFileDialog
from PyQt5 import uic
from PyQt5.QtCore import QPoint

template = '''<?xml version="1.0" encoding="UTF-8"?>
<ui version="4.0">
 <class>MainWindow</class>
 <widget class="QMainWindow" name="MainWindow">
  <property name="geometry">
   <rect>
    <x>0</x>
    <y>0</y>
    <width>923</width>
    <height>456</height>
   </rect>
  </property>
  <property name="windowTitle">
   <string>MainWindow</string>
  </property>
  <widget class="QWidget" name="centralwidget">
   <widget class="QWidget" name="formLayoutWidget">
    <property name="geometry">
     <rect>
      <x>10</x>
      <y>20</y>
      <width>891</width>
      <height>351</height>
     </rect>
    </property>
    <layout class="QFormLayout" name="formLayout">
     <item row="2" column="1">
      <layout class="QVBoxLayout" name="verticalLayout">
       <item>
        <widget class="QPushButton" name="save_button">
         <property name="text">
          <string>Сохранить изменения</string>
         </property>
        </widget>
       </item>
       <item>
        <widget class="QPushButton" name="pushButton_2">
         <property name="text">
          <string/>
         </property>
        </widget>
       </item>
       <item>
        <spacer name="verticalSpacer">
         <property name="orientation">
          <enum>Qt::Vertical</enum>
         </property>
         <property name="sizeHint" stdset="0">
          <size>
           <width>20</width>
           <height>40</height>
          </size>
         </property>
        </spacer>
       </item>
       <item>
        <layout class="QHBoxLayout" name="horizontalLayout">
         <item>
          <widget class="QPushButton" name="db_button">
           <property name="text">
            <string>Открыть путь к базе данных</string>
           </property>
          </widget>
         </item>
         <item>
          <widget class="QPushButton" name="file_button">
           <property name="text">
            <string>Открыть путь к изменяемому файлу </string>
           </property>
          </widget>
         </item>
        </layout>
       </item>
      </layout>
     </item>
     <item row="1" column="1">
      <widget class="QPlainTextEdit" name="plainTextEdit"/>
     </item>
    </layout>
   </widget>
  </widget>
  <widget class="QMenuBar" name="menubar">
   <property name="geometry">
    <rect>
     <x>0</x>
     <y>0</y>
     <width>923</width>
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
        self.db_button.clicked.connect(self.open)
        self.file_button.clicked.connect(self.open)
        self.save_button.clicked.connect(self.save)

    def open(self):
        self.dialog = QFileDialog.getOpenFileName(self, 'Открыть путь к файлу', '', filter='Лист Excel (*.xlsx)')[0]
        self.setGeometry(400, 400, *SCREEN_SIZE)
        data = 0
        # ДОЛЖНА БЫТЬ ДАТА КОТОРАЯ ПЕРЕНОСИТСЯ В unzip
        self.unzip(data)
        # ТРЕБУЕТСЯ ДОПИСАТЬ ИБО НИЧЕ НЕ РАБОТАЕТ

    def unzip(self):
        pass

    def text_to_PlainText(self):
        pass

    def save(self):
        pass


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = Redactor()
    ex.show()
    sys.exit(app.exec())
