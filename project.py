import io
import sys
import os
from datetime import date, datetime, timedelta
import openpyxl
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QMessageBox, QSizePolicy
import subprocess
from PyQt5 import uic, QtCore
from PyQt5.QtCore import Qt
from PyQt5.QtCore import QThread, pyqtSignal
import time
from play import TicTacToe

template = '''<?xml version="1.0" encoding="UTF-8"?>
<ui version="4.0">
 <class>MainWindow</class>
 <widget class="QMainWindow" name="MainWindow">
  <property name="geometry">
   <rect>
    <x>0</x>
    <y>0</y>
    <width>760</width>
    <height>305</height>
   </rect>
  </property>
  <property name="windowTitle">
   <string>MainWindow</string>
  </property>
  <widget class="QWidget" name="centralwidget">
   <layout class="QGridLayout" name="gridLayout">
    <item row="0" column="0">
     <widget class="QTabWidget" name="tabWidget">
      <property name="sizePolicy">
       <sizepolicy hsizetype="Ignored" vsizetype="Ignored">
        <horstretch>1</horstretch>
        <verstretch>1</verstretch>
       </sizepolicy>
      </property>
      <property name="minimumSize">
       <size>
        <width>735</width>
        <height>246</height>
       </size>
      </property>
      <property name="font">
       <font>
        <pointsize>8</pointsize>
       </font>
      </property>
      <property name="contextMenuPolicy">
       <enum>Qt::ActionsContextMenu</enum>
      </property>
      <property name="whatsThis">
       <string>&lt;html&gt;&lt;head/&gt;&lt;body&gt;&lt;p&gt;dfsdafasdf&lt;/p&gt;&lt;/body&gt;&lt;/html&gt;</string>
      </property>
      <property name="layoutDirection">
       <enum>Qt::LeftToRight</enum>
      </property>
      <property name="autoFillBackground">
       <bool>false</bool>
      </property>
      <property name="currentIndex">
       <number>0</number>
      </property>
      <property name="elideMode">
       <enum>Qt::ElideNone</enum>
      </property>
      <property name="movable">
       <bool>false</bool>
      </property>
      <widget class="QWidget" name="tab">
       <attribute name="title">
        <string>Меню программы</string>
       </attribute>
       <widget class="QWidget" name="gridLayoutWidget">
        <property name="geometry">
         <rect>
          <x>90</x>
          <y>30</y>
          <width>486</width>
          <height>83</height>
         </rect>
        </property>
        <layout class="QGridLayout" name="gridLayout_2">
         <item row="0" column="1">
          <widget class="QDateEdit" name="select_date">
           <property name="sizePolicy">
            <sizepolicy hsizetype="Minimum" vsizetype="Fixed">
             <horstretch>1</horstretch>
             <verstretch>1</verstretch>
            </sizepolicy>
           </property>
          </widget>
         </item>
         <item row="2" column="0">
          <layout class="QHBoxLayout" name="horizontalLayout">
           <item>
            <widget class="QPushButton" name="file_button">
             <property name="sizePolicy">
              <sizepolicy hsizetype="Minimum" vsizetype="Fixed">
               <horstretch>1</horstretch>
               <verstretch>1</verstretch>
              </sizepolicy>
             </property>
             <property name="text">
              <string>Открыть путь к изменяемому файлу</string>
             </property>
            </widget>
           </item>
           <item>
            <widget class="QPushButton" name="db_button">
             <property name="sizePolicy">
              <sizepolicy hsizetype="Minimum" vsizetype="Fixed">
               <horstretch>1</horstretch>
               <verstretch>1</verstretch>
              </sizepolicy>
             </property>
             <property name="text">
              <string>Открыть путь к базе данных</string>
             </property>
            </widget>
           </item>
          </layout>
         </item>
         <item row="0" column="0">
          <widget class="QPushButton" name="save_button">
           <property name="sizePolicy">
            <sizepolicy hsizetype="Minimum" vsizetype="Fixed">
             <horstretch>1</horstretch>
             <verstretch>1</verstretch>
            </sizepolicy>
           </property>
           <property name="focusPolicy">
            <enum>Qt::NoFocus</enum>
           </property>
           <property name="layoutDirection">
            <enum>Qt::LeftToRight</enum>
           </property>
           <property name="text">
            <string>Сохранить изменения</string>
           </property>
          </widget>
         </item>
         <item row="1" column="1">
          <widget class="QLabel" name="label">
           <property name="sizePolicy">
            <sizepolicy hsizetype="Expanding" vsizetype="Expanding">
             <horstretch>0</horstretch>
             <verstretch>0</verstretch>
            </sizepolicy>
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
         </item>
         <item row="2" column="1">
          <widget class="QSpinBox" name="days">
           <property name="sizePolicy">
            <sizepolicy hsizetype="Minimum" vsizetype="Fixed">
             <horstretch>1</horstretch>
             <verstretch>1</verstretch>
            </sizepolicy>
           </property>
          </widget>
         </item>
         <item row="1" column="0">
          <widget class="QProgressBar" name="progress">
           <property name="sizePolicy">
            <sizepolicy hsizetype="Expanding" vsizetype="Expanding">
             <horstretch>0</horstretch>
             <verstretch>0</verstretch>
            </sizepolicy>
           </property>
           <property name="value">
            <number>24</number>
           </property>
          </widget>
         </item>
        </layout>
       </widget>
      </widget>
      <widget class="QWidget" name="tab_2">
       <attribute name="title">
        <string>О создателе</string>
       </attribute>
       <widget class="QWidget" name="gridLayoutWidget_2">
        <property name="geometry">
         <rect>
          <x>0</x>
          <y>0</y>
          <width>711</width>
          <height>181</height>
         </rect>
        </property>
        <layout class="QGridLayout" name="gridLayout_3">
         <item row="1" column="0">
          <widget class="QLabel" name="program">
           <property name="sizePolicy">
            <sizepolicy hsizetype="Expanding" vsizetype="Expanding">
             <horstretch>0</horstretch>
             <verstretch>0</verstretch>
            </sizepolicy>
           </property>
           <property name="font">
            <font>
             <family>Segoe UI Black</family>
             <pointsize>10</pointsize>
             <weight>75</weight>
             <bold>true</bold>
            </font>
           </property>
           <property name="text">
            <string>TextLabel</string>
           </property>
          </widget>
         </item>
         <item row="1" column="1">
          <spacer name="horizontalSpacer">
           <property name="orientation">
            <enum>Qt::Horizontal</enum>
           </property>
           <property name="sizeType">
            <enum>QSizePolicy::Fixed</enum>
           </property>
           <property name="sizeHint" stdset="0">
            <size>
             <width>10</width>
             <height>10</height>
            </size>
           </property>
          </spacer>
         </item>
         <item row="1" column="2">
          <widget class="QLabel" name="creator">
           <property name="sizePolicy">
            <sizepolicy hsizetype="Expanding" vsizetype="Expanding">
             <horstretch>0</horstretch>
             <verstretch>0</verstretch>
            </sizepolicy>
           </property>
           <property name="font">
            <font>
             <family>Segoe UI Black</family>
             <pointsize>10</pointsize>
             <weight>75</weight>
             <bold>true</bold>
            </font>
           </property>
           <property name="text">
            <string>TextLabel</string>
           </property>
          </widget>
         </item>
         <item row="2" column="0" colspan="3">
          <widget class="QLabel" name="link">
           <property name="sizePolicy">
            <sizepolicy hsizetype="Expanding" vsizetype="Expanding">
             <horstretch>0</horstretch>
             <verstretch>0</verstretch>
            </sizepolicy>
           </property>
           <property name="font">
            <font>
             <family>Playbill</family>
             <pointsize>10</pointsize>
            </font>
           </property>
           <property name="text">
            <string>TextLabel</string>
           </property>
          </widget>
         </item>
        </layout>
       </widget>
      </widget>
     </widget>
    </item>
   </layout>
  </widget>
  <widget class="QMenuBar" name="menubar">
   <property name="geometry">
    <rect>
     <x>0</x>
     <y>0</y>
     <width>760</width>
     <height>21</height>
    </rect>
   </property>
   <widget class="QMenu" name="menu">
    <property name="title">
     <string>Игры</string>
    </property>
    <addaction name="play"/>
    <addaction name="play_2"/>
   </widget>
   <addaction name="menu"/>
  </widget>
  <widget class="QStatusBar" name="statusbar"/>
  <action name="play">
   <property name="text">
    <string>Змейка</string>
   </property>
  </action>
  <action name="play_2">
   <property name="text">
    <string>Крестики  нолики на двоих</string>
   </property>
  </action>
 </widget>
 <resources/>
 <connections/>
</ui>
'''
SCREEN_SIZE = [735, 257]


class Redactor(QMainWindow):
    def __init__(self):
        super(Redactor, self).__init__()
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
        self.progress.setValue(0)
        self.program.setText('Программа для массового изменения xlsx файлов \n'
                             'предназначена для автоматизации процесса\n '
                             'изменения данных в больших объемах Excel файлов.\n'
                             ' С ее помощью пользователи могут легко и \n'
                             'быстро вносить изменения в структуру, форматирование\n'
                             ' и содержание xlsx файлов без \n'
                             'необходимости открывать каждый файл вручную.')
        self.link.setText("<a href=\"https://vk.com/ya_wirr/\">Помощь/Связь со мной!</a>")
        self.link.setTextFormat(Qt.RichText)
        self.link.setTextInteractionFlags(Qt.TextBrowserInteraction)
        self.link.setOpenExternalLinks(True)
        self.db_button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.file_button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.save_button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.select_date.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.days.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.program.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.link.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.tab.setLayout(self.gridLayout_2)
        self.tab.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.tab_2.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.program.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.link.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.link.setAlignment(Qt.AlignCenter)
        self.program.setAlignment(Qt.AlignCenter)
        self.creator.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.creator.setAlignment(Qt.AlignCenter)
        self.tab_2.setLayout(self.gridLayout_3)
        self.play_2.triggered.connect(self.play_)
        self.play.triggered.connect(self.play_1)
        self.creator.setText('Благодаря данной программе пользователи могут\n значительно упростить\n'
                             ' и ускорить процесс изменения данных\n в больших объемах xlsx файлов, \n'
                             'что делает ее незаменимым инструментом для\n работы с данными в офисной \n'
                             'среде и других областях, где требуется массовая\n обработка Excel файлов')
        self.file_button.setStyleSheet('QPushButton {'
                                       'border: 2px solid rgba(255, 0, 17, 239)'
                                       '}'
                                       'QPushButton:hover {'
                                       'background-color: rgba(199, 250, 90, 100);'
                                       'border-image: url(Без имени.png);'
                                       '}')

    def open_file(self):
        '''
        Открывает и читает основной файл.
        '''
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
        """
        Сохраняет файл и меняет название на нужное.
        """
        directory = os.path.abspath(name)
        if not (os.path.exists(directory)):
            os.mkdir(directory)
        self.data.save(directory + '/' + f'{date}-sm.xlsx')
        self.count += 1
        if self.count == len(self.db_data) * self.days.value():
            self.сongratulations()
            self.count = 0

    def open_db(self):
        '''
        Открывает базу данных и читает ее.
        '''
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
        '''
        Основная часть программы.
        Создает файлы и выставляет в нужны ячейки нужную информацию.
        '''
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
        '''
        Выставляет сегодняшнюю дату.
        '''
        year = int(str(date.today()).split('-')[0])
        day = int(str(date.today()).split('-')[1])
        mouth = int(str(date.today()).split('-')[2])
        self.select_date.setDate(QtCore.QDate(year, day, mouth))
        self.select_date.setDisplayFormat("dd.MM.yyyy")

    def check(self):
        '''
        Проверяет присутствует ли файл в переменной.
        '''
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
        '''
        Выводит сообщение об ошибке.
        '''
        error_dialog = QMessageBox(self)
        error_dialog.setIcon(QMessageBox.Critical)
        error_dialog.setWindowTitle(title)
        error_dialog.setText(message)
        error_dialog.exec_()

    def information_message(self, title, message):
        '''
            Выводит сообщение, которое оповещает пользователя об успешном создании файлов/файла.
        '''
        error_dialog = QMessageBox(self)
        error_dialog.setIcon(QMessageBox.Information)
        error_dialog.setWindowTitle(title)
        error_dialog.setText(message)
        error_dialog.exec_()

    def сongratulations(self):
        '''
            Запускается когда закончился цикл создания файлов.
        '''
        self.progress_bar()
        self.statusBar().showMessage('Успешно!')
        self.information_message('Успешно!', 'Создание и изменение файлов прошло успешно.')

    def progress_bar(self):
        for i in range(101):
            time.sleep(0.01)
            self.progress.setValue(i)

    def play_(self):
        subprocess.run(["python", "play.py"])

    def play_1(self):
        subprocess.run(["python", "Play_2.py"])

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = Redactor()
    ex.show()
    sys.exit(app.exec())
