from  PyQt5 import QtWidgets, QtGui, QtCore
from  PyQt5.QtCore import QFileInfo
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QMessageBox

import docx
import re
import pandas as  pd #Установил ещё openpyxl, т.к. выдавал ошибку
from functools import partial
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from Ui_Diploma import Ui_DiplomaWindow

from Classes.SendMail import SendMail
from settings import settings

class Window3(QtWidgets.QMainWindow):
    def __init__(self):
        super(Window3, self).__init__()
        self.ui = Ui_DiplomaWindow()
        self.ui.setupUi(self)
        
        self.ui.pushButton.clicked.connect(self.diploma_sample)
        self.ui.pushButton_2.clicked.connect(self.databased)
        self.ui.pushButton_3.clicked.connect(self.add_labels_to_diploma_more)
        self.ui.pushButton_4.clicked.connect(self.send_diploma_to_email)
        self.ui.pushButton_6.clicked.connect(self.delete_label_text_diploma)
        self.ui.pushButton_7.clicked.connect(self.delete_label_text_databased)

        self.__mail = SendMail(settings.mail_server,
                               settings.mail_port,
                               settings.mail_login,
                               settings.mail_password)
        
    
    #Выбор файлов
    def diploma_sample(self):
        self.file_dialog = QFileDialog(self, 'Выбрать файл', 'C://')
        self.file_dialog.setFileMode(QFileDialog.ExistingFile)
        self.file_dialog.setNameFilter("Microsoft Word (*.doc *.docx)")
        if self.file_dialog.exec_() == QFileDialog.Accepted:
            self.file_choice = self.file_dialog.selectedFiles()[0]
            if QFileInfo(self.file_choice).suffix() not in ["doc", "docx"]:
                QMessageBox.warning(self, "Ошибка", "Выбранный файл не является документом Word (.doc или .docx)")
                return 
            self.ui.label_2.setText(self.file_choice)


    
    def databased(self):
        self.file_dialog = QFileDialog(self, 'Выбрать файл', 'C://')
        self.file_dialog.setFileMode(QFileDialog.ExistingFile)
        self.file_dialog.setNameFilter("Microsoft Excel (*.xlsx *.xls)")
        if self.file_dialog.exec_() == QFileDialog.Accepted:
            self.file_choice = self.file_dialog.selectedFiles()[0]
            if QFileInfo(self.file_choice).suffix() not in ['xlsx', 'xls']:
                QMessageBox.warning(self, "Ошибка", "Выбранный файл не является документом Excel (.xlsx или .xls)")
                return 
            self.ui.label_3.setText(self.file_choice)

    def delete_label_text_diploma(self):
        self.ui.label_2.setText('')

    def delete_label_text_databased(self):
        self.ui.label_3.setText('')

    def get_url_to_diploma(self):
        if not self.ui.label_2.text() or not self.ui.label_3.text():
            QMessageBox.warning(self, "Ошибка", "Не выбраны файлы")
            return '', ''
        return  self.ui.label_2.text(), self.ui.label_3.text()

    

    def create_data(self,item: list) -> dict:  # Создание словаря с метками и их значениями
        self.full_name = f'{item[0]} {item[1]} {item[2]}'
        self.place = item[3]
        self.full_name_dictionary_and_place = {
                            '{{full_name}}': self.full_name,
                            '{{place}}': self.place,     
                                        }
        return self.full_name_dictionary_and_place


    def add_labels_to_diploma_more(self):
            #Считаваю данные из label'ов и передаю их в переменные в виде строки
            self.file_choice_diploma, self.file_choice_data = self.get_url_to_diploma()
            # Открытие выбранного файла с участниками и местами
            self.df = pd.read_excel(self.file_choice_data)
            self.data_list = self.df.values.tolist()  # преобразование всего датафрейма в список списков
            #Прохожу по всем вложенным спискам в списке
            for self.i, self.item in enumerate(self.data_list):
                # Создание словаря с метками и их значениями
                self.dictor = self.create_data(self.item)
                # Открытие выбранного документа 
                self.doc = docx.Document(self.file_choice_diploma)
                # Установка стилей документа
                self.style = self.doc.styles['Normal']
                self.font = self.style.font
                self.font.size = docx.shared.Pt(14)
                self.font.name = 'Times New Roman'
                # Замена меток на значения в документе

                for self.paragraph in self.doc.paragraphs:
                    self.new_text_diploma, self.count_entry_diploma = re.subn('|'.join(self.dictor.keys()), lambda match: self.dictor[match.group()], self.paragraph.text)
                if self.count_entry_diploma > 0:
                    self.paragraph.text = self.new_text_diploma

                # Сохранение документа с ФИО участника в качестве имени файла
                self.name_file_diploma = ' '.join(self.item[:4])
                self.doc.save(f'{self.name_file_diploma} место.docx')

    def create_mail(self, item):
        self.mail = item[4]
        return self.mail

    def send_diploma_to_email(self):
        #Считаваю данные из label'ов и передаю их в переменные в виде строки
        file_choice_diploma, file_choice_data = self.get_url_to_diploma()
        # Открытие выбранного файла с участниками и местами
        df = pd.read_excel(file_choice_data)
        data_list = df.values.tolist()  # преобразование всего датафрейма в список списков

        messages = []

        for i, item in enumerate(data_list):
            # Создание словаря с метками и их значениями
            dictor = self.create_data(item)
            # Открытие выбранного документа 
            doc = docx.Document(file_choice_diploma)
            # Установка стилей документа
            style = doc.styles['Normal']
            font = style.font
            font.size = docx.shared.Pt(14)
            font.name = 'Times New Roman'
            # Замена меток на значения в документе

            paragraph = None
            count_entry_diploma = 0
            new_text_diploma = None

            for paragraph in doc.paragraphs:
                new_text_diploma, count_entry_diploma = re.subn('|'.join(dictor.keys()), lambda match: dictor[match.group()], paragraph.text)
            if count_entry_diploma > 0:
                paragraph.text = new_text_diploma

            # Сохранение документа с ФИО участника в качестве имени файла
            name_file_diploma = ' '.join(item[:4])
            doc.save(f'{name_file_diploma} место.docx')

            mail_to_deliver = self.create_mail(item)
            subject = 'Диплом'
            text = 'Вы выиграли в олимпиаде! Диплом во вложении письма'
            files = [f'{name_file_diploma} место.docx']
            # Параметры соединения с SMTP


            msg = MIMEMultipart()
            msg['From'] = settings.mail_login
            msg['To'] = mail_to_deliver
            msg['Subject'] = subject
            # тут текст из text
            msg.attach(MIMEText(text))
            # Тут добавляем файл ворд в письмо
            if files:
                for file in files:
                    with open(file, 'rb') as file_user:
                        attach = MIMEApplication(file_user.read(), _subtype='docx')
                        attach.add_header('Content-Disposition', 'attachment', filename=file)
                        msg.attach(attach)

            messages.append(msg)

        self.__mail.send_message(messages)
    

