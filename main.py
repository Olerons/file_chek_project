import sys, os
import openpyxl as ox
import win32com.client as win32

from QT import Ui_MainWindow
from check_base import DataBase, DataBaseTable
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QWidget, QTableWidgetItem


class MainPustograf(QMainWindow, Ui_MainWindow):
    """
    main-класс "Анализа пустографок"
    Собирает интерфейс и подключает кнопки
        
        switch_interface - Позволяет переключаться между интерфейсами прграммы
        open_folder -      Выбор папки с файлами через диалогове окно
        exam -             Создание необходимых корневых директорий (/Base/номер школы)
                           Перебор файлов каталога. Отсеивание не excel файлов
        convert -          Преобразование xls в xlsx файл (РАБОТАЕТ ТОЛЬКО С УСТАНОВЛЕННЫМ OFFICE)
        exam_process -     Процесс анализа найденных таблиц
        exam_value_f -     Вспомогательная функция для анализа по определенным ключам
        open_table -       Окно с таблицей данных из БД
    """
    def __init__(self):
        self.cwd = os.getcwd()
        self.folder_name = ''
        self.dbase = DataBase()
        super().__init__()
        self.setupUi(self)
        
        self.toolButton.clicked.connect(self.open_table)
        self.esc_btn.clicked.connect(self.switch_interface)
        self.start_btn.clicked.connect(self.switch_interface)
        self.folder_btn.clicked.connect(self.open_folder)
        self.id_dict = {'id1' : 'выпуск 9 класса',
                        'id2' : 'выпуск 11 класса',
                        'id3' : 'классные рук',
                        'id4' : 'инд. обучение',
                        'id5' : 'база многодетных',
                        'id6' : 'социальный паспорт',
                        'id7' : 'База по ПМПК',
                        'id8' : 'База по ин яз',
                        'id9' : 'дети-инвалиды',
                        'id10' : 'программы нач. шк.'}


    def switch_interface(self):
        """
        Переход на второй интерфейс по нажатию кнопки start_btn, если выбрана папка, названа школа
        Переход на первый интерфейс по нажатию кнопки esc_btn, если процесс достиг 100%
        """
        if self.sender() == self.start_btn and self.folder_name and self.school_lineEdit.text(): 
            self.main_widget.hide()
            self.main_widget2.show()
            self.exam()
        elif self.sender() == self.esc_btn and self.progressBar.value() == 100:
            self.school_lineEdit.setText('')
            self.folder_text.setText('')
            self.main_widget2.hide()
            self.main_widget.show()


    def open_folder(self):
        """ Создание диалога для выбора папки с файлами, заполнение текстового поля """
        self.folder_name = QFileDialog.getExistingDirectory(self, 'Выберите папку', self.cwd)
        self.folder_text.setText(self.folder_name)


    def exam(self):
        """
        Создание в корневой папке программы папки Base
        Создание в папке Base папки с наименованием школы
        Перебор файлов пользователя.
            принудительная конвертация xls файлов в xlsx
            запуск анализа
        """
        if not os.path.isdir('Base'):
            os.mkdir('Base')
        if not os.path.isdir(self.cwd + '/Base/' + self.school_lineEdit.text()):
            os.mkdir('Base/' + self.school_lineEdit.text())
        self.save_path = f'{self.cwd}/Base/{self.school_lineEdit.text()}/'
        
        self.progress = 0
        self.up_data_db = []
        
        for file in os.listdir(self.folder_name):
            if file.split('.')[-1] == 'xls':
                convert_file = self.convert(file)
                self.exam_process(convert_file)
            elif file.split('.')[-1] == 'xlsx':
                self.exam_process(file)
        
        for lbl in self.lbl_group:
            self.up_data_db.append(lbl.text())
        
        sch2 = ['2', '2 ВСОШ', 'ВСОШ 2', '2ВСОШ', 'ВСОШ2', 'ВСОШ', '2(ВСОШ)']
        gim1 = ['1Г', 'Г1', '1 Г', 'Г 1', '1 ГИМНАЗИЯ', 'ГИМНАЗИЯ 1', '1 ГИМ', 'ГИМ 1']
        
        if self.school_lineEdit.text().upper() in sch2:
            self.up_data_db.append('2(всош)')
        elif self.school_lineEdit.text().upper() in gim1:
            self.up_data_db.append('Г1')
        else:
            self.up_data_db.append(self.school_lineEdit.text())
        
        self.progressBar.setValue(100)
        self.dbase.updDB(tuple(self.up_data_db))
        

    def convert(self, filename):
        """
        Конвертация файла из xls в xlsx.
        !!! РАБОТАЕТ ТОЛЬКО ПРИ УСТАНОВЛЕННОМ EXCEL 2010 И НОВЕЕ !!!
        Это связано с шифровкой исходных файлов.
        """
        try:
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            # Преобразуем path в абсолютный путь для корректной работы
            wb = excel.Workbooks.Open(os.path.abspath(self.folder_name + '/' + filename))
            excel.DisplayAlerts = False # отключаем вопрос о перезаписи на всякий случай
            wb.SaveAs(os.path.abspath(self.folder_name + '/convert' + filename + 'x'), FileFormat=51)
            excel.DisplayAlerts = True # включаем обратно
            wb.Close()
            excel.Application.Quit()
            return self.folder_name + '/convert' + filename + 'x'
        
        except Exception as exc:
            print('ОШИБКА  ', exc)


    def exam_process(self, file):
        """
        Анализ найденых таблиц.
        По id таблицы устанавливаются поля для проверки данных
        """
        self.exam_key = '-'
        try:
            self.wb = ox.load_workbook(filename = os.path.abspath(self.folder_name + '/' + file))
            activsheet = self.wb.worksheets[0]
            file_id = activsheet['A1'].value
        except:
            file_id = ''
                
        if file_id == 'id1':
            exam_value = (['A3', 'B3', 'C3', 'O3'],
                        ['Наименование образовательной организации',
                         'Класс', 'Фамилия выпускника',
                         'Форма обучения (очная, очно-заочная, заочная)'])
            self.exam_key = self.exam_value_f(activsheet, exam_value, 'C4')
            self.ok1_lbl.setText(self.exam_key)
            
        elif file_id == 'id2':
            exam_value = [['A3', 'B3', 'C3', 'P3'],
                        ['Наименование образовательной организации',
                         'Класс', 'Фамилия выпускника',
                         'Педагогический класс']]
            self.exam_key = self.exam_value_f(activsheet, exam_value, 'C4')
            self.ok2_lbl.setText(self.exam_key)
            
        elif file_id == 'id3':
            exam_value = [['A3', 'B3', 'C3', 'H3'],
                          ['ОУ', 'Ф.И.О. классного руководителя полностью',
                           'Класс', '№ приказа по ОУ о дате назначения ']]
            self.exam_key = self.exam_value_f(activsheet, exam_value, 'B4')
            self.ok3_lbl.setText(self.exam_key)
            
        elif file_id == 'id4':
            exam_value = [['A2', 'B2', 'C2', 'I2'],
                          ['ОУ', 'Фамилия, имя учащегося',
                           'Класс', 'Общая учебная нагрузка  в час/нед']]
            self.exam_key = self.exam_value_f(activsheet, exam_value, 'B4')
            self.ok4_lbl.setText(self.exam_key)
            
        elif file_id == 'id5':
            exam_value = [['A3', 'B3', 'C3', 'H3'],
                          ['№ п/п семьи', '№ ОУ', 'ФИО отца',
                           'ФИО детей']]
            self.exam_key = self.exam_value_f(activsheet, exam_value, 'C4')
            self.ok5_lbl.setText(self.exam_key)
            
        elif file_id == 'id6':
            exam_value = [['A3', 'B3', 'C3', 'H3'],
                          ['Всего учащихся в ОУ', 'Полные семьи  (количество)',
                           'Неполные семьи  (количество)', 'Количество опекаемых детей']]
            self.exam_key = self.exam_value_f(activsheet, exam_value, 'B4')
            self.ok6_lbl.setText(self.exam_key)
            
        elif file_id == 'id7':
            exam_value = [['A3', 'B3', 'C3', 'E3'],
                          ['Класс', 'Количество учащихся', 'ФИО ребенка',
                           'Примечание ( особые отметки) \nУказать вариант программ обучения для учащихся 1 и 2 классов)']]
            self.exam_key = self.exam_value_f(activsheet, exam_value, 'C4')
            self.ok7_lbl.setText(self.exam_key)
            
        elif file_id == 'id8':
            exam_value = [['A2', 'B3', 'C3', 'E3'],
                          ['Класс', 'немецкий', 'английский',
                           'второй иностранный язык']]
            self.exam_key = self.exam_value_f(activsheet, exam_value, 'A6')
            self.ok8_lbl.setText(self.exam_key)
            
        elif file_id == 'id9':
            exam_value = [['A3', 'B3', 'C3', 'E3'],
                          ['ФИО', 'Дата рождения', 'Класс',
                           'Домашний адрес']]
            self.exam_key = self.exam_value_f(activsheet, exam_value, 'A4')
            self.ok9_lbl.setText(self.exam_key)
            
        elif file_id == 'id10':
            exam_value = [['A3', 'B3', 'C3', 'D3'],
                          ['Класс', 'Количество учащихся', 'ФИО учителя  (полностью)',
                           'Программа обучения']]
            self.exam_key = self.exam_value_f(activsheet, exam_value, 'C4')
            self.ok10_lbl.setText(self.exam_key)
            
        
    def exam_value_f(self, activsheet, exam_value, empty_value):
        """
        Анализ значений таблицы
        X - если шапка не соответствует шаблону
        Пусто - если под шапкой нет значений
        Образец - если не удален образец заполнения
        Ок - если все отлично
        """
        self.progress += 10
        self.progressBar.setValue(self.progress)
        for i in range(4):
            if activsheet[exam_value[0][i]].value != exam_value[1][i]:
                return 'X'
        if self.exam_key == '-' and activsheet[empty_value].value == None:
            return 'Пусто'
        elif self.exam_key == '-' and activsheet[empty_value].value == 'ОБРАЗЕЦ ОБРАЗЕЦ ОБРАЗЕЦ':
            return 'Образец'
        else:
            self.wb.save(f'{self.save_path}{self.id_dict[activsheet["A1"].value]}.xlsx')
            return 'Ok'


    def open_table(self):
        self.form = QWidget()
        dbtable = DataBaseTable(self.form)
        dbtable.esc_btn.clicked.connect(lambda : self.form.close())
        dbtable.clear_btn.clicked.connect(lambda : dbtable.clear())



if __name__ == '__main__':
    app = QApplication(sys.argv)
    m = MainPustograf()
    m.show()
    sys.exit(app.exec_())