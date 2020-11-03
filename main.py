# Полуавтоматическая генерация индивидуального плана преподавателя
# motor1n develop PyQt5 - 2020 year


import sys
import xlrd
import time
import datetime as dt
from PyQt5 import uic
from docxtpl import DocxTemplate
from PyQt5.QtWidgets import (QApplication, QMainWindow, QFileDialog,
                             QTreeWidgetItemIterator, QTableWidgetItem,
                             QComboBox, QProgressBar)


# Текущий год:
CURRENT_YEAR = int(dt.datetime.today().strftime('%Y'))
# Текущий месяц:
CURRENT_MONTH = int(dt.datetime.today().strftime('%m'))
# Если пользователь воспользовался приложением во втором семестре,
# то уменьшим CURRENT_YEAR на единицу:
if CURRENT_MONTH in range(1, 7):
    CURRENT_YEAR -= 1
# Заместитель директора по учебной и научной работе:
DEPUTY = 'О.А. Тарасова'
# Педагогическая нагрузка на 1 ставку (часов):
RATE = 1524


class PlanForm(QMainWindow):
    def __init__(self):
        super().__init__()
        uic.loadUi('iplan-design.ui', self)
        # Заполняем выбор учебного года (плюс-минус год)
        self.cb5.addItems([f'{CURRENT_YEAR - 1} - {CURRENT_YEAR}',
                           f'{CURRENT_YEAR} - {CURRENT_YEAR + 1}',
                           f'{CURRENT_YEAR + 1} - {CURRENT_YEAR + 2}'])
        # По-умолчанию задаём среднее значение,
        # как наиболее вероятное при заполнении индивидуального плана
        self.cb5.setCurrentIndex(1)
        # Задаём дату на 1,5 года ранее текущей даты для избрвния по конкурсу на QDateEdit:
        self.de.setDate(dt.date.today() - dt.timedelta(days=548))
        # Включаем возможность выбирать дату при помощи PopUp-календаря:
        self.de.setCalendarPopup(True)
        # Сигнал --> слот: pb1 --> pb_open
        self.pb1.clicked.connect(self.pb_open)
        # Сигнал --> слот: pb2 --> pb_start
        #self.pb2.clicked.connect(self.pb_start)
        # Сигнал --> слот: pb3 --> pb_save
        self.pb3.clicked.connect(lambda checked, tab=self.tw1: self.pb_save(tab))
        # Сигнал --> слот: pb01 --> pb_load
        self.pb01.clicked.connect(lambda checked, tree=self.tree1, tab=self.tw1: self.load(tree, tab))

    def pb_open(self):
        fname = QFileDialog.getOpenFileName(self, 'Выбрать файл', '',
                                            'Excel 2007–365 (.xlsx)(*.xlsx)')[0]
        workbook = xlrd.open_workbook(fname)
        sh = workbook.sheet_by_index(0)
        # Учебная работа (вся):
        self.learn_work = round(sh.cell(sh.nrows - 3, 1).value)
        # Почасовая оплата:
        self.hourly_pay = round(sh.cell(sh.nrows - 3, 2).value)
        # Учебная работа (ставка):
        self.learn_rate = self.learn_work - self.hourly_pay
        # Внеучебная работа:
        self.extra_work = round(sh.cell(sh.nrows - 1, 1).value)
        # Словарь для дополнения в context и дальнейшего внесения данных в шаблон документа:
        self.up1 = dict()
        # Словарь соответствия столбцов исходной таблицы столбцам итоговой
        s = {1: 'm', 2: 'a', 3: 'b', 5: 'd', 6: 'c', 7: 'e', 9: 'f', 14: 'g', 16: 'h'}
        m1 = 1  # Счётчик отфильтрованных по нечётному семестру строк
        m2 = 1  # Счётчик отфильтрованных по чётному семестру строк
        # Суммы часов за семестры по раздичным видам учебной работы:
        as1, as2, as3, as4 = 0, 0, 0, 0  # Осенний семестр
        cs1, cs2, cs3, cs4 = 0, 0, 0, 0  # Весенний семестр
        # Пробегаем по строкам таблицы с нужными данными
        for rownum in range(1, sh.nrows - 3):
            # Считываем строку из исходной таблицы
            row = sh.row_values(rownum)
            # Считаем, что данные в строку нечётного семестра не вносились
            ok1 = False
            # Считаем, что данные в строку чётного семестра не вносились
            ok2 = False
            # Заходим в цикл по элементам строки
            for col in range(len(row)):
                # Округляем все вещественные до целых,
                # строки оставляем без изменения
                if row[col].__class__.__name__ == 'float':
                    tmp = round(row[col])
                else:
                    tmp = row[col]
                if col in s.keys():
                    if int(row[4]) % 2 == 1:  # Нечётный семестр
                        # Вносим данные в строку
                        self.up1[f'a{s[col]}0{m1}'] = tmp
                        if col == 7:
                            as1 += tmp  # Сумма за семестр: лекции
                        elif col == 9:
                            as2 += tmp  # Сумма за семестр: практика
                        elif col == 14:
                            as3 += tmp  # Сумма за семестр: экзамены
                        elif col == 16:
                            as4 += tmp  # Сумма за семестр: зачёты
                        ok1 = True  # В строку внесены данные
                    elif int(row[4]) % 2 == 0:  # Чётный семестр
                        # Вносим данные в строку
                        self.up1[f'c{s[col]}0{m2}'] = tmp
                        if col == 7:
                            cs1 += tmp  # Сумма за семестр: лекции
                        elif col == 9:
                            cs2 += tmp  # Сумма за семестр: практика
                        elif col == 14:
                            cs3 += tmp  # Сумма за семестр: экзамены
                        elif col == 16:
                            cs4 += tmp  # Сумма за семестр: зачёты
                        ok2 = True  # В строку внесены данные
            # Если строка была внесена, то переключаемся на следующую
            if ok1:
                m1 += 1
            if ok2:
                m2 += 1
        # Всего по плану за осенний семестр
        self.up1['as1'] = as1
        self.up1['as2'] = as2
        self.up1['as3'] = as3
        self.up1['as4'] = as4
        self.up1['lrnAP'] = as1 + as2 + as3 + as4
        # Всего по плану за весенний семестр
        self.up1['cs1'] = cs1
        self.up1['cs2'] = cs2
        self.up1['cs3'] = cs3
        self.up1['cs4'] = cs4
        self.up1['lrnSP'] = cs1 + cs2 + cs3 + cs4
        # План на год
        self.up1['dy1'] = tmp1 = as1 + cs1
        self.up1['dy2'] = tmp2 = as2 + cs2
        self.up1['dy3'] = tmp3 = as3 + cs3
        self.up1['dy4'] = tmp4 = as4 + cs4
        self.up1['lrnYP'] = tmp1 + tmp2 + tmp3 + tmp4

    def load(self, tree, tab):
        """Загрузка отмеченных элементов QTreeWidgetItem в таблицу QTableWidget"""
        checklist = list()  # Список для найденых выделенных check
        iter = QTreeWidgetItemIterator(tree, QTreeWidgetItemIterator.Checked)
        while iter.value():
            currentItem = iter.value()
            #print(currentItem.text(0), currentItem.text(1))
            checklist.append((currentItem.text(0), currentItem.text(1)))
            iter += 1
        #print(checklist)
        # Если ничего не выбрано,
        # то выведем сообщение об этом в статус-бар и вернём пустой return
        if not checklist:
            self.statusBar().showMessage('Внеучебная работа: ничего не выбрано')
            return
        else:
            tab.setColumnWidth(0, 500)  # Задаём размер первого столбца пошире
            tab.clearContents()  # Очищаем содержимое таблицы
            # Заполняем QTableWidget данными из QTreeWidget
            for i, elem in enumerate(checklist):
                for j, val in enumerate(elem):
                    tab.setItem(i, j, QTableWidgetItem(val))
            self.statusBar().showMessage(f'Таблица сформирована')
        # Помещаем кнопки QComboBox в поле "Срок выполнения"
        # на таблицу QTableWidget
        i = 0
        for j in range(len(checklist)):
            # Создаём объект QComboBox
            cbox = QComboBox()
            # Задаём его содержимое
            cbox.addItems(('сентябрь', 'октябрь', 'ноябрь', 'декабрь', 'январь',
                           'февраль', 'март', 'апрель', 'май', 'июнь',
                           '1 семестр', '2 семестр', 'в течение года'))
            # Помещаем cbox в ячейку таблицы
            tab.setCellWidget(i, 3, cbox)
            i += 1

    def pb_save(self, tab):
        """Получение данных из интерфейса и запись в файл индивидуального плана"""
        # Фамилия Имя Отчество:
        name = self.cb1.currentText()
        # Учёная степень, звание:
        degree = self.cb2.currentText()
        # Должность:
        position = self.cb3.currentText()
        # Кафедра:
        cathedra = self.cb4.currentText()
        # Дата избрания по конкурсу:
        election = self.de.dateTime().toString('dd.MM.yyyy')
        # Учебный год (первый семестр):
        self.current_year = int(self.cb5.currentText()[:4])
        # Словарь для дополнения в context и дальнейшего внесения данных в шаблон документа
        self.up01 = dict()
        # Словарь соответствия столбцов исходной таблицы столбцам итоговой
        s = {0: 'A', 1: 'B', 2: 'C', 3: 'D', 4: 'E', 5: 'F'}
        # Получение списка элементов заголовка таблицы:
        # lst = [tab.horizontalHeaderItem(i).text() for i in range(tab.columnCount())]
        # Запись остальных строк в файл:
        summ = 0
        for i in range(tab.rowCount()):
            #row = []
            for j in range(tab.columnCount()):
                #itm = tab.item(i, j)
                if tab.item(i, j) is not None:
                    #row.append(itm.text())
                    #self.up01[f'lrn{s[0]}0{i + 1}'] = itm.text()
                    # Поле: Виды работы
                    self.up01[f'mtd{s[0]}0{i + 1}'] = tab.item(i, 0).text()
                    # Поле: Трудоёмкость в часах
                    labour = tab.item(i, 1).text().split()[0]
                    self.up01[f'mtd{s[2]}0{i + 1}'] = labour
                    # Поле: Срок выполнения
                    # tab.cellWidget(i, 3).currentText() - смотрим содержимое
                    # ячеек поля "Срок выполнения" - это названия месяцев
                    self.up01[f'mtd{s[3]}0{i + 1}'] = tab.cellWidget(i, 3).currentText()
                    # Поле: Запланировано
                    itm = tab.item(i, 4).text()
                    planned = f'{int(int(itm) / float(labour))}\u2219{labour}={itm}'
                    self.up01[f'mtd{s[4]}0{i + 1}'] = planned
                    summ += int(itm)
                    break
        # Записываем общую сумму в документ docx
        self.up01[f'mtdYP'] = summ

        # Файл шаблона:
        doc = DocxTemplate('iplan-template.dotx')
        context = {
            'cathedra': cathedra,
            'deputy': DEPUTY,
            'name': name,
            'degree': degree,
            'position': position,
            'election': election,
            'year_one': str(self.current_year),
            'year_two': str(self.current_year + 1),
            'hourly': self.hourly_pay,
            'user_rate': self.learn_rate
        }
        # Объединяем все данные для внесения в документ
        context.update(self.up1)
        context.update(self.up01)
        fname = QFileDialog.getSaveFileName(self, 'Сохранить документ', '',
                                            'Word 2007–365 (.docx)(*.docx)')[0]
        doc.render(context)
        doc.save(fname)
        self.statusBar().showMessage('Идёт формирование документа...')
        TIME_LIMIT = 100
        count = 0
        while count < TIME_LIMIT:
            count += 1
            time.sleep(0.1)
            self.progress.setValue(count)
        self.statusBar().showMessage(f'Документ {fname} сформирован')
        self.progress.setValue(0)


# Отслеживаем ошибки PyQt5
def except_hook(cls, exception, traceback):
    sys.__excepthook__(cls, exception, traceback)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = PlanForm()
    ex.show()
    sys.excepthook = except_hook
    sys.exit(app.exec_())
