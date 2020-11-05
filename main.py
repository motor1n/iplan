# iPlan 0.0.2
# Автоматическая генерация индивидуального плана преподавателя
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
# то для корректности уменьшим CURRENT_YEAR на единицу:
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
        # Учебная работа - сигнал pb_lrn --> слот learn
        self.pb_lrn.clicked.connect(self.learn)
        # Учебно-методическая работа - сигнал pb01 --> слот extra
        self.pb01.clicked.connect(lambda checked, tree=self.tree1,
                                         tab=self.tw1: self.extra(tree, tab))
        # Организационная работа - сигнал pb02 --> слот extra
        self.pb02.clicked.connect(lambda checked, tree=self.tree2,
                                         tab=self.tw2: self.extra(tree, tab))
        # Научно-исследовательская работа - сигнал pb03 --> слот extra
        self.pb03.clicked.connect(lambda checked, tree=self.tree3,
                                         tab=self.tw3: self.extra(tree, tab))
        # Воспитательная работа - сигнал pb04 --> слот extra
        self.pb04.clicked.connect(lambda checked, tree=self.tree4,
                                         tab=self.tw4: self.extra(tree, tab))
        # Повышение квалификации - сигнал pb05 --> слот extra
        self.pb05.clicked.connect(lambda checked, tree=self.tree5,
                                         tab=self.tw5: self.extra(tree, tab))
        # Сигнал pb3 --> слот savedocx
        self.pb_save.clicked.connect(lambda checked,
                                        tables=(self.tw1, self.tw2, self.tw3,
                                                self.tw4, self.tw5): self.savedocx(tables))

    def learn(self):
        """Заполнение данных по учебной работе"""
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

    def extra(self, tree, tab):
        """Заполнение данных по внеучебной работе"""
        # Загрузка отмеченных элементов QTreeWidgetItem в таблицу QTableWidget
        # Список для найденых выделенных check
        checklist = list()
        # Создаём итератор для прохода по элементам QTreeWidget
        iter = QTreeWidgetItemIterator(tree, QTreeWidgetItemIterator.Checked)
        while iter.value():
            # Читаем строку QTreeWidgetItem
            currentItem = iter.value()
            # Значение toolTip ячейки "Трудоёмкость"
            print('Трудоёмкость:', currentItem.toolTip(1))
            checklist.append((currentItem.text(0), currentItem.text(1), currentItem.text(2)))
            iter += 1
        # Если ничего не выбрано,
        # то выведем сообщение об этом в статус-бар и вернём пустой return
        if not checklist:
            self.statusBar().showMessage('Внеучебная работа: не выбрано')
            return
        else:
            # Задаём размер первого столбца пошире
            tab.setColumnWidth(0, 500)
            # Очищаем содержимое таблицы
            tab.clearContents()
            # Заполняем QTableWidget данными из QTreeWidget
            for i, elem in enumerate(checklist):
                for j, val in enumerate(elem):
                    tab.setItem(i, j, QTableWidgetItem(val))

                    #QTableWidgetItem(val).setToolTip('100')
                    #print(QTableWidgetItem(val).text())
                    print(QTableWidgetItem(val).toolTip())
                    #print(QTableWidgetItem(val).statusTip())
                    #print(QTableWidgetItem(val).whatsThis())
            msg = f'Внеучебная работа: выбрано позиций: {len(checklist)}'
            self.statusBar().showMessage(msg)
        # Помещаем кнопки QComboBox
        # в поле "Срок выполнения" на таблицу QTableWidget
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

    def savedocx(self, tables):
        """Сбор данных из интерфейса и запись в docx"""
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
        self.up2 = dict()
        # Словарь соответствия столбцов исходной таблицы столбцам итоговой
        d1 = {0: 'A', 1: 'B', 2: 'C', 3: 'D', 4: 'E', 5: 'F'}
        # Словарь соответствия номеров вкладок разделам внеучебной работы
        d2 = {0: 'mtd', 1: 'org', 2:'sci', 3: 'edu', 4: 'upg'}
        # Получение списка элементов заголовка таблицы:
        # lst = [tab.horizontalHeaderItem(i).text() for i in range(tab.columnCount())]
        # Запись данных о внеучебной работе из таблиц интерфейса в файл docx:
        for t in range(len(tables)):
            summ = 0
            for i in range(tables[t].rowCount()):
                for j in range(tables[t].columnCount()):
                    if tables[t].item(i, j) is not None:
                        # Поле: Виды работы
                        self.up2[f'{d2[t]}{d1[0]}0{i + 1}'] = tables[t].item(i, 0).text()
                        # Поле: Трудоёмкость в часах
                        labour = tables[t].item(i, 1).text().split()[0]
                        self.up2[f'{d2[t]}{d1[2]}0{i + 1}'] = labour
                        # Поле: Срок выполнения
                        # tab.cellWidget(i, 3).currentText() - смотрим содержимое
                        # ячеек поля "Срок выполнения" - названия месяцев учебного года
                        self.up2[f'{d2[t]}{d1[3]}0{i + 1}'] = tables[t].cellWidget(i, 3).currentText()
                        # Поле: Запланировано
                        itm = tables[t].item(i, 4).text()
                        # Проверяем - первые символы в трудоёмкости цифры?
                        if labour.isnumeric():
                            planned = f'{int(int(itm) / float(labour))}\u2219{labour}={itm}'
                            self.up2[f'{d2[t]}{d1[4]}0{i + 1}'] = planned
                        summ += int(itm)
                        break
            # Записываем общую сумму по разделу внеучебной работы в документ docx
            self.up2[f'{d2[t]}YP'] = summ

        # Файл шаблона:
        doc = DocxTemplate('iplan-template.dotx')
        # Начальный словарь для рендеринга
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
        # Объединяем данные учебной (up1) и внеучебной (up2) работы
        # в словарь рендеринга для дальнейшего внесения в документ
        context.update(self.up1)
        context.update(self.up2)
        # Диалоговое окно сохранения файла docx
        fname = QFileDialog.getSaveFileName(self, 'Сохранить документ', '',
                                            'Word 2007–365 (.docx)(*.docx)')[0]
        self.statusBar().showMessage('Идёт формирование документа...')
        # Рендерим инфу в шаблон
        doc.render(context)
        # Сохраняем конечный документ
        doc.save(fname)
        # Используем progressBar
        TIME_LIMIT = 100
        count = 0
        while count < TIME_LIMIT:
            count += 1
            time.sleep(0.1)
            self.progress.setValue(count)
        # Выводим сообщение в статус-бар
        self.statusBar().showMessage(f'Документ {fname} сформирован')
        # Обнуляем progressBar после операции
        self.progress.setValue(0)


def except_hook(cls, exception, traceback):
    """Функция для отслеживания ошибок PyQt5"""
    sys.__excepthook__(cls, exception, traceback)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = PlanForm()
    ex.show()
    # Ловим и показываем ошибки PyQt5 в терминале
    sys.excepthook = except_hook
    sys.exit(app.exec_())
