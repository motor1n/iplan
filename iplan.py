#!/usr/bin/env python3
# -*- coding: utf-8 -*-

# iPlan 0.1.5.3
# Автоматическая генерация индивидуального плана преподавателя
# developed on PyQt5

__author__ = 'Vladislav Motorin <motorin@yandexlyceum.ru>'


import sys
import xlrd
import datetime as dt
from PyQt5 import uic
from docxtpl import DocxTemplate
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtWidgets import (QApplication, QMainWindow, QFileDialog,
                             QTreeWidgetItemIterator, QTableWidgetItem,
                             QComboBox, QMessageBox, QProgressDialog)


# Текущий год:
CURRENT_YEAR = int(dt.datetime.today().strftime('%Y'))
# Текущий месяц:
CURRENT_MONTH = int(dt.datetime.today().strftime('%m'))
# Если вдруг преподаватель воспользовался приложением во втором семестре,
# то для корректности уменьшим CURRENT_YEAR на единицу:
if CURRENT_MONTH in range(1, 7):
    CURRENT_YEAR -= 1
# Заместитель директора по учебной и научной работе:
DEPUTY = 'О.А. Тарасова'
# Педагогическая нагрузка на 1 ставку (часов):
RATE = 1524
# Виды внеучебной работы:
TYPES_EXTRAWORK = {
    0: 'Учебно-методическая работа',
    1: 'Организационная работа',
    2: 'Научно-исследовательская',
    3: 'Воспитательная работа',
    4: 'Повышение квалификации'
}


class Thread1(QThread):
    """Поток для рендеринга и сохранения файла"""
    # Создаём собственный сигнал,
    # принимающий параметр типа str:
    signal = pyqtSignal(str)

    def __init__(self, fname, context, parent=None):
        """Инициализация потока"""
        # fname - имя сохраняемого файла
        # content - дянные для рендеринга
        QThread.__init__(self, parent)
        # Проверяем, имеет ли правильное
        # расширение сохраняемый файл:
        if fname.endswith('.docx'):
            self.fname = fname
        else:
            self.fname = f'{fname}.docx'
        self.context = context

    def run(self):
        """Рендеринг и сохраниение докумнета"""
        # Подключаем файл шаблона .dotx:
        doc = DocxTemplate('iplan-template.dotx')
        # Рендерим инфу в шаблон
        self.signal.emit('Рендерим инфу в шаблон')
        doc.render(self.context)
        # Сохраняем конечный документ
        self.signal.emit('Сохраняем конечный документ')
        try:
            doc.save(self.fname)
        except Exception:
            self.signal.emit('error')


class PlanForm(QMainWindow):
    def __init__(self):
        super().__init__()
        uic.loadUi('iplan-design.ui', self)
        # Файл нагрузки ещё не открыт:
        self.fileopen = False
        # Ошибок открытия файла ещё не было:
        self.errorOpen = False
        # Кнопка "Открыть..." дезактивирована,
        # сначала пользователь должен ввести свои данные:
        self.pb_lrn.setDisabled(True)
        # Интерфейс внеучебной работы дезактивирован изначально,
        # чтобы пользователь сначала заполнил учебную работу:
        self.tabs.setDisabled(True)
        # Кнопка pb_save также дезактивирована,
        # поскольку на данный момент ещё нечего сохранять:
        self.pb_save.setDisabled(True)
        # Таблицы QTableWidget ещё не заполнениы:
        self.complete_tabs = False
        # Словарь - состояние заполненности таблиц QTableWidget:
        self.condition_tabs = dict()
        # Кортеж кнопок QComboBox на заполнение данных пользователя:
        self.cbX = (self.cb1, self.cb2, self.cb3, self.cb4)
        # Кортеж объектов QTreeWidget
        self.treeX = (self.tree1, self.tree2,
                      self.tree3, self.tree4, self.tree5)
        # Экспандинг древовидной структуры объектов QTreeWidget:
        for tree in self.treeX:
            tree.expandAll()
        # Кортеж кнопок "Записать в таблицу"
        self.pb0X = (self.pb01, self.pb02, self.pb03, self.pb04, self.pb05)
        # Кортеж таблиц интерфейса для сохранения в документ
        self.tables = (self.tw1, self.tw2, self.tw3, self.tw4, self.tw5)
        # Заполняем выбор учебного года (плюс-минус год)
        self.cb5.addItems([f'{CURRENT_YEAR - 1} - {CURRENT_YEAR}',
                           f'{CURRENT_YEAR} - {CURRENT_YEAR + 1}',
                           f'{CURRENT_YEAR + 1} - {CURRENT_YEAR + 2}'])
        # По-умолчанию задаём среднее значение,
        # как наиболее вероятное при заполнении индивидуального плана
        self.cb5.setCurrentIndex(1)
        # Задаём дату на 1,5 года ранее текущей даты для избрания по конкурсу на QDateEdit:
        self.de.setDate(dt.date.today() - dt.timedelta(days=548))
        # Включаем возможность выбирать дату при помощи PopUp-календаря:
        self.de.setCalendarPopup(True)
        # Учебная работа - сигнал pb_lrn --> слот learn
        self.pb_lrn.clicked.connect(self.learn)
        # Связываем сигналы от кнопок "Записать в таблицу" с сответствующими слотами:
        for i in range(len(self.pb0X)):
            self.pb0X[i].clicked.connect(lambda checked,
                                                tree=self.treeX[i],
                                                tab=self.tables[i]: self.extra(tree, tab))
        # Сигнал pb3 --> слот savedocx
        self.pb_save.clicked.connect(lambda checked,
                                            tables=self.tables: self.savedocx(tables))
        # Сигналы отслеживания изменений таблиц QTableWidget (кортеж tables),
        # но без "Повышения квалификации", поскольку она не обязательна.
        for tab in self.tables[:-1]:
            tab.cellChanged.connect(self.complete_alltabs)
            # И попутно заполнение словаря self.condition_tabs значениями False,
            # т.е. пока ещё ни одна таблица не заполнена полностью.
            self.condition_tabs[tab.objectName()] = False
        # Сигналы отслеживания изменений QComboBox на заполнение данных пользователя:
        for cb in self.cbX:
            cb.currentTextChanged.connect(self.user)
        # Сигнал вабранной вкладки QTabWidget:
        self.tabs.tabBarClicked.connect(self.show_currtab_name)
        msg = 'Изучите порядок заполнения индивидуального плана и приступайте к работе'
        self.statusBar().showMessage(msg)

    def user(self):
        """Контроль заполнения данных о пользователе"""
        if '---' not in [cb.currentText() for cb in self.cbX]:
            # Если данные заполнены, активируем кнопку "Открыть..."
            self.pb_lrn.setDisabled(False)
            msg = 'Проверьте дату избрания по конкурсу, ' \
                  'текущий учебный год и откройте ваш файл учебной нагрузки.'
            self.statusBar().showMessage(msg)

    def learn(self):
        """Заполнение данных по учебной работе"""
        if self.fileopen:
            # noinspection PyArgumentList
            msg = QMessageBox.information(self, 'Инфо', '<h4>Файл уже был открыт,'
                                                        '<br>но можно выбрать другой.</h4>')
            fname, _ = QFileDialog.getOpenFileName(self,
                                                   'Выбрать файл', None,
                                                   'Microsoft Excel 2007–365 (*.xlsx)')
        else:
            fname, _ = QFileDialog.getOpenFileName(self,
                                                   'Выбрать файл', None,
                                                   'Microsoft Excel 2007–365 (*.xlsx)')
        try:
            workbook = xlrd.open_workbook(fname)
            # Читаем первый лист:
            sh = workbook.sheet_by_index(0)
            # Учебная работа (вся):
            self.learn_work = round(sh.cell(sh.nrows - 3, 1).value)
            # Почасовая оплата:
            self.hourly_pay = round(sh.cell(sh.nrows - 3, 2).value)
            # Учебная работа (ставка):
            self.learn_rate = self.learn_work - self.hourly_pay
            # Внеучебная работа:
            self.extra_work = round(sh.cell(sh.nrows - 1, 1).value)
            # Расчёт долей процентов: учебная работа
            self.percent_user_rate = round(self.learn_rate /
                                           (RATE * sh.cell(sh.nrows - 3, 0).value) * 100, 1)
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
            # Флаг: файл открыт
            self.fileopen = True
            # Сообщение: файл учебной нагрузки открыт
            # noinspection PyArgumentList
            msg = QMessageBox.information(self, 'Инфо',
                                          '<h4>Файл учебной нагрузки открыт,'
                                          '<br>можно продолжить работу.</h4>')
            self.statusBar().showMessage('Заполните данные по внеучебной работе')
            # Учебная работа заполнена,
            # делаем активной для заполнеия внеучебную работу
            self.tabs.setDisabled(False)
        except FileNotFoundError:
            if self.fileopen and not self.errorOpen:
                message = '<h4>Вы уже открыли файл</h4>'
            else:
                message = '<h4>Вы не открыли файл,<br>попробуйте ещё раз.</h4>'
            msg = QMessageBox.information(self, 'Инфо', message)
            self.errorOpen = True

    def extra(self, tree, tab):
        """Заполнение данных по внеучебной работе"""
        # Загрузка отмеченных элементов QTreeWidgetItem в таблицу QTableWidget
        # Список для найденых выделенных check
        self.checklist = list()
        # Создаём итератор для прохода по элементам QTreeWidget
        it = QTreeWidgetItemIterator(tree, QTreeWidgetItemIterator.Checked)
        while it.value():
            # Читаем строку QTreeWidgetItem
            current_item = it.value()
            # currentItem.text(0) - текст в ячейке "Виды работы"
            # currentItem.toolTip(1) - всплывающая подсказка ячейки "Трудоёмкость"
            # currentItem.text(2) - текст в ячейке "Форма отчётности"
            if tab.objectName() == 'tw3':
                # Если таблица "Научная работа", то дополняется столбец "Объём п.л. или стр."
                self.checklist.append((current_item.text(0), current_item.toolTip(1),
                                       current_item.text(2), current_item.text(3)))
            else:
                self.checklist.append((current_item.text(0), current_item.toolTip(1),
                                       current_item.text(2)))
            it += 1
        # Если ничего не выбрано,
        # то выведем сообщение об этом в статус-бар и вернём пустой return
        if not self.checklist:
            self.statusBar().showMessage(f'{self.get_currtab_name()}: ничего не выбрано')
            tab.clearContents()
            return
        else:
            # Задаём размер первого столбца пошире
            tab.setColumnWidth(0, 500)
            # Очищаем содержимое таблицы
            tab.clearContents()
            # Заполняем QTableWidget данными из QTreeWidget
            for i, elem in enumerate(self.checklist):
                for j, val in enumerate(elem):
                    tab.setItem(i, j, QTableWidgetItem(val))
        # Выводим текущую инфу в статус бар
        msg = f'{self.get_currtab_name()} | Выбрано позиций: {len(self.checklist)}'
        self.statusBar().showMessage(msg)
        # Помещаем кнопки QComboBox в поле "Срок выполнения" на таблицу QTableWidget
        i = 0
        for j in range(len(self.checklist)):
            # Создаём объект QComboBox
            cbox = QComboBox()
            # Задаём его содержимое
            cbox.addItems(('сентябрь', 'октябрь', 'ноябрь', 'декабрь', 'январь',
                           'февраль', 'март', 'апрель', 'май', 'июнь',
                           '1 семестр', '2 семестр', 'в течение года'))
            # Помещаем cbox в ячейку таблицы
            if tab == self.tw3:
                tab.setCellWidget(i, 4, cbox)
            else:
                tab.setCellWidget(i, 3, cbox)
            i += 1

    def get_currtab_name(self):
        """Название текущей внеучебной работы"""
        current_tab = TYPES_EXTRAWORK[self.tabs.currentIndex()]
        return current_tab

    def show_currtab_name(self, curr_tab):
        """Выводим имя текушей внеучебной работы в статус-бар"""
        self.statusBar().showMessage(TYPES_EXTRAWORK[curr_tab])

    def count_fill_rows(self, tab):
        """Количество заполненых строк в таблице"""
        count = 0
        for i in range(tab.rowCount()):
            for j in range(tab.columnCount()):
                if tab.item(i, j):
                    count += 1
                    break
        return count

    def is_tabfull(self, tab):
        """Проверка заполненности одной таблицы QTableWidget"""
        for i in range(self.count_fill_rows(tab)):
            for j in range(tab.columnCount()):
                # Проверяем все ячейки кроме столбца "Срок выполнения"
                if tab == self.tw3:
                    if j != 4 and (tab.item(i, j) is None
                                   or tab.item(i, j).text() == str()):
                        return False
                else:
                    if j != 3 and (tab.item(i, j) is None
                                   or tab.item(i, j).text() == str()):
                        return False
        return True

    def complete_alltabs(self, *args):
        """Проверка заполненности таблиц QTableWidget"""
        # Проверяем все, кроме "Повышения квалификации"
        for tab in self.tables[:-1]:
            if self.is_tabfull(tab):
                self.condition_tabs[tab.objectName()] = True
            else:
                self.condition_tabs[tab.objectName()] = False
        # Если всё заполнено, активируем кнопку "Сохранить..."
        if all(self.condition_tabs.values()):
            self.complete_tabs = True
            self.pb_save.setDisabled(False)
            self.statusBar().showMessage('Сохраните документ')
        else:
            self.complete_tabs = False
            self.pb_save.setDisabled(True)

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
        d2 = {0: 'mtd', 1: 'org', 2: 'sci', 3: 'edu', 4: 'upg'}
        # Строка "Всего" в таблице "Распределение времени по семестрам и основным видам работы,
        # общие суммы по плану внеучебной работы (осенний, весенний, год)
        AP = 0
        SP = 0
        YP = 0
        # Список для записи часов по внеучебной работе на титульный лист:
        perext = list()
        try:
            # Запись данных о внеучебной работе из таблиц интерфейса в файл docx:
            for t in range(len(tables)):
                # Общая сумма по разделу внеучебной работы
                summ = 0
                # Осенний семестр
                autumn = 0
                # Весенний семестр
                spring = 0
                for i in range(tables[t].rowCount()):
                    for j in range(tables[t].columnCount()):
                        if tables[t].item(i, j) is not None:
                            # Поле: Виды работы
                            self.up2[f'{d2[t]}{d1[0]}0{i + 1}'] = tables[t].item(i, 0).text()
                            # Поле: Форма отчётности
                            self.up2[f'{d2[t]}{d1[1]}0{i + 1}'] = tables[t].item(i, 2).text()
                            if tables[t] != self.tw3:
                                # Поле: Трудоёмкость в часах
                                self.up2[f'{d2[t]}{d1[2]}0{i + 1}'] = tables[t].item(i, 1).text()
                                # Поле: Срок выполнения
                                # tab.cellWidget(i, 3).currentText() - смотрим содержимое
                                # ячеек поля "Срок выполнения" - названия месяцев учебного года
                                period = tables[t].cellWidget(i, 3).currentText()
                                self.up2[f'{d2[t]}{d1[3]}0{i + 1}'] = period
                                # Поле: Запланировано
                                itm = tables[t].item(i, 4).text()
                                self.up2[f'{d2[t]}{d1[4]}0{i + 1}'] = int(itm)
                            else:
                                # Поле: Объём п.л. или стр.
                                self.up2[f'{d2[t]}{d1[2]}0{i + 1}'] = tables[t].item(i, 3).text()
                                # Поле: Трудоёмкость в часах
                                self.up2[f'{d2[t]}{d1[4]}0{i + 1}'] = tables[t].item(i, 1).text()
                                # Поле: Срок выполнения
                                # tab.cellWidget(i, 3).currentText() - смотрим содержимое
                                # ячеек поля "Срок выполнения" - названия месяцев учебного года
                                period = tables[t].cellWidget(i, 4).currentText()
                                self.up2[f'{d2[t]}{d1[3]}0{i + 1}'] = period
                                # Поле: Запланировано
                                itm = tables[t].item(i, 5).text()
                                self.up2[f'{d2[t]}{d1[5]}0{i + 1}'] = int(itm)
                            # Распределение часов по осеннему и весеннему семестрам:
                            if period in ('сентябрь', 'октябрь', 'ноябрь',
                                          'декабрь', 'январь', '1 семестр'):
                                autumn += int(itm)
                            elif period in ('февраль', 'март', 'апрель',
                                            'май', 'июнь', '2 семестр'):
                                spring += int(itm)
                            # Если пользователь выбирает "в течение года",
                            # то в этом случае часы рспределяются пополам на оба семестра:
                            elif period == 'в течение года':
                                autumn += int(itm) // 2
                                spring += int(itm) - int(itm) // 2
                            summ += int(itm)
                            break
                # Записываем общую сумму по внеучебной работе
                self.up2[f'{d2[t]}YP'] = summ
                # Список часов по внеучебной работе на титульный лист. Индексы:
                # 0 - учебно-методическая
                # 1 - организационная
                # 2 - научно-исследовательская
                # 3 - воспитательная
                # 4 - повышение квалификации
                perext.append(summ)
                # Записываем суммы по осеннему и весеннему семестру:
                self.up2[f'{d2[t]}AP'] = autumn
                self.up2[f'{d2[t]}SP'] = spring
                # Суммируем "Всего":
                AP += autumn
                SP += spring
                YP += summ
            # Записываем "Всего" = <учебная_работа> + <внеучебная_работа>:
            self.up2['AP'] = self.up1['lrnAP'] + AP
            self.up2['SP'] = self.up1['lrnSP'] + SP
            self.up2['YP'] = self.up1['lrnYP'] + YP
            # Записываем проценты долей: учебно-методическая работа:
            self.up2['per_mtd'] = round(perext[0] / sum(perext[:-1]) * 100, 1)
            # Записываем проценты долей: организационная
            self.up2['per_org'] = round(perext[1] / sum(perext[:-1]) * 100, 1)
            # Записываем проценты долей: научно-исследовательская:
            self.up2['per_sci'] = round(perext[2] / sum(perext[:-1]) * 100, 1)
            # Записываем проценты долей: воспитательная:
            self.up2['per_edu'] = round(perext[3] / sum(perext[:-1]) * 100, 1)
            # Предупреждающее сообщение о неправильном распределении часов:
            '''
            if not self.percheck(self.percent_user_rate, self.up2['per_mtd'],
                                 self.up2['per_org'], self.up2['per_sci'], self.up2['per_edu']):
                msg = QMessageBox.warning(self, 'Внимание!',
                                          '<h4>Ошибка.<br>Распределение часов неправильное.</h4>')
            '''
            # Словарь для рендеринга:
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
                'user_rate': self.learn_rate,
                'per_ur': self.percent_user_rate
            }
            # Объединяем данные учебной (up1) и внеучебной (up2) работы
            # в словарь рендеринга для дальнейшего внесения в документ:
            context.update(self.up1)
            context.update(self.up2)
            # Диалоговое окно сохранения файла docx:
            saveDialog = QFileDialog()
            saveDialog.setDefaultSuffix('docx')
            fname, _ = saveDialog.getSaveFileName(self, 'Сохранить документ', '',
                                                  'Microsoft Word 2007–365 (*.docx)')
            if fname != str():
                self.statusBar().showMessage('Идёт процесс формирование документа...')
                # Создаём поток thread1 и передаём туда имя файла и данные для рендеринга:
                self.thread1 = Thread1(fname, context)
                # Сигнал запуска потока hread1 отправляем на слот thread1_start:
                self.thread1.started.connect(self.thread1_start)
                # Сигнал завершения потока thread1 отправляем на слот thread1_stop:
                self.thread1.finished.connect(self.thread1_stop)
                # Qt.QueuedConnection - сигнал помещается в очередь обработки событий интерфейса Qt:
                self.thread1.signal.connect(self.thread1_process, Qt.QueuedConnection)
                # Делаем кнопки неактивными:
                self.pb_lrn.setDisabled(True)
                self.pb_save.setDisabled(True)
                # Запускаем поток рендеринга:
                self.thread1.start(priority=QThread.IdlePriority)
            else:
                msg = QMessageBox.warning(self, 'Внимание!',
                                          '<h4>Вы не задали имя файла<br>для сохранения.</h4>')
        except Exception:
            msg = QMessageBox.warning(self, 'Внимание!',
                                      '<h4>Ещё не все поля заполнены.</h4>')

    def percheck(self, lrn, mtd, org, sci, edu):
        """Проверка правильности соотношения процентов"""
        # Учебная работа --- не более 60%
        # Научная и учебно-методическая --- 30% и более
        # Организационная и воспитательная --- не более 20%
        if lrn <= 60 and org <= 20 and edu <= 20 and mtd >= 30 and sci >= 30:
            return True
        else:
            return False

    def thread1_start(self):
        """Вызывается при запуске потока thread1"""
        # Выводим окно QProgressDialog на ожидание рендеринга.
        # HTML-сообщение с иконкой:
        self.save_error = False
        msg = '<table border = "0"> <tbody> <tr>' \
              '<td> <img src = "pic/save-icon.png"> </td>' \
              '<td> <h4>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Идёт сохранение документа,<br>' \
              '&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;подождите пожалуйста.</h4> </td>'
        self.dialog = QProgressDialog(msg, None, 0, 0, self)
        self.dialog.setModal(True)
        self.dialog.setWindowTitle('Инфо')
        self.dialog.setRange(0, 0)
        self.dialog.show()

    def thread1_process(self, s):
        """Вызывается сигналами которые отправляет поток thread1"""
        # Параметр s - это сигнал полученный из потока thread1
        if s == 'error':
            self.dialog.close()
            self.save_error = True
            msg = QMessageBox.warning(self, 'Внимание!',
                                      '<h4>Не удалось сохранить файл.<br>'
                                      'Возможно, у вас нет доступа<br>к целевой папке.</h4>')
            self.statusBar().showMessage('Не удалось создать файл')

    def thread1_stop(self):
        """Вызывается при завершении потока thread1"""
        self.dialog.close()
        if not self.save_error:
            # Выводим информационное сообщение:
            msg = QMessageBox.information(self, 'Инфо',
                                          '<h4>Индивидуальный план готов.</h4>')
            self.statusBar().showMessage('Документ сохранён')
        # Делаем кнопки "Открыть..." и "Сохранить..." активными:
        self.pb_lrn.setDisabled(False)
        self.pb_save.setDisabled(False)


def except_hook(cls, exception, traceback):
    """Функция для отслеживания ошибок PyQt5"""
    sys.__excepthook__(cls, exception, traceback)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = PlanForm()
    ex.show()
    # Ловим и показываем ошибки PyQt5 в терминале:
    sys.excepthook = except_hook
    sys.exit(app.exec_())
