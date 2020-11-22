# iPlan 0.0.5
# Автоматическая генерация индивидуального плана преподавателя
# motor1n develop PyQt5 - 2020 year


import sys, xlrd
import datetime as dt
from PyQt5 import uic
from docxtpl import DocxTemplate
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtWidgets import (QApplication, QMainWindow, QFileDialog,
                             QTreeWidgetItemIterator, QTableWidgetItem,
                             QComboBox, QMessageBox)


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


class Thread1(QThread):
    """Поток для рендеринга и сохранения файла"""
    # Создаём собственный сигнал,
    # принимающий параметр типа str:
    signal = pyqtSignal(str)
    # Инициализация потока
    # fname - имя сохраняемого файла
    # content - дянные для рендеринга
    def __init__(self, fname, contect, parent=None):
        QThread.__init__(self, parent)
        self.fname = fname
        self.context = contect

    # Обязательный для любого потока метод run,
    # в котором происходит основной процесс:
    def run(self):
        # Подключаем файл шаблона .dotx:
        doc = DocxTemplate('iplan-template.dotx')
        # Рендерим инфу в шаблон
        self.signal.emit('Рендерим инфу в шаблон')
        doc.render(self.context)
        # Сохраняем конечный документ
        self.signal.emit('Сохраняем конечный документ')
        doc.save(self.fname)


class Thread2(QThread):
    # Создаём собственный сигнал,
    # передающий параметр типа int:
    signal = pyqtSignal(int)
    # Инициализация потока
    def __init__(self,  parent=None):
        QThread.__init__(self, parent)

    # Счётчик для прогресс-бара
    def run(self):
        for i in range(101):
            i += 1
            QThread.msleep(30)
            self.signal.emit(i)


class PlanForm(QMainWindow):
    def __init__(self):
        super().__init__()
        uic.loadUi('iplan-design.ui', self)
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
        self.compete_tabs = False
        # Словарь - состояние заполненности таблиц QTableWidget:
        self.condition_tabs = dict()
        # Кортеж кнопок QComboBox на заполнение данных пользователя:
        self.cbX = (self.cb1, self.cb2, self.cb3, self.cb4)
        # Кортеж объектов QTreeWidget
        self.treeX = (self.tree1, self.tree2, self.tree3, self.tree4, self.tree5)
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
            self.pb0X[i].clicked.connect(lambda checked, tree=self.treeX[i],
                                                tab=self.tables[i]: self.extra(tree, tab))
        # Сигнал pb3 --> слот savedocx
        self.pb_save.clicked.connect(lambda checked, tables=self.tables: self.savedocx(tables))
        # Сигналы отслеживания изменений таблиц QTableWidget (кортеж tables),
        # но без "Повышения квалификации", поскольку она не обязательна
        for tab in self.tables[:-1]:
            tab.cellChanged.connect(self.comlete_alltabs)
            # И попутно заполнение словаря self.condition_tabs значениями False,
            # т.е. пока ещё ни одна таблица не заполнена полностью
            self.condition_tabs[tab.objectName()] = False
        # Сигналы отслеживания изменений QComboBox на заполнение данных пользователя:
        for cb in self.cbX:
            cb.currentTextChanged.connect(self.user)

    def user(self):
        """Контроль заполнения данных о пользователе"""
        if '---' not in [cb.currentText() for cb in self.cbX]:
            # Если данные заполнены, активируем кнопку "Открыть..."
            self.pb_lrn.setDisabled(False)

    def learn(self):
        """Заполнение данных по учебной работе"""
        fname = QFileDialog.getOpenFileName(self, 'Выбрать файл', '',
                                            'Excel 2007–365 (.xlsx)(*.xlsx)')[0]
        # Флаг: файл учебной нагрузки открыт
        msg = QMessageBox.information(self, 'Инфо',
                                      '<h4>Файл учебной нагрузки открыт,<br>можно продолжить работу.</h4>')
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
        self.percent_user_rate = round(self.learn_rate / (RATE * sh.cell(sh.nrows - 3, 0).value) * 100, 1)
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
        # Учебная работа заполнена,
        # делаем активной для заполнеия внеучебную работу
        self.tabs.setDisabled(False)

    def extra(self, tree, tab):
        """Заполнение данных по внеучебной работе"""
        # Загрузка отмеченных элементов QTreeWidgetItem в таблицу QTableWidget
        # Список для найденых выделенных check
        self.checklist = list()
        # Создаём итератор для прохода по элементам QTreeWidget
        iter = QTreeWidgetItemIterator(tree, QTreeWidgetItemIterator.Checked)
        while iter.value():
            # Читаем строку QTreeWidgetItem
            currentItem = iter.value()
            # currentItem.text(0) - текст в ячейке "Виды работы"
            # currentItem.toolTip(1) - всплывающая подсказка ячейки "Трудоёмкость"
            # currentItem.text(2) - текст в ячейке "Форма отчётности"
            if tab.objectName() == 'tw3':
                # Если таблица "Научная работа", то дополняется столбец "Объём п.л. или стр."
                self.checklist.append((currentItem.text(0), currentItem.toolTip(1), currentItem.text(2), currentItem.text(3)))
            else:
                self.checklist.append((currentItem.text(0), currentItem.toolTip(1), currentItem.text(2)))
            iter += 1
        # Если ничего не выбрано,
        # то выведем сообщение об этом в статус-бар и вернём пустой return
        if not self.checklist:
            self.statusBar().showMessage('Внеучебная работа: не выбрано')
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
            msg = f'Внеучебная работа - выбрано позиций: {len(self.checklist)}'
            self.statusBar().showMessage(msg)
        # Помещаем кнопки QComboBox
        # в поле "Срок выполнения" на таблицу QTableWidget
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

    def is_tabfull(self, tab):
        """Проверка заполненности одной таблицы QTableWidget"""
        for i in range(len(self.checklist)):
            for j in range(tab.columnCount()):
                # Проверяем все ячейки кроме столбца "Срок выполнения"
                if tab == self.tw3:
                    if j != 4 and tab.item(i, j) is None:
                        return False
                else:
                    if j != 3 and tab.item(i, j) is None:
                        return False
        return True

    def comlete_alltabs(self):
        """Проверка заполненности всех таблиц QTableWidget"""
        if self.is_tabfull(self.sender()):
            self.condition_tabs[self.sender().objectName()] = True
        # Если всё заполнено, активируем кнопку "Сохранить..."
        if all(self.condition_tabs.values()):
            self.pb_save.setDisabled(False)

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
        # Строка "Всего" в таблице "Распределение времени по семестрам и основным видам работы,
        # общие суммы по плану внеучебной работы (осенний, весенний, год)
        AP = 0
        SP = 0
        YP = 0
        # Список для записи часов по внеучебной работе на титульный лист:
        perext = list()
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
            # Записываем суммы по осеннему и весеннему семестру
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
        # Записываем проценты долей: учебно-методическая работа
        self.up2['per_mtd'] = round(perext[0] / sum(perext[:-1]) * 100, 1)
        # Записываем проценты долей: организационная
        self.up2['per_org'] = round(perext[1] / sum(perext[:-1]) * 100, 1)
        # Записываем проценты долей: научно-исследовательская
        self.up2['per_sci'] = round(perext[2] / sum(perext[:-1]) * 100, 1)
        # Записываем проценты долей: воспитательная
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
        # в словарь рендеринга для дальнейшего внесения в документ
        context.update(self.up1)
        context.update(self.up2)
        # Диалоговое окно сохранения файла docx
        fname = QFileDialog.getSaveFileName(self, 'Сохранить документ', '',
                                            'Word 2007–365 (.docx)(*.docx)')[0]
        self.statusBar().showMessage('Идёт формирование документа...')
        # Создаём поток thread1 и передаём туда имя файла и данные для рендеринга:
        self.thread1 = Thread1(fname, context)
        # Создаём поток thread2 для счётчика прогресс-бара:
        self.thread2 = Thread2()
        # Сигнал запуска потока hread1 отправляем на слот thread1_start
        self.thread1.started.connect(self.thread1_start)
        # Сигнал завершения потока thread1 отправляем на слот thread1_stop
        self.thread1.finished.connect(self.thread1_stop)
        # Сигнал завершения потока thread2 отправляем на слот thread2_stop
        self.thread2.finished.connect(self.thread2_stop)
        # Cигнал из потока thread1 отправляем в основную программу на слот thread_process
        # Qt.QueuedConnection - сигнал помещается в очередь обработки событий интерфейса Qt.
        self.thread1.signal.connect(self.thread1_process, Qt.QueuedConnection)
        self.thread2.signal.connect(self.thread2_process, Qt.QueuedConnection)
        # Делаем кнопки неактивными
        self.pb_lrn.setDisabled(True)
        self.pb_save.setDisabled(True)
        # Запускаем поток рендеринга
        # IdlePriority - самый низкий приоритет
        self.thread1.start(priority=QThread.IdlePriority)

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
        """Вызывается при запуске потока"""
        # Запускаем поток прогресс-бара
        # InheritPriority - автоматический приоритет
        self.thread2.start(priority=QThread.InheritPriority)

    def thread1_process(self, s):
        """Вызывается сигналами которые отправляет поток"""
        # Параметр s - это сигнал полученный из потока thread1
        self.statusBar().showMessage(s)

    def thread1_stop(self):
        """Вызывается при завершении потока"""
        self.statusBar().showMessage('Рендеринг выполняется...')

    def thread2_process(self, i):
        """Вызывается сигналами которые отправляет поток"""
        # Счётчик из потока thread2 увеличивает прогресс-бар
        self.progress.setValue(i)

    def thread2_stop(self):
        # Выводим сообщение в статус-бар
        self.statusBar().showMessage('Документ сформирован')
        # Делаем кнопки "Открыть..." и "Сохранить..." активными:
        self.pb_lrn.setDisabled(False)
        self.pb_save.setDisabled(False)
        # Обнуляем прогресс-бар
        self.progress.setValue(0)
        # Выводим информационное сообщение
        msg = QMessageBox.information(self, 'Инфо',
                                      '<h4>Индивидуальный план готов.</h4>')


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
