import re
import sys
import csv
import lxml
import requests
import pandas as pd
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QTextDocument, QStandardItemModel, QStandardItem
from requests.adapters import HTTPAdapter
from urllib3.util import Retry
from striprtf.striprtf import rtf_to_text
from PyQt6 import QtCore, QtPrintSupport
from PyQt6.QtWidgets import QApplication, QWidget, QHeaderView, QTableWidgetItem, QTableWidget, QSizePolicy, QFileDialog
from bs4 import BeautifulSoup
from yattag import Doc
from form_regions import Ui_Form

link_license = 'https://rkn.gov.ru/activity/connection/register/license/'
headers = {
    'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
    'accept-language': 'ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7',
    'cache-control': 'max-age=0',
    'dnt': '1',
    'priority': 'u=0, i',
    'sec-ch-ua': '"Not)A;Brand";v="99", "Google Chrome";v="127", "Chromium";v="127"',
    'sec-ch-ua-mobile': '?0',
    'sec-ch-ua-platform': '"Windows"',
    'sec-fetch-dest': 'document',
    'sec-fetch-mode': 'navigate',
    'sec-fetch-site': 'none',
    'sec-fetch-user': '?1',
    'sec-gpc': '1',
    'upgrade-insecure-requests': '1',
    'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/127.0.0.0 Safari/537.36',
    }

class Window(QWidget, Ui_Form):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.setWindowTitle('Parse license')
        self.le_inn.textChanged.connect(self.inn_input)
        self.btn_quit.clicked.connect(app.exit)
        self.btn_search.clicked.connect(self.inn_search)
        self.btn_save.clicked.connect(self.save_xlsx)
        self.btn_print.clicked.connect(self.print_result)
        self.comboBox.activated.connect(self.get_region)
        self.btn_save.setDisabled(True)
        self.btn_print.setDisabled(True)
        self.result_list = []
        self.table_license.setStyleSheet('QHeaderView::section {color: white; background-color:grey; border:grey;}')
        self.table_license.setAlternatingRowColors(True)
        self.table_license.verticalHeader().hide()
        self.table_license.verticalHeader().setStretchLastSection(False)
        self.load_regions()
        self.region = 0

    def load_regions(self):
        model = QStandardItemModel(0, 1)
        with open('regions.csv', 'r', newline='', encoding='utf-8') as file:
            file.readline()
            for reg_id, reg_name in csv.reader(file):
                item = QStandardItem(reg_name)
                item.setData(reg_id, Qt.ItemDataRole.UserRole)
                model.appendRow(item)
        self.comboBox.setModel(model)
        self.comboBox.setCurrentIndex(-1)

    def get_region(self, index):
        self.region = int(self.comboBox.itemData(index,Qt.ItemDataRole.UserRole))
        if self.region: self.btn_search.setDisabled(False)
        elif not self.le_inn.text(): self.btn_search.setDisabled(True)

    def draw_table(self):
        params = ['Сокращенное наименование', 'Регистрационный номер лицензии', 'День начала оказания услуг', 'Срок действия до', 'Территория действия лицензии', 'Лицензируемый вид деятельности с указанием выполняемых работ, составляющих лицензируемый вид деятельности']
        column_names = ['Лицензиат', '№ лицензии', 'Дата начала', 'Действует до', 'Территория действия', 'Вид деятельности', 'Технические параметры']
        if self.le_inn.text():
            params = params[1:]
            column_names = column_names[1:]
            license_data = self.result_list[0]
            self.label_name.setText(license_data.get('Сокращенное наименование'))
            self.label_address.setText(license_data.get('Адрес места нахождения'))
            self.label_phone.setText(license_data.get('Номер телефона'))
            self.label_mail.setText(license_data.get('Адрес электронной почты'))
        self.table_license.setRowCount(len(self.result_list))
        self.table_license.setColumnCount(len(params)+1)
        self.table_license.setHorizontalHeaderLabels(column_names)
        self.table_license.horizontalHeader().setMaximumSectionSize(300)
        for i, license_ in enumerate(self.result_list):
            if 'Сокращенное наименование' not in license_: license_['Сокращенное наименование'] = license_['Полное наименование лицензиата']
            for j, param in enumerate(params):
                item = QTableWidgetItem(license_.get(param))
                item.setTextAlignment(Qt.AlignmentFlag.AlignTop)
                self.table_license.setItem(i, j, item)
            if data:=license_.get('data'):
                self.table_license.setCellWidget(i, len(params), self.inner_tbl(data)) 
        self.table_license.resizeColumnsToContents()
        self.table_license.resizeRowsToContents()

    def inner_tbl(self,data):
        names_dict = {'region':'регион', 'friq':'частота', 'chanel':'канал', 'power':'мощность'}
        column_names = [names_dict[k] for k in data[0]]
        tbl = QTableWidget()
        tbl.setColumnCount(len(data[0]))
        tbl.setRowCount(len(data))
        tbl.verticalHeader().hide()
        tbl.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum)
        tbl.setHorizontalHeaderLabels(column_names)
        v_header = tbl.verticalHeader()
        h_header = tbl.horizontalHeader()
        tbl.setFixedHeight((v_header.sectionSize(0) + 1) * tbl.rowCount() + h_header.height())
        for i, d in enumerate(data):
            for j, col in enumerate(d.keys()):
                tbl.setItem(i, j, QTableWidgetItem(d.get(col)))
        tbl.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        tbl.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        tbl.setStyleSheet('QHeaderView::section {color: #373737; background-color:silver;border:grey};')
        tbl.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        return tbl

    @QtCore.pyqtSlot()
    def inn_input(self):
        self.btn_search.setEnabled(len(self.le_inn.text())==10)

    def inn_search(self):
        inn_input = self.le_inn.text()
        status_input = 1  # 1-действующая, 0-все
        session = requests.Session()
        retries = Retry(total=5,
                        backoff_factor=0.5,
                        status_forcelist=[500, 502, 503, 504],
                        allowed_methods=frozenset(['GET', 'POST']))
        session.mount('http://', HTTPAdapter(max_retries=retries))
        session.mount('https://', HTTPAdapter(max_retries=retries))
        response_license = session.get(link_license, headers=headers)
        soup = BeautifulSoup(response_license.content, 'lxml')
        token = soup.find('meta', {'name': 'csrf-token-value'})['content']
        cookies = response_license.cookies.get_dict() | {'csrf-token-name': 'csrftoken', 'csrf-token-value': token}
        data = {'act': 'search',
                'org_name_full': '',
                'org_inn': inn_input,
                'lic_num': '',
                'lic_status_id': status_input,
                'service_id': 14 if self.region else 0,
                'region_id': self.region,
                'searchBtn': 'Найти'}
        post_response = session.get(link_license, headers=headers, cookies=cookies, params=data)
        soup = BeautifulSoup(post_response.text, 'lxml')
        res_list = soup.find(id='ResList1')
        if res_list:
            license_links = [link_license + a.get('href') for a in res_list.find_all('a')]
            pages = soup.find(class_='paginationList')
            if pages:
                for link in [a.get('href') for a in pages.find_all('a')]:
                    page_response = session.get('https://rkn.gov.ru' + link, headers=headers, cookies=cookies)
                    soup = BeautifulSoup(page_response.text, 'lxml')
                    license_links.extend([link_license + a.get('href') for a in soup.find(id='ResList1').find_all('a')])
            self.parse_page(license_links, cookies, session)
            self.parse_files(session)
            self.draw_table()
            self.btn_save.setDisabled(False)
            self.btn_print.setDisabled(False)
        else:
            self.label_name.setText('Записей не найдено')
            self.label_address.setText('')
            self.label_phone.setText('')
            self.label_mail.setText('')
            self.table_license.setRowCount(0)

    def parse_page(self, links, cookies, session):
        self.result_list = []
        for url in links:
            text = session.get(url, headers=headers, cookies=cookies).text
            license_dict = {}
            soup = BeautifulSoup(text, 'lxml')
            tbl_list = soup.find(class_='TblList').find_all('tr')
            for tr in tbl_list:
                td1, td2 = tr.find_all('td')
                if td2.string:
                    license_dict |= {td1.string: td2.string}
                elif td1.string in ['Номер телефона', 'Адрес электронной почты']:
                    license_dict |= {td1.string: td2.div.string}
            if getfile_link := soup.find(href=lambda href: href and re.search('getFile', href))['href']:
                file_url = link_license + getfile_link
                license_dict |= {'url': file_url}
            self.result_list.append(license_dict)

    def parse_files(self, session):
        for license in self.result_list:
            if 'url' in license:
                text = session.get(license.get('url'), headers=headers).content
                text = rtf_to_text(text.decode('cp1251'))
                pattern = ''
                m = re.search(r'Территория оказания\s?\n?услуги связи(\|)(Частота \(МГц\)\|)?(Частотный канал \(ТВК\))?(\|)?', text)
                if m:
                    _, freq, chan, __ = m.groups()
                    if freq and chan: pattern = re.compile(r'\|(?P<region>[^|]+?)\|(?P<friq>(\d{2,3},\d)|-)\|(?P<chanel>\d{1,2}|-)\|(?P<power>\d,?\d{,2})')
                    elif freq: pattern = re.compile(r'\|(?P<region>[^|]+?)\|(?P<friq>\d{2,3},\d)\|(?P<power>\d,?\d{,2})')
                    elif chan: pattern = re.compile(r'\|(?P<region>[^|]+?)\|(?P<chanel>\d{1,2})\|(?P<power>\d,?\d{,2})')
                    license['data'] = [m.groupdict() for m in pattern.finditer(text)] if pattern else []
                else: license['data'] = []

    def save_xlsx(self):
        file_name, _ = QFileDialog.getSaveFileName(self, 'Save as xlsx', '', 'Excel files(*.xlsx)')
        if file_name:
            df = pd.DataFrame(self.result_list)
            license_df = df.loc[df.data.map(len) != 0][['Сокращенное наименование', 'Полное наименование лицензиата', 'Территория действия лицензии', 'data']]
            res_df = license_df.explode('data', ignore_index=True)
            result = res_df.merge(res_df.data.apply(pd.Series), left_index=True, right_index=True).drop('data', axis=1)
            result.to_excel(file_name, sheet_name='license', index=False)

    def print_result(self):
        style_sheet = '''table {border-collapse:collapse; width:100%}
                            th {background-color:lightblue; border: 1px solid gray; height:1em}
                            td {border: 1px solid gray; padding:0 1em 0 1em; vertical-align:top}
                            td.params {padding:0}
                            tr.head_inner {background-color:lightgray; font-weight:normal; text-align:center}
                            '''
        text_doc = QTextDocument()
        text_doc.setDefaultStyleSheet(style_sheet)
        text_doc.setHtml(self.get_text_doc())
        prev_dialog = QtPrintSupport.QPrintPreviewDialog()
        prev_dialog.paintRequested.connect(text_doc.print)
        prev_dialog.exec()

    def get_text_doc(self):
        params = ['Сокращенное наименование', 'Регистрационный номер лицензии', 'День начала оказания услуг', 'Срок действия до',
                  'Территория действия лицензии',
                  'Лицензируемый вид деятельности с указанием выполняемых работ, составляющих лицензируемый вид деятельности']
        params = params[1:] if self.le_inn.text() else params[:-1]
        names_dict = {'region': 'регион', 'friq': 'частота', 'chanel': 'канал', 'power': 'мощность'}
        license_data = self.result_list[0]
        doc, tag, text, line = Doc().ttl()

        with tag('html'):
            # doc.stag('br')
            if self.le_inn.text():
                with tag('ul'):
                    line('li', 'Контрагент: ' + license_data.get('Сокращенное наименование'))
                    line('li', 'Адрес: ' + license_data.get('Адрес места нахождения'))
                    line('li', 'Телефон: ' + license_data.get('Номер телефона'))
                    line('li', 'Email: ' + license_data.get('Адрес электронной почты'))
                doc.stag('br')
            with tag('table', klass='license', cellspacing='0', cellpadding='0', width='100%'):
                with tag('tr', klass='head'):
                    if self.le_inn.text(): header = '<th>№ лицензии</th><th>Дата начала</th><th>Действует до</th><th>Территория действия</th><th>Лицензируемый вид деятельности</th><th>Технические параметры</th>'
                    else: header = '<th>Лицензиат</th><th>№ лицензии</th><th>Дата начала</th><th>Действует до</th><th>Территория действия</th><th>Технические параметры</th>'
                    doc.asis(header)
                for lic in self.result_list:
                    with tag('tr'):
                        for param in params:
                            line('td', lic.get(param))
                        if data:=lic.get('data'):
                            column_names = [names_dict[k] for k in data[0]]
                            with tag('td', klass='params'):
                                with tag('table', klass='inner_tbl', cellspacing='0', cellpadding='0', width='100%'):
                                    with tag('tr', klass='head_inner'):
                                        doc.asis(''.join(f'<td>{name}</td>' for name in column_names))
                                    for d in data:
                                        with tag('tr'):
                                            doc.asis(''.join(f'<td>{v}</td>' for k,v in d.items()))
                        else: doc.asis('<td></td>')
        return doc.getvalue()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = Window()
    window.show()
    sys.exit(app.exec())