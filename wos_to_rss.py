#!/bin/env python3
# -*- coding: utf-8 -*-
import os
import time
import platform
import configparser
import xlrd
from datetime import datetime, timedelta
from bs4 import BeautifulSoup
import sqlite3
from DrissionPage import ChromiumPage
from DrissionPage.easy_set import set_headless
import translators as ts

if platform.system() == "Linux":
    from pyvirtualdisplay import Display

class WoS2RSS:

    def __init__(self, IN_SCHOOL, EMAIL, EMAIL_PASSWORD, UNIVERSITY, STUDENT_ID, STUDENT_PASSWORD, TRANSLATOR):
        self.XLS_PATH = './savedrecs.xls'
        self.XML_PATH = './wos.xml'
        self.DB_PATH  = './wos.db'
        self.IN_SCHOOL = IN_SCHOOL
        self.EMAIL = EMAIL
        self.EMAIL_PASSWORD = EMAIL_PASSWORD
        self.UNIVERSITY = UNIVERSITY
        self.STUDENT_ID = STUDENT_ID
        self.STUDENT_PASSWORD = STUDENT_PASSWORD
        self.TRANSLATOR = TRANSLATOR

        if platform.system() == "Linux":
            self.display = Display(visible=0, size=(1920, 1080))
            self.display.start()

        set_headless(False)
        self.page = ChromiumPage()
        self.page.set.window.fullscreen()
        self.page.clear_cache()

    def quit(self):
        self.page.ele('xpath://button[@data-ta="wos-header-user_name"]').click()
        time.sleep(1)
        self.page.ele('xpath://a[@aria-label="CloseSessionAndLogout"]').click()
        time.sleep(2)

        self.page.quit()
        if platform.system() == "Linux":
            self.display.stop()

        if os.path.exists(wos2rss.XLS_PATH):
            os.remove(wos2rss.XLS_PATH)

    def convert_to_timestamp(self, date_string):
        date_obj = datetime.strptime(date_string, "%a, %d %b %Y %H:%M:%S %z")
        return date_obj.timestamp()
 
    def get_xls_file(self):

        if self.page.ele('xpath://h1[@tabindex="-1"]').text.replace('alerting results for:', '').strip() == '0':
            return  False
        self.page.ele('xpath://button[contains(., "Export")]').wait.display(timeout=15)
        self.page.ele('xpath://button[contains(., "Export")]').click()
        time.sleep(1)
        self.page.ele('xpath://button[@id="exportToExcelButton"]').click()
        time.sleep(1)
        self.page.ele('xpath://span[contains(@class, "dropdown-text cdx-appirence-color")]').click()
        time.sleep(1)
        self.page.ele('xpath://div[@title="Author, Title, Source, Abstract"]').click()
        time.sleep(1)
        self.page.ele('xpath://span[contains(@class, "ng-star-inserted") and text()="Export"]').click()
        time.sleep(5)
        return True
    
    def check_notifications(self, alert_details):
        for alert_detail in alert_details:
            if 'notifications_none' not in alert_detail.text:
                return alert_detail
        return False

    def read_xls(self):
        workbook = xlrd.open_workbook(self.XLS_PATH)
        worksheet = workbook.sheet_by_index(0)
        headers = [worksheet.cell_value(0, col) for col in range(worksheet.ncols)]
        data = []
        for row in range(1, worksheet.nrows):
            row_data = {headers[col]: worksheet.cell_value(row, col) for col in range(worksheet.ncols)}
            data.append(row_data)
        return data

    def update_xml_file(self):
        # Create xml file
        if not os.path.exists(self.XML_PATH):
            content = '<rss xmlns:atom="http://www.w3.org/2005/Atom" version="2.0"><channel>'\
                    + '<title><![CDATA[' + "Web of Science" + ']]></title>' \
                    + '<link>' + 'https://www.webofscience.com/' + '</link>'\
                    + '<description><![CDATA[' + "Web of Science 订阅" + ']]></description>' \
                    + '<language>zh-cn</language>' \
                    + '</channel></rss>'
            with open(self.XML_PATH, 'w', encoding='utf-8') as file:
                file.write(content)
        # Create db file
        if not os.path.exists(self.DB_PATH):
            db_conn = sqlite3.connect(self.DB_PATH)
            db_cursor = db_conn.cursor()
            db_cursor.execute('CREATE TABLE wos (ArticleTitle TEXT, ArticleTitleZH TEXT, SourceTitle TEXT, AuthorFullNames TEXT, PublicationDate TEXT, Abstract TEXT, AbstractZH TEXT, DOI TEXT)')
            db_conn.commit()
            db_conn.close()

        db_conn = sqlite3.connect(self.DB_PATH)
        db_cursor = db_conn.cursor()

        # Read xml file
        xmlContent = ''
        with open(self.XML_PATH, 'r', encoding='utf-8') as file:
            xmlContent = file.read()
        xmlContent = BeautifulSoup(xmlContent.replace('link>','temptlink>'),'lxml')

        # Read xls file
        xlsList = self.read_xls()

        # Get exist papers
        existPaper = [row[0] for row in db_cursor.execute("SELECT DOI FROM wos").fetchall()]

        # Concreate new content
        for element in xlsList:
            if not element["DOI"] in existPaper:
                titleInZH, abstractInZH = '', ''
                if not self.TRANSLATOR == 'disabled':
                    titleInZH    = ts.translate_text(element["Article Title"], translator=self.TRANSLATOR,to_language='zh')
                    abstractInZH = ts.translate_text(element["Abstract"], translator=self.TRANSLATOR,to_language='zh')
                elementContent = '<title><![CDATA[' + element["Article Title"] + ']]></title>' \
                                + '<description><![CDATA[' + "<p><b>" + titleInZH + "</b></p>" \
                                + "<p>" + element["Source Title"].title() + ", " + element["Publication Date"] + "</p>" \
                                + "<p><u>" + element["Author Full Names"] + "</u></p>" \
                                + "<p>" + "<b>摘要:</b>" + abstractInZH + "</p>" \
                                + "<p>" + "<b>Abstract:</b>" + element["Abstract"] + "</p>" + ']]></description>'\
                                + '<temptlink>' + "http://dx.doi.org/" + element["DOI"] + '</temptlink>'\
                                + '<pubDate>' + time.strftime("%a, %d %b %Y %H:%M:%S %z", time.localtime(int(time.time()))) + '</pubDate>'
                
                parent = xmlContent.select_one('channel')  # get parent element
                new_item = xmlContent.new_tag('item')  # create new item element
                new_item.string = elementContent
                parent.append(new_item)

                dbdata = (element["Article Title"], titleInZH, element["Source Title"], element["Author Full Names"], 
                            element["Publication Date"], element["Abstract"], abstractInZH, element["DOI"])
                db_cursor.execute("INSERT INTO wos VALUES (?,?,?,?,?,?,?,?)", dbdata)
                            
        xmlContent = BeautifulSoup(str(xmlContent.body.contents[0]).replace('&amp;','&').replace('&lt;','<').replace('&gt;','>'),'lxml')
        items = xmlContent.find_all('item')
        # sort <item> by pubDate
        sorted_items = sorted(items, key=lambda x: self.convert_to_timestamp(x.select_one('pubDate').text), reverse=True)

        # remove <item> older than 2 weeks if there are more than 100 <item>
        if len(sorted_items) > 100:
            # get timestamp of 2 weeks ago
            two_week_ago = datetime.now() - timedelta(days=14)
            two_week_ago_timestamp = time.mktime(two_week_ago.timetuple())
            for item in sorted_items.copy():
                pub_date = item.select_one('pubDate').text
                pub_date_timestamp = self.convert_to_timestamp(pub_date)
                if pub_date_timestamp < two_week_ago_timestamp:
                    sorted_items.remove(item)

        # remove all <item> in xmlContent
        items_in_xmlContent = xmlContent.find_all('item')
        for item in items_in_xmlContent:
            item.extract()
        
        # append sorted <item> to xmlContent
        parent_element = xmlContent.find('channel')
        for sorted_item in sorted_items:
            parent_element.append(sorted_item)

        with open(self.XML_PATH, 'w', encoding='utf-8') as f:
            f.write(str(xmlContent.body.contents[0]).replace('&lt;','<').replace('&gt;','>').replace('temptlink','link'))

        db_conn.commit()
        db_conn.close()

        os.remove(self.XLS_PATH)
    
    def update_alerts(self):
        
        alert_details = self.page.eles('xpath://div[@class="alert-details"]')
        while self.check_notifications(alert_details):
            self.check_notifications(alert_details).click()
            time.sleep(2)
            if self.get_xls_file():
                self.update_xml_file()
                time.sleep(1)
            self.page.back()
            time.sleep(1)
            # page element would be refreshed after page back
            alert_details = self.page.eles('xpath://div[@class="alert-details"]')

    def email_cert(self):
        # wait loading
        self.page.get('https://webofscience.com')
        self.page.ele('xpath://img[@src="/public/assets/img/wos-1.svg"]',timeout=15)
        # login
        if self.page.ele('xpath://input[@name="email"]',timeout=10):
            self.page.ele('xpath://input[@name="email"]').clear()
            self.page.ele('xpath://input[@name="email"]').input(self.EMAIL)
            self.page.ele('xpath://input[@name="password"]').clear()
            self.page.ele('xpath://input[@name="password"]').input(self.EMAIL_PASSWORD)
            self.page.ele('xpath://button[@id="signIn-btn"]').click()
            if self.page.ele('xpath://a[text()="continue and establish a new session"]'):
                self.page.ele('xpath://a[text()="continue and establish a new session"]').click()
            # wait loading
            self.page.ele('xpath://button[@data-ta="wos-header-user_name"]', timeout=15)

        # switch to lang-en
        lang_button = self.page.ele('xpath://*[normalize-space(text())="简体中文"]',timeout=2)
        if lang_button:
            lang_button.click()
            self.page.ele('xpath://button[@lang="en"]').wait.display(timeout=15)
            self.page.ele('xpath://button[@lang="en"]').click()
        
        time.sleep(1)

    def carsi_cert(self):
        self.page.get('https://webofscience.com')
        time.sleep(1)
        self.page.ele('xpath://span[@class="ng-tns-c82-2 ng-star-inserted"]').wait.display(timeout=15)
        self.page.ele('xpath://span[@class="ng-tns-c82-2 ng-star-inserted"]').click()
        time.sleep(1)
        self.page.ele('xpath://span[contains(., "CHINA CERNET Federation")]').click()
        time.sleep(1)
        self.page.ele('xpath://button[contains(., "Go to institution")]').click()

        # select school
        univ_input = self.page.ele('xpath://input[@id="show"]')
        univ_input.wait.display(timeout=15)
        univ_input.click()
        univ_input.clear()
        univ_input.input(self.UNIVERSITY)
        time.sleep(1)
        self.page.ele('xpath://ul[@class="typeahead dropdown-menu pre-scrollable test-5 select-ul"]/li[1]/a').click()
        time.sleep(1)
        self.page.ele('xpath://button[@id="idpSkipButton"]').click()
        time.sleep(1)

        # Warning bypass
        warning_button = self.page.ele('xpath://a[@class="logout_button" and text()="继续登录"]',timeout=2)
        if warning_button:
            warning_button.click()

        # school certification
        self.page.ele('xpath://input[@name="username"]').wait.display(timeout=15)
        self.page.ele('xpath://input[@name="username"]').clear()
        self.page.ele('xpath://input[@name="username"]').input(self.STUDENT_ID)
        self.page.ele('xpath://input[@name="password"]').clear()
        self.page.ele('xpath://input[@name="password"]').input(self.STUDENT_PASSWORD)
        time.sleep(2)
        self.page.ele('xpath://button[@id="dl"]').click()

        # wait loading
        time.sleep(10)
        # pop windows bypass
        if self.page.ele('xpath://*[@id="onetrust-accept-btn-handler"]',timeout=1):
            self.page.ele('xpath://*[@id="onetrust-accept-btn-handler"]').click()
        if self.page.ele('xpath://button[contains(@class, "_pendo-button-primaryButton")]"]',timeout=1):
            self.page.ele('xpath://button[contains(@class, "_pendo-button-primaryButton")]').click()
        if self.page.ele('xpath://button[contains(@class, "_pendo-button-secondaryButton")]',timeout=1):
            self.page.ele('xpath://button[contains(@class, "_pendo-button-secondaryButton")]').click()
        if self.page.ele('xpath://span[contains(@class, "_pendo-close-guide")',timeout=1):
            self.page.ele('xpath://span[contains(@class, "_pendo-close-guide")').click()
        if self.page.ele('xpath://div[contains(@class, "cdk-overlay-container")',timeout=1):
            self.page.ele('xpath://div[contains(@class, "cdk-overlay-container")').click()
        time.sleep(1)
        # switch to lang-en
        lang_button = self.page.ele('xpath://*[normalize-space(text())="简体中文"]',timeout=1)
        if lang_button:
            lang_button.click()
            self.page.ele('xpath://button[@lang="en"]').wait.display(timeout=15)
            self.page.ele('xpath://button[@lang="en"]').click()
        time.sleep(1)
        if self.page.ele('xpath://button[@title="Sign in to access"]').wait.display(timeout=15):
            self.page.ele('xpath://button[@title="Sign in to access"]').click()
        # wait loading
        self.page.ele('xpath://img[@src="/public/assets/img/wos-1.svg"]',timeout=15)
        # login
        if self.page.ele('xpath://input[@name="email"]',timeout=10):
            self.page.ele('xpath://input[@name="email"]').clear()
            self.page.ele('xpath://input[@name="email"]').input(self.EMAIL)
            self.page.ele('xpath://input[@name="password"]').clear()
            self.page.ele('xpath://input[@name="password"]').input(self.EMAIL_PASSWORD)
            self.page.ele('xpath://button[@id="signIn-btn"]').click()
            if self.page.ele('xpath://a[text()="continue and establish a new session"]'):
                self.page.ele('xpath://a[text()="continue and establish a new session"]').click()
            # wait loading
            self.page.ele('xpath://button[@data-ta="wos-header-user_name"]', timeout=15)
        time.sleep(1)
        
if __name__ == '__main__':

    config = configparser.ConfigParser()
    config.read('./config.ini', encoding='utf-8')
    IN_SCHOOL = config.getboolean('ID', 'IN_SCHOOL')
    EMAIL = config.get('ID', 'EMAIL')
    EMAIL_PASSWORD = config.get('ID', 'EMAIL_PASSWORD')
    UNIVERSITY, STUDENT_ID, STUDENT_PASSWORD = '', '', ''
    if not IN_SCHOOL:
        UNIVERSITY = config.get('ID', 'UNIVERSITY')
        STUDENT_ID = config.get('ID', 'STUDENT_ID')
        STUDENT_PASSWORD = config.get('ID', 'STUDENT_PASSWORD')
    TRANSLATOR = config.get('Translator', 'TRANSLATOR')

    wos2rss = WoS2RSS(IN_SCHOOL, EMAIL, EMAIL_PASSWORD, UNIVERSITY, STUDENT_ID, STUDENT_PASSWORD, TRANSLATOR)
  
    # Get xls file from WOS
    if wos2rss.IN_SCHOOL:
        wos2rss.email_cert()
    else:
        wos2rss.carsi_cert()

    if os.path.exists(wos2rss.XLS_PATH):
        os.remove(wos2rss.XLS_PATH)

    # download xls file and transcode to xml file
    wos2rss.update_alerts()
    
    wos2rss.quit()
