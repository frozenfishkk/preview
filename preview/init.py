from PyQt5.QtWidgets import QMainWindow, QApplication
from custom_preview_ui import Ui_Form
from PyQt5 import QtCore
from PyQt5.QtGui import QTextDocument, QFont, QTextOption, QFontDatabase
from config import *
from PyQt5.QtCore import Qt

import json
import os
import re
import sys
import xlrd


class Init(QMainWindow, Ui_Form):
    app = QApplication(sys.argv)

    def __init__(self, path, excel):
        super().__init__()
        self.setWindowFlags(Qt.WindowStaysOnTopHint)
        self.setupUi(self)
        self.baseConfigPath = os.path.expanduser("~") + '\\PreviewConf\\'
        self.configName = 'config.json'
        self.configDict = {}
        self.InitConfig()
        self.text_document = QTextDocument()
        self.button_preview.clicked.connect(self.button_preview_clicked)
        self.spinbox_font.valueChanged.connect(self.spinbox_font_size_changed)
        self.spinbox_width.valueChanged.connect(self.spinbox_width_value_changed)
        self.clipboard = Init.app.clipboard()
        self.text_document.setMaximumBlockCount(10)
        self.text_edit.setWordWrapMode(QTextOption.WrapAnywhere)
        self.damagePath = path + '\\战斗\\英雄攻击效果表.xlsx'
        self.statePath = path + '\\战斗\\状态表.xlsx'
        self.confPath = path + '\\战斗\\属性配置表.xlsx'
        self.excel = excel
        self.damage = ''
        self.state = ''
        self.format_text = ''
        self.stateNum = ''
        self.stateData = {}
        self.damageData = {}
        self.confData = {}
        cwd_path = os.path.dirname(__file__)
        font_path = os.path.join(cwd_path, 'Source Han Sans CN Regular.ttf')
        font = QFontDatabase.addApplicationFont(font_path)
        font_family = QFontDatabase.applicationFontFamilies(font)[0]
        self.custom_font = QFont(font_family)
        if excel not in self.configDict:
            self.configDict[excel] = {'font': 14, 'width': 319}
        self.custom_font.setPointSize(self.configDict[excel]['font'])
        self.spinbox_font.setProperty('value', self.configDict[excel]['font'])
        self.spinbox_width.setProperty("value", self.configDict[excel]['width'])
        self.text_browser_right.setGeometry(
            QtCore.QRect(480, 40, self.configDict[excel]['width'], 381))

    def InitConfig(self):
        if not os.path.exists(self.baseConfigPath):
            os.makedirs(self.baseConfigPath)
        configPath = self.baseConfigPath + self.configName
        if os.path.exists(configPath):
            with open(configPath, 'r', encoding='utf-8') as f:
                self.configDict = json.loads(f.read())

    def saveDefaultConfig(self):
        configPath = self.baseConfigPath + self.configName
        self.configDict[self.excel] = {'font': self.spinbox_font.value(),
                                                     'width': self.spinbox_width.value()}
        with open(configPath, 'w', encoding='utf-8') as f:
            f.write(json.dumps(self.configDict))

    def button_preview_clicked(self):
        input_text = f'{self.text_edit.toPlainText()}'
        state_to_replace = re.findall(r"\|state\d+\|", input_text)
        stateNum_to_replace = re.findall(r"\|stateNum\d+\|", input_text)
        damage_to_replace = re.findall(r"\|damage\d+\|", input_text)
        for formatStr in state_to_replace:
            text = self.formatState(self.getID(formatStr))
            input_text = input_text.replace(formatStr, text)
        for formatStr in stateNum_to_replace:
            text = self.formatState(self.getID(formatStr), True)
            input_text = input_text.replace(formatStr, text)
        for formatStr in damage_to_replace:
            text = self.formatDamage(self.getID(formatStr))
            input_text = input_text.replace(formatStr, text)
        input_text = re.sub(r'<link="BKT_\d+">', r'<link="BKT_">', input_text)
        self.format_text = ((input_text.replace("<color=", "<font color=").replace("</color>", "</font>")
            .replace('<style="Link"><link="BKT_">', '<span style="text-decoration: underline;">').replace(
            '</link></style>', '</span>')))
        self.text_document.setHtml(self.format_text)
        self.text_document.setDefaultFont(self.custom_font)
        self.text_browser_right.setDocument(self.text_document)
        self.text_browser_right.document().setDefaultTextOption(self.text_option)
        self.saveDefaultConfig()

    def getID(self, formatStr):
        match = re.search(r"\d+", formatStr)
        if not match:
            return 0
        ID = int(match.group())
        return ID

    def spinbox_font_size_changed(self, value):
        self.custom_font.setPointSize(value)
        self.text_document.setDefaultFont(self.custom_font)

    def spinbox_width_value_changed(self, value):
        self.text_browser_right.setGeometry(QtCore.QRect(480, 40, value, 381))

    def formatState(self, ID, onlyNum=False):
        data = self.getStateDataByID(ID)
        text = ''
        count = 0
        existKey = []
        for key, value in data.items():
            value = self.formatValue(value)
            if not value:
                continue
            if key in STATE_MAP:
                existKey.append(key)
                if "%" in STATE_MAP[key]:
                    value = self.getPercentValue(value)
                if onlyNum:
                    text += value
                    continue
                count += 1
                if count > 1:
                    text += "，"
                text += STATE_MAP[key].replace("%", "") + str(value)
        for key, value in data.items():
            if key in existKey:
                continue
            value = self.formatValue(value)
            if not value:
                continue
            if key in STATE_CN_MAP:
                confData = self.getConfData(STATE_CN_MAP[key])
                cname = confData['显示名称']
                if not value:
                    continue
                if confData["是否百分比"]:
                    value = self.getPercentValue(value)
                if onlyNum:
                    value = value.replace("-", "")
                    text += str(value)
                    continue
                count += 1
                if count > 1:
                    text += "，"
                if "-" not in value:
                    text += f'{cname}提高{value}'
                else:
                    value = value.replace("-", "")
                    text += f'{cname}降低{value}'
        text = f"<font color=#f29133>{text}</font>"
        return text

    def formatDamage(self, ID):
        text = ""
        data = self.getDamageData(ID)
        if not data:
            return ""
        count = 0
        for key, replaceValue in DAMAGE_MAP.items():
            value = data.get(key)
            if not value:
                continue
            value = self.formatValue(value)
            count += 1
            if count > 1:
                text += "，"
            if "%" in replaceValue:
                value = self.getPercentValue(value)
                replaceValue = replaceValue.replace("%", "")
                text += f"{replaceValue}{value}"
            else:
                text += f"{replaceValue}{value}"
        text = f"<font color=#f29133>{text}</font>"
        return text

    def getPercentValue(self, value):
        value = f'{str(int(value / 100))}%'
        return value

    def formatValue(self, value):
        if type(value) in (int, float):
            return value
        elif type(value) == str:
            if ";" in value:
                value = value.split(";")[-1]
                value = int(value)
        return value

    def getStateDataByID(self, ID):
        if not self.stateData:
            workbook = xlrd.open_workbook(filename=self.statePath)
            sheet = workbook.sheet_by_name(sheet_name='状态表')
            rows = sheet.nrows
            sheetHeader = sheet.row_values(0)
            IDIndex = sheetHeader.index("状态ID")
            for row in range(1, rows):
                row_value = sheet.row_values(row)
                state_id = row_value[IDIndex]
                self.stateData[state_id] = {k: v for k, v in zip(sheetHeader, row_value)}
        return self.stateData.get(ID, {})

    def getDamageData(self, ID):
        if not self.damageData:
            workbook = xlrd.open_workbook(filename=self.damagePath)
            sheet = workbook.sheet_by_name(sheet_name="Sheet1")
            nrows = sheet.nrows
            sheetHeader = sheet.row_values(0)
            IDIndex = sheetHeader.index("攻击效果ID")
            for row in range(1, nrows):
                rowValues = sheet.row_values(row)
                ID = int(rowValues[IDIndex]) if rowValues[IDIndex] else 0
                self.damageData[ID] = {k: v for k, v in zip(sheetHeader, rowValues)}
        return self.damageData.get(ID, {})

    def getConfData(self, enName):
        if not self.confData:
            workbook = xlrd.open_workbook(filename=self.confPath)
            sheet = workbook.sheet_by_name(sheet_name="Sheet1")
            nrows = sheet.nrows
            sheetHeader = sheet.row_values(0)
            IDIndex = sheetHeader.index("属性（不可更改）")
            for row in range(1, nrows):
                rowValues = sheet.row_values(row)
                ID = rowValues[IDIndex] if rowValues[IDIndex] else ""
                if not ID:
                    continue
                self.confData[ID] = {k: v for k, v in zip(sheetHeader, rowValues)}
        return self.confData.get(enName, {})


