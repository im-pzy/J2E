from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, \
    QPlainTextEdit, QPushButton, QTableWidget, QTableWidgetItem, \
    QLabel, QMenu, QAbstractItemView, \
    QMessageBox, QFileDialog, \
    QHBoxLayout, QVBoxLayout, \
    QShortcut
from PyQt5.QtCore import Qt, QStandardPaths, pyqtSignal
from PyQt5.QtGui import QIcon
from xlsxwriter import Workbook
import json, os

DEFAULT_JSON_DATA = []
DEFAULT_DATA_INFO = (0, 0)  # (col,row)
DEFAULT_INFO = "✅已就绪"
DEFAULT_FIELDS = []


class MainWindow(QMainWindow):

    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle("JsonArray转Excel工具")
        self.setWindowIcon(QIcon(os.path.join(os.path.dirname(__file__), "icons/J2E.ico")))
        self.resize(1000, 500)
        self.centralWidget = QWidget(self)  # 中心部件和布局
        self.setCentralWidget(self.centralWidget)

        # 实例化
        self.dataModel = DataModel()  # 数据结构
        self.importJsonButton = ImportJsonButton(self, self.dataModel)  # 导入文件按钮
        self.formatButton = FormatButton(self, self.dataModel)  # 格式化按钮
        self.outputExcelButton = OutputExcelButton(self, self.dataModel)  # 导出文件按钮
        self.editBox = EditBox(self, self.dataModel)  # json输入框
        self.excelTable = ExcelTable(self, self.dataModel)  # 表格
        self.noDataLabel = NoDataLabel(self)    # 无数据标签
        self.statusBar = StatusBar(self.statusBar(), self.dataModel)  # 状态栏

        # 触发更新表格和状态栏
        self.editBox.textChanged.connect(self.excelTable.updateTable)
        self.editBox.textChanged.connect(self.statusBar.updateStatusBar)
        self.importJsonButton.fileSelectedSignal.connect(self.excelTable.updateTable)
        self.importJsonButton.fileSelectedSignal.connect(self.statusBar.updateStatusBar)
        self.formatButton.clicked.connect(self.statusBar.updateStatusBar)
        self.outputExcelButton.clicked.connect(self.statusBar.updateStatusBar)

        # 按钮布局
        buttonLayout = QHBoxLayout()
        buttonLayout.addWidget(self.importJsonButton)
        buttonLayout.addWidget(self.formatButton)
        buttonLayout.addWidget(self.outputExcelButton)

        # 左边布局
        leftLayout = QVBoxLayout()
        leftLayout.addWidget(self.editBox)
        leftLayout.addLayout(buttonLayout)

        # 整体布局
        self.layout = QHBoxLayout(self.centralWidget)
        self.layout.addLayout(leftLayout)
        self.layout.addWidget(self.excelTable)
        self.layout.setStretchFactor(self.excelTable, 1)  # 表格具有拉伸因子，意味拉伸时只有右边会拉伸

        # 默认表格隐藏，只显示noDataLabel
        self.excelTable.hide()
        self.layout.replaceWidget(self.excelTable, self.noDataLabel)


class DataModel:

    def __init__(self):
        self.setDefault()

    def setDefault(self):
        self._jsonData = DEFAULT_JSON_DATA
        self._dataInfo = DEFAULT_DATA_INFO
        self._info = DEFAULT_INFO
        self._fields = DEFAULT_FIELDS

    @property
    def jsonData(self):
        return self._jsonData

    @jsonData.setter
    def jsonData(self, oldJsonData):
        jsonData = [
            {
                key: "true" if val is True else
                "false" if val is False else
                "null" if val is None else
                str(val)
                for key, val in d.items()
            }
            for d in oldJsonData
        ]
        self._jsonData = jsonData

    @property
    def dataInfo(self):
        return self._dataInfo

    @dataInfo.setter
    def dataInfo(self, dataInfo):
        self._dataInfo = dataInfo

    @property
    def info(self):
        return self._info

    @info.setter
    def info(self, info):
        self._info = info

    @property
    def fields(self):
        return self._fields

    @fields.setter
    def fields(self, fields):
        self._fields = fields


class StatusBar:

    def __init__(self, statusBar=None, dataModel=None):
        self.statusBar = statusBar
        self.dataModel = dataModel
        self.initUI()
        self.updateStatusBar()

    def initUI(self):
        self.rightLabel = QLabel()
        self.statusBar.addPermanentWidget(self.rightLabel, -1)

    def updateStatusBar(self):
        self.statusBar.showMessage(self.dataModel.info)
        self.rightLabel.setText(f"共计{self.dataModel.dataInfo[0]}个字段，{self.dataModel.dataInfo[1]}条数据")


class EditBox(QPlainTextEdit):

    def __init__(self, parent=None, dataModel=None):
        super().__init__(parent)
        self.dataModel = dataModel
        self.textChanged.connect(self.updateData)
        self.initUI()

    def initUI(self):
        self.setPlaceholderText(
            '请在此处粘贴Json Array文本\n\n如:\n[{"name": "Tom", "age": 20},\n{"name": "Bob", "age": 18},\n{"name": "Lucy", "age": 19}]')
        self.setTabStopWidth(24)
        self.setContextMenuPolicy(Qt.CustomContextMenu)
        self.customContextMenuRequested.connect(self.onContextMenu)

    def onContextMenu(self, position):
        menu = QMenu(self)
        copyAction = menu.addAction("复制")
        copyAction.setIcon(QIcon(os.path.join(os.path.dirname(__file__), "./icons/copy.png")))
        copyAction.triggered.connect(self.copy)
        cutAction = menu.addAction("剪切")
        cutAction.setIcon(QIcon(os.path.join(os.path.dirname(__file__), "./icons/cut.png")))
        cutAction.triggered.connect(self.cut)
        pasteAction = menu.addAction("粘贴")
        pasteAction.setIcon(QIcon(os.path.join(os.path.dirname(__file__), "./icons/paste.png")))
        pasteAction.triggered.connect(self.paste)
        menu.exec_(self.viewport().mapToGlobal(position))

    def updateData(self):
        if self.toPlainText() == "":  # 用于清空文本后dataModel归为默认值
            self.dataModel.setDefault()
            return
        flag = self.checkJsonArray(self.toPlainText())
        if flag[0] is True:
            self.dataModel.jsonData = json.loads(self.toPlainText())
            maxLine = max(self.dataModel.jsonData, key=lambda obj: len(obj), default=None)  # 返回字段数最多的一行
            if maxLine is not None:
                self.dataModel.fields = list(maxLine.keys())
                self.dataModel.dataInfo = (len(maxLine), len(self.dataModel.jsonData))
            self.dataModel.info = "✅已识别到键入Json Array文本"
        else:
            self.dataModel.info = flag[1]

    def checkJsonArray(self, text):
        try:
            parsedJson = json.loads(text)
            if not isinstance(parsedJson, list):
                return (False, "⚠输入的不是Json Array，最外层应为数组")
            for dictItem in parsedJson:
                if not isinstance(dictItem, dict):
                    return (False, "⚠输入的不是Json Array，内层应为对象")
            return (True, None)
        except Exception as e:
            return (False, f"⚠Json解析出错：{e}")


class NoDataLabel(QLabel):

    def __init__(self, parent=None):
        super().__init__("无数据", parent)
        self.setAlignment(Qt.AlignCenter)
        self.setStyleSheet("QLabel { font-size: 20px; border: 1px solid grey; background: white}")


class ExcelTable(QTableWidget):

    def __init__(self, parent=None, dataModel=None):
        super().__init__(parent)
        self.dataModel = dataModel
        self.initUI()

    def initUI(self):
        self.setEditTriggers(QTableWidget.NoEditTriggers)  # 禁用编辑
        self.setSelectionMode(QAbstractItemView.SingleSelection)  # 单选限制

    def contextMenuEvent(self, event):
        selectedItem = self.selectedItems()
        if selectedItem:
            selectedText = selectedItem[0].text()
            menu = QMenu(self)
            copyAction = menu.addAction('复制')
            copyAction.setIcon(QIcon("./icons/copy.png"))
            copyAction.triggered.connect(self.copyData)
            copyAction.setData(selectedText)
            menu.exec_(event.globalPos())  # 显示菜单
        else:
            print("[Info][Table]未选中单元格")

    def copyData(self):
        sender = self.sender()
        text = sender.data()
        clipboard = QApplication.clipboard()
        clipboard.setText(text)

    def updateTable(self):
        if self.dataModel.jsonData != DEFAULT_JSON_DATA:
            self.parent().parent().noDataLabel.hide()
            self.parent().parent().layout.replaceWidget(self.parent().parent().noDataLabel, self)
            self.show()
            self.setColumnCount(self.dataModel.dataInfo[0])
            self.setRowCount(self.dataModel.dataInfo[1])
            self.setHorizontalHeaderLabels(self.dataModel.fields)

            for row, dictItem in enumerate(self.dataModel.jsonData):
                for col, key in enumerate(self.dataModel.fields):
                    value = dictItem.get(key, "")
                    self.setItem(row, col, QTableWidgetItem(str(value)))
            self.resizeColumnsToContents()
        else:

            self.setColumnCount(0)
            self.setRowCount(0)
            self.parent().parent().noDataLabel.show()
            self.parent().parent().layout.replaceWidget(self, self.parent().parent().noDataLabel)
            self.hide()


class ImportJsonButton(QPushButton):
    fileSelectedSignal = pyqtSignal(str)

    def __init__(self, parent=None, dataModel=None):
        super().__init__("导入json文件", parent)
        self.dataModel = dataModel
        self.setToolTip("支持.json和.txt格式的json标准文件")
        self.clicked.connect(self.handleFile)

    def handleFile(self):
        filename, _ = QFileDialog.getOpenFileName(
            self,
            "导入json文件",
            QStandardPaths.writableLocation(QStandardPaths.DesktopLocation),
            "json文件(*.json;*.txt);;所有文件 (*)")
        if filename:
            flag = self.checkJsonArray(filename)
            if flag[0] is True:
                jsonString = flag[1]
                self.dataModel.jsonData = json.loads(jsonString)
                maxLine = max(self.dataModel.jsonData, key=lambda obj: len(obj), default=None)  # 返回字段数最多的一行
                if maxLine is not None:
                    self.dataModel.fields = list(maxLine.keys())
                    self.dataModel.dataInfo = (len(maxLine), len(self.dataModel.jsonData))
                self.parent().parent().editBox.setPlainText(jsonString)
                self.dataModel.info = f"✅已导入Json Array文件{filename}"
                # 告知ExcelTable和StatusBar文件已经导入，这里用自定义信号的原因是：如果用clicked信号，当check未通过时也会发出信号
                # 当然这里也有解决方案：可以通过在类中定义变量决定是否触发clicked信号
                self.fileSelectedSignal.emit("FileImported")
            else:
                QMessageBox.critical(self, f"无法导入{filename}", f"错误信息：\n\n\n{flag[1]}")
        else:
            print("[Info][IMButton]未选择任何文件")

    def checkJsonArray(self, filename):
        encodings = ['UTF-8', 'ASCII', 'GBK', 'UTF-16', 'UTF-32', 'ISO-8859-1', 'BIG5']
        for encoding in encodings:
            try:
                with open(filename, 'r', encoding=encoding) as file:
                    jsonString = json.dumps(json.load(file), ensure_ascii=False)
                parsedJson = json.loads(jsonString)
                if not isinstance(parsedJson, list):
                    return (False, "导入的不是Json Array，最外层应为数组")
                for dictItem in parsedJson:
                    if not isinstance(dictItem, dict):
                        return (False, "导入的不是Json Array，内层应为对象")
                return (True, jsonString)
            except UnicodeDecodeError:
                pass
            except Exception as e:
                return (False, f"导入文件时出错：{e}")


class FormatButton(QPushButton):

    def __init__(self, parent=None, dataModel=None):
        super().__init__("格式化", parent)
        self.dataModel = dataModel
        self.setToolTip("格式化编辑框中的json文本")
        self.clicked.connect(self.format)

    def format(self):
        flag = self.checkJsonArray(self.parent().parent().editBox.toPlainText())
        if self.parent().parent().editBox.toPlainText() == "":  # 未输入内容
            self.dataModel.info = "⚠请输入Json Array后再进行格式化"
            printError("[Error][FButton]请输入Json Array后再进行格式化")
        elif not flag:  # 输入错误的内容
            self.dataModel.info = "⚠请输入正确格式的Json Array后再进行格式化"
            printError("[Error][FButton]请输入正确格式的Json Array后再进行格式化")
        else:  # 前面过滤了为空和输入错误的情况，那么剩下一种情况就是输入正确
            # 设计indent可选为2、4、None（紧凑为一行），以及一个{}为一行的模式，用一个下拉实现
            tmp = json.dumps(self.dataModel.jsonData, indent=4, ensure_ascii=False)
            self.parent().parent().editBox.setPlainText(tmp)
            self.dataModel.info = "✅已格式化Json Array文本"
            print("[Info][FButton]已格式化Json Array文本")

    def checkJsonArray(self, text):
        try:
            parsedJson = json.loads(text)
            if not isinstance(parsedJson, list):
                return False
            for dictItem in parsedJson:
                if not isinstance(dictItem, dict):
                    return False
            return True
        except Exception as e:
            return False


class OutputExcelButton(QPushButton):

    def __init__(self, parent=None, dataModel=None):
        super().__init__("导出Excel文件", parent)
        self.dataModel = dataModel
        self.setToolTip("支持.xlsx、.xls、.csv和.txt格式的Excel标准文件")
        self.clicked.connect(self.handleFile)
        self.saveShortcut = QShortcut(Qt.ControlModifier + Qt.Key_S, self)
        self.saveShortcut.activated.connect(self.handleFile)

    def handleFile(self):
        if self.dataModel.jsonData == DEFAULT_JSON_DATA:
            QMessageBox.critical(self, f"无法导出", f"错误信息：\n\n\n请输入或者导入有效的Json Array")
            return
        filename, fileFilter = QFileDialog.getSaveFileName(
            self,
            "导出Excel文件",
            QStandardPaths.writableLocation(QStandardPaths.DesktopLocation),
            "Excel文件(*.xlsx);;csv文件(逗号分隔)(*.csv);;txt文件(*.txt)")
        if filename:
            if fileFilter == "Excel文件(*.xlsx)":
                workbook = Workbook(filename)
                worksheet = workbook.add_worksheet()
                textFormat = workbook.add_format({'num_format': '@'})  # 单元格纯文本格式
                headers = self.dataModel.fields
                worksheet.write_row("A1", headers)
                for index, dictItem in enumerate(self.dataModel.jsonData):
                    worksheet.write_row(f"A{index + 2}", dictItem.values(), textFormat)

                worksheet.autofit()
                workbook.close()
                self.dataModel.info = f"✅已导出文件到{filename}"
            elif fileFilter == "csv文件(逗号分隔)(*.csv)" or "txt文件(*.txt)":
                with open(filename, "a") as f:
                    for key in self.dataModel.jsonData[0]:
                        f.write(f'{key},')
                    f.write("\n")
                    for dictItem in self.dataModel.jsonData:
                        for value in dictItem.values():
                            f.write(f'{value},')
                        f.write("\n")
                self.dataModel.info = f"✅已导出文件到{filename}"
        else:
            print("[Info][OButton]未选择任何文件")


def printError(text):
    print(f"\033[1;31m{text}\033[0m")


def test():
    pass


def main():
    import sys
    app = QApplication(sys.argv)
    # translator = QTranslator()
    # if translator.load("qt_zh_CN.qm"):
    #     app.installTranslator(translator)
    jsonToExcel = MainWindow()
    jsonToExcel.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
