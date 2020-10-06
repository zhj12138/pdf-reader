from PyQt5.QtGui import QFont
# from PyQt5.QtWebEngineWidgets import QWebEngineView
from PyQt5.QtWidgets import QSplitter, QPlainTextEdit

# from mdConvert import md2html


class MdEditor(QSplitter):
    def __init__(self):
        super(MdEditor, self).__init__()
        self.writePart = QPlainTextEdit()
        # self.showPart = QWebEngineView()
        # self.showPart.setZoomFactor(1.5)
        self.writePart.setFont(QFont("", 13))
        # self.showPart.setMinimumWidth(20)
        # self.writePart.textChanged.connect(self.updateShowPart)
        self.addWidget(self.writePart)
        # self.addWidget(self.showPart)

    # def updateShowPart(self):
    #     text = self.writePart.toPlainText()
    #     html = md2html(text)
    #     self.showPart.setHtml(html)
