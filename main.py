# 主程序
import sys
import fitz
from PyQt5.QtCore import Qt, QSize, QRect
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *


class PDFReader(QMainWindow):

    def __init__(self):
        super(PDFReader, self).__init__()
        self.menubar = self.menuBar()
        self.generateMenuBar()
        self.toolbar = self.addToolBar("工具栏")
        self.generateToolBar()
        layout = QHBoxLayout(self)
        self.toc = QTreeWidget()
        self.file_path = ""
        self.page_num = 0
        self.doc = None
        self.dock = QDockWidget()
        self.generateDockWidget()
        self.addDockWidget(Qt.LeftDockWidgetArea, self.dock)
        self.dock.setVisible(False)
        self.trans_a = 200
        self.trans_b = 200
        self.trans = fitz.Matrix(self.trans_a / 100, self.trans_b / 100).preRotate(0)
        self.scrollarea = QScrollArea(self)
        # self.scrollarea.setMinimumSize(500, 500)
        # temp = QHBoxLayout()
        self.pdfview = QLabel()
        # temp.addWidget(self.pdfview)
        # self.scrollarea.setLayout(temp)
        # self.scrollarea.
        self.scrollarea.setWidget(self.pdfview)
        # self.scrollarea.setMinimumSize(0, 0)
        # self.pdfview.setMinimumSize(500, 500)
        # self.pdfview.set
        self.generatePDFView()
        layout.addWidget(self.scrollarea)
        self.widget = QWidget()
        self.widget.setLayout(layout)
        self.setCentralWidget(self.widget)
        self.setWindowTitle('PDF Reader')
        desktop = QApplication.desktop()
        rect = desktop.availableGeometry()
        self.setGeometry(rect)
        self.setWindowIcon(QIcon('icon/reader.png'))
        # self.setGeometry(100, 100, 1000, 600)
        self.show()

    def generatePDFView(self):
        if not self.file_path or not self.doc:
            return
        pix = self.doc[self.page_num].getPixmap(matrix=self.trans)
        fmt = QImage.Format_RGBA8888 if pix.alpha else QImage.Format_RGB888
        pageImage = QImage(pix.samples, pix.width, pix.height, pix.stride, fmt)
        pixmap = QPixmap()
        pixmap.convertFromImage(pageImage)
        # print(pixmap.size())
        # pixmap.scaled(1000, 1000)
        # print(self.width(), "  ", self.height())
        # self.pdfview.setScaledContents(True)
        self.pdfview.setPixmap(QPixmap(pixmap))
        self.pdfview.resize(pixmap.size())
        # self.scrollarea.setWidget(self.pdfview)
        # print(self.pdfview.size())
        # self.pdfview.adjustSize()

    def generateFile(self):
        file = self.menubar.addMenu('文件')
        file.setFont(QFont("", 13))
        openFile = QAction(QIcon('icon/file.png'), '打开文件', file)
        closeFile = QAction("关闭文件", file)
        saveFile = QAction(QIcon('icon/save.png'), '保存文件', file)
        saveFile.setShortcut('Ctrl+S')

        openFile.triggered.connect(self.openfile)
        closeFile.triggered.connect(self.closefile)
        saveFile.triggered.connect(self.savefile)

        file.addAction(openFile)
        file.addMenu('最近文件')
        file.addAction(saveFile)
        file.addAction(closeFile)

    def generatePage(self):
        page = self.menubar.addMenu('页面')
        page.setFont(QFont("", 13))
        insertPage = QAction(QIcon('icon/insert.png'), '添加页面', page)
        deletePage = QAction(QIcon('icon/delete.png'), '删除当前页面', page)
        extractPage = QAction(QIcon('icon/pdf.png'), '提取pdf页面', page)

        self.pageConnect(deletePage, extractPage, insertPage)

        page.addAction(insertPage)
        page.addAction(deletePage)
        page.addAction(extractPage)

    def generateInfile(self):
        infile = self.menubar.addMenu('导入')
        infile.setFont(QFont("", 13))
        inHTML = QAction(QIcon('icon/html.png'), '导入HTML', infile)
        inPic = QAction(QIcon('icon/pic.png'), '导入图片', infile)
        inDocx = QAction(QIcon('icon/docx.png'), '导入docx', infile)
        inMarkdown = QAction(QIcon('icon/markdown_2.png'), '导入markdown', infile)

        inHTML.triggered.connect(self.inhtml)
        inPic.triggered.connect(self.inpic)
        inDocx.triggered.connect(self.indocx)
        inMarkdown.triggered.connect(self.inmarkdown)

        infile.addAction(inDocx)
        infile.addAction(inPic)
        infile.addAction(inHTML)
        infile.addAction(inMarkdown)

    def generateOutfile(self):
        outfile = self.menubar.addMenu('导出')
        outfile.setFont(QFont("", 13))
        toToC = QAction('导出目录', outfile)
        toPic = QAction('导出为图片', outfile)
        toHTML = QAction(QIcon('icon/html.png'), '导出为HTML', outfile)
        toTXT = QAction(QIcon('icon/txt.png'), '导出为TXT', outfile)
        toDocx = QAction(QIcon('icon/word.png'), '导出为Docx', outfile)

        toToC.triggered.connect(self.totoc)
        toPic.triggered.connect(self.topic)
        self.tofileConnect(toDocx, toHTML, toTXT)

        outfile.addAction(toToC)
        outfile.addAction(toPic)
        outfile.addSeparator()
        outfile.addAction(toHTML)
        outfile.addAction(toTXT)
        outfile.addAction(toDocx)

    def generateShare(self):
        share = self.menubar.addMenu('分享')
        share.setFont(QFont("", 13))
        toKindle = QAction(QIcon('icon/kindle.png'), '发送到kindle', share)
        toQQ = QAction(QIcon('icon/QQ.png'), '分享到QQ', share)
        toWechat = QAction(QIcon('icon/wechat.png'), '分享到微信', share)
        toEmail = QAction(QIcon('icon/email.png'), '通过邮件发送', share)

        self.shareConnect(toEmail, toKindle, toQQ, toWechat)

        share.addAction(toKindle)
        share.addAction(toQQ)
        share.addAction(toWechat)
        share.addAction(toEmail)

    def generateMenuBar(self):
        qss = '''
        QMenuBar{
            min-height: 35px;
            font-size: 22px;
        }
        QMenu::item{

        }'''
        self.menubar.setStyleSheet(qss)
        self.generateFile()
        self.generatePage()
        self.generateInfile()
        self.generateOutfile()
        self.generateShare()

    def generateToolBar(self):
        self.toolbar.setMinimumSize(QSize(200, 200))
        self.toolbar.setIconSize(QSize(100, 100))
        ToC = QAction(QIcon('icon/目录 (5).png'), '目录', self.toolbar)
        openFile = QAction(QIcon('icon/file.png'), '打开文件', self.toolbar)
        saveFile = QAction(QIcon('icon/Save (3).png'), '保存文件', self.toolbar)
        prePage = QAction(QIcon('icon/分页 下一页 (1).png'), '上一页', self.toolbar)
        nextPage = QAction(QIcon('icon/分页 下一页.png'), '下一页', self.toolbar)
        insertPage = QAction(QIcon('icon/insert.png'), '添加页面', self.toolbar)
        deletePage = QAction(QIcon('icon/delete.png'), '删除当前页面', self.toolbar)
        extractPage = QAction(QIcon('icon/pdf.png'), '提取pdf页面', self.toolbar)
        enlargePage = QAction(QIcon('icon/放大 (1).png'), '放大', self.toolbar)
        shrinkPage = QAction(QIcon('icon/缩小.png'), '缩小', self.toolbar)
        toHTML = QAction(QIcon('icon/html (3).png'), '导出为HTML', self.toolbar)
        toTXT = QAction(QIcon('icon/txt.png'), '导出为TXT', self.toolbar)
        toDocx = QAction(QIcon('icon/word.png'), '导出为Docx', self.toolbar)
        toKindle = QAction(QIcon('icon/kindle.png'), '发送到kindle', self.toolbar)
        toQQ = QAction(QIcon('icon/QQ.png'), '分享到QQ', self.toolbar)
        toWechat = QAction(QIcon('icon/wechat.png'), '分享到微信', self.toolbar)
        toEmail = QAction(QIcon('icon/email.png'), '通过邮件发送', self.toolbar)

        nextPage.setShortcut(Qt.Key_Right)
        prePage.setShortcut(Qt.Key_Left)

        self.shareConnect(toEmail, toKindle, toQQ, toWechat)
        openFile.triggered.connect(self.openfile)
        ToC.triggered.connect(self.handleDock)
        saveFile.triggered.connect(self.savefile)
        prePage.triggered.connect(self.prepage)
        nextPage.triggered.connect(self.nextpage)
        enlargePage.triggered.connect(self.enlargepage)
        shrinkPage.triggered.connect(self.shrinkpage)
        self.tofileConnect(toDocx, toHTML, toTXT)
        self.pageConnect(deletePage, extractPage, insertPage)

        self.toolbar.addActions([ToC])
        self.toolbar.addSeparator()
        self.toolbar.addActions([openFile, saveFile])
        self.toolbar.addSeparator()
        self.toolbar.addActions([prePage, nextPage])
        self.toolbar.addSeparator()
        self.toolbar.addActions([insertPage, deletePage, extractPage])
        self.toolbar.addSeparator()
        self.toolbar.addActions([enlargePage, shrinkPage])
        self.toolbar.addSeparator()
        self.toolbar.addActions([toHTML, toTXT, toDocx])
        self.toolbar.addSeparator()
        self.toolbar.addActions([toKindle, toQQ, toWechat, toEmail])

    def generateDockWidget(self):
        if not self.file_path:
            return
        self.dock.setWidget(self.toc)
        self.generateTreeWidget()

    def generateTreeWidget(self):
        if not self.doc:
            return
        self.toc.setColumnCount(1)
        self.toc.setHeaderLabels(['目录'])
        # tree.setMinimumSize(500, 500)
        self.toc.setWindowTitle('目录')
        toc = self.doc.getToC()
        nodelist = [self.toc]
        floorlist = [0]
        tempdict = {}
        if not toc:
            return tempdict
        first = True
        for line in toc:
            floor, title, page = line
            if first:
                node = QTreeWidgetItem(self.toc)
                node.setText(0, title)
                nodelist.append(node)
                floorlist.append(floor)
                tempdict[title] = page
                first = False
            else:
                while floorlist[-1] >= floor:
                    nodelist.pop()
                    floorlist.pop()
                node = QTreeWidgetItem(nodelist[-1])
                node.setText(0, title)
                nodelist.append(node)
                floorlist.append(floor)
                tempdict[title] = page
        return tempdict

    def tofileConnect(self, toDocx, toHTML, toTXT):
        toHTML.triggered.connect(self.tohtml)
        toTXT.triggered.connect(self.totxt)
        toDocx.triggered.connect(self.todocx)

    def pageConnect(self, deletePage, extractPage, insertPage):
        insertPage.triggered.connect(self.insertpage)
        deletePage.triggered.connect(self.deletepage)
        extractPage.triggered.connect(self.extractpage)

    def shareConnect(self, toEmail, toKindle, toQQ, toWechat):
        toKindle.triggered.connect(self.tokindle)
        toQQ.triggered.connect(self.toqq)
        toWechat.triggered.connect(self.towechat)
        toEmail.triggered.connect(self.toemail)

    def handleDock(self):
        try:
            if self.dock.isVisible():
                self.dock.setVisible(False)
            else:
                self.dock.setVisible(True)
        except AttributeError:
            pass

    def openfile(self):
        fDialog = QFileDialog()
        self.file_path, _ = fDialog.getOpenFileName(self, "打开文件", ".", 'PDF file (*.pdf)')
        self.toc.clear()
        self.page_num = 0
        self.getDoc()
        self.generateDockWidget()
        self.generatePDFView()

    def getDoc(self):
        if self.file_path:
            self.doc = fitz.open(self.file_path)

    def closefile(self):
        pass

    def savefile(self):
        pass

    def prepage(self):
        self.page_num -= 1
        self.generatePDFView()

    def nextpage(self):
        self.page_num += 1
        self.scrollarea.verticalScrollBar().setValue(0)
        # scrollBar.setValue(200)
        self.generatePDFView()

    def enlargepage(self):
        self.trans_a += 5
        self.trans_b += 5
        self.trans = fitz.Matrix(self.trans_a / 100, self.trans_b / 100).preRotate(0)
        self.generatePDFView()

    def shrinkpage(self):
        self.trans_a -= 5
        self.trans_b -= 5
        self.trans = fitz.Matrix(self.trans_a / 100, self.trans_b / 100).preRotate(0)
        self.generatePDFView()

    def insertpage(self):
        pass

    def deletepage(self):
        pass

    def extractpage(self):
        pass

    def inhtml(self):
        pass

    def inmarkdown(self):
        pass

    def indocx(self):
        pass

    def inpic(self):
        pass

    def tohtml(self):
        pass

    def totxt(self):
        pass

    def totoc(self):
        pass

    def topic(self):
        pass

    def todocx(self):
        pass

    def tokindle(self):
        pass

    def toqq(self):
        pass

    def towechat(self):
        pass

    def toemail(self):
        pass


if __name__ == '__main__':
    app = QApplication(sys.argv)
    reader = PDFReader()
    sys.exit(app.exec_())
