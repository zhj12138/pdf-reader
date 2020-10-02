# 主程序
import sys
import fitz
from PyQt5.QtCore import Qt, QSize, QRect
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from mydialogs import *
from myemail import email_to
from mythreads import EmailThread
import time
from Vkeyboard import *
from convert import *
from mydatabase import MyDb


class PDFReader(QMainWindow):

    def __init__(self):
        super(PDFReader, self).__init__()
        self.db = MyDb()  # 数据库
        self.menubar = self.menuBar()
        self.recentfile = None
        self.generateMenuBar()
        self.generateRecentMenu()
        self.toolbar = self.addToolBar("工具栏")
        self.generateToolBar()
        layout = QHBoxLayout(self)
        self.toc = QTreeWidget()
        self.file_path = ""
        self.page_num = 0
        self.doc = None
        self.book_open = False
        self.dock = QDockWidget()
        self.generateDockWidget()
        self.addDockWidget(Qt.LeftDockWidgetArea, self.dock)
        self.dock.setVisible(False)
        self.trans_a = 200
        self.trans_b = 200
        self.trans = fitz.Matrix(self.trans_a / 100, self.trans_b / 100).preRotate(0)
        self.scrollarea = QScrollArea(self)
        self.pdfview = QLabel()
        self.tocDict = {}
        self.scrollarea.setWidget(self.pdfview)
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
        self.pdfview.setPixmap(QPixmap(pixmap))
        self.pdfview.resize(pixmap.size())

    def generateFile(self):
        file = self.menubar.addMenu('文件')
        file.setFont(QFont("", 13))
        openFile = QAction(QIcon('icon/file.png'), '打开文件', file)
        closeFile = QAction("关闭文件", file)
        saveFile = QAction(QIcon('icon/save.png'), '保存文件', file)
        saveFile.setShortcut('Ctrl+S')

        openFile.triggered.connect(self.onOpen)
        closeFile.triggered.connect(self.onClose)
        saveFile.triggered.connect(self.onSave)

        file.addAction(openFile)
        self.recentfile = file.addMenu('最近文件')
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

    def generateRecentMenu(self):
        self.recentfile.clear()
        fileList = self.db.getAllRencentFile()
        sortlist = sorted(fileList, key=lambda d: d.opentime, reverse=True)
        actionList = []
        for file in sortlist:
            action = QAction(file.path, self.recentfile)
            self.recentfile.addAction(action)
            actionList.append(action)
        for action in actionList:
            action.triggered.connect(lambda: self.open_file(action.text()))

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
        turnPage = QAction(QIcon('icon/跳转.png'), '跳转', self.toolbar)
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
        openFile.triggered.connect(self.onOpen)
        ToC.triggered.connect(self.onDock)
        saveFile.triggered.connect(self.onSave)
        prePage.triggered.connect(self.onPrepage)
        nextPage.triggered.connect(self.nextpage)
        turnPage.triggered.connect(self.turnpage)
        enlargePage.triggered.connect(self.enlargepage)
        shrinkPage.triggered.connect(self.shrinkpage)
        self.tofileConnect(toDocx, toHTML, toTXT)
        self.pageConnect(deletePage, extractPage, insertPage)

        self.toolbar.addActions([ToC])
        self.toolbar.addSeparator()
        self.toolbar.addActions([openFile, saveFile])
        self.toolbar.addSeparator()
        self.toolbar.addActions([prePage, nextPage, turnPage])
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
        self.tocDict = tempdict
        self.toc.clicked.connect(self.bookmark_jump)

    def bookmark_jump(self, index):
        item = self.toc.currentItem()
        self.page_num = self.tocDict[item.text(0)] - 1
        self.updatePdfView()

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

    def onDock(self):
        try:
            if self.dock.isVisible():
                self.dock.setVisible(False)
            else:
                self.dock.setVisible(True)
        except AttributeError:
            pass

    def onOpen(self):
        fDialog = QFileDialog()
        filename, _ = fDialog.getOpenFileName(self, "打开文件", ".", 'PDF file (*.pdf)')
        self.open_file(filename)

    def open_file(self, filename):
        if not filename:
            return
        if not self.db.fileInDB(filename):
            self.db.addRecentFile(filename)
        else:
            self.db.updateRecentFile(filename)
        if os.path.exists(filename):
            self.file_path = filename
        else:
            QMessageBox.about(self, "提醒", "文件不存在")
            self.db.deleteRecentFile(filename)
        self.toc.clear()
        self.page_num = 0
        self.book_open = True
        self.generateRecentMenu()
        self.getDoc()
        self.generateDockWidget()
        self.generatePDFView()

    def getDoc(self):
        if self.file_path:
            self.doc = fitz.open(self.file_path)

    def onClose(self):
        self.file_path = ""
        self.book_open = False
        self.toc.clear()
        self.pdfview.clear()
        self.getDoc()
        self.generatePDFView()
        self.generateDockWidget()

    def onSave(self):
        if not self.book_open:
            return
        self.doc.save(self.doc.name, incremental=True, encryption=fitz.PDF_ENCRYPT_KEEP)

    def onPrepage(self):
        self.page_num -= 1
        if self.page_num < 0:
            self.page_num += self.doc.pageCount
        self.updatePdfView()

    def updatePdfView(self):
        self.scrollarea.verticalScrollBar().setValue(0)
        self.generatePDFView()

    def nextpage(self):
        self.page_num += 1
        if self.page_num >= self.doc.pageCount:
            self.page_num -= self.doc.pageCount
        self.updatePdfView()

    def turnpage(self):
        if not self.book_open:
            return
        allpages = self.doc.pageCount
        print(allpages)
        page, ok = QInputDialog.getInt(self, "选择页面", "输入目标页面({}-{})".format(1, allpages), min=1, max=allpages)
        if ok:
            self.page_num = page - 1
            self.updatePdfView()

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
        if not self.book_open:
            return
        dig = InsertDialog(self)
        dig.picSignal.connect(self.insertpage_pic)
        dig.pdfSignal.connect(self.insertpage_pdf)
        dig.show()

    def insertpage_pic(self, pic_list):
        if not self.book_open:
            return
        for pic in pic_list:
            img = fitz.open(pic)
            rect = img[0].rect  # pic dimension
            pdfbytes = img.convertToPDF()  # make a PDF stream
            img.close()  # no longer needed
            imgPDF = fitz.open("pdf", pdfbytes)  # open stream as PDF
            page = self.doc.newPage(self.page_num, width=rect.width,  # new page with ...
                                    height=rect.height)  # pic dimension
            page.showPDFpage(rect, imgPDF, 0)  # image fills the page
        self.generatePDFView()

    def insertpage_pdf(self, file_path, start, end):
        if not self.book_open:
            return
        doc2 = fitz.open(file_path)
        self.doc.insertPDF(doc2, from_page=start - 1, to_page=end - 1, start_at=self.page_num)
        self.generatePDFView()

    def deletepage(self):
        if not self.book_open:
            return
        self.doc.deletePage(self.page_num)
        self.updatePdfView()

    def extractpage(self):
        if not self.book_open:
            return
        allpages = self.doc.pageCount
        start, ok = QInputDialog.getInt(self, "选择开始页面", "输入开始页面(1-{})".format(allpages), min=1, max=allpages)
        if ok:
            end, ok = QInputDialog.getInt(self, "选择结束页面", "输入结束页面({}-{})".format(start, allpages), min=start,
                                          max=allpages)
            if ok:
                file_name, _ = QFileDialog.getSaveFileName(self, "保存文件", ".", "PDF File(*.pdf)")
                if file_name:
                    doc2 = fitz.open()
                    doc2.insertPDF(self.doc, from_page=start - 1, to_page=end - 1)
                    doc2.save(file_name)

    def inhtml(self):
        dig = inHtmlDialog(self)
        dig.finishSignal.connect(self.open_file)
        dig.show()

    def inmarkdown(self):
        filename, _ = QFileDialog.getOpenFileName(self, "选择文件", ".", "markdown file(*.md)")
        if filename:
            toname, _ = QFileDialog.getSaveFileName(self, "保存文件", ".", "pdf file(*.pdf)")
            if toname:
                t = convertThread(mdToPdf, (filename, toname))
                t.finishSignal.connect(lambda: self.open_file(toname))
                t.start()
                time.sleep(1)

    def indocx(self):
        filename, _ = QFileDialog.getOpenFileName(self, "选择文件", ".", "docx file(*.docx)")
        if filename:
            toname, _ = QFileDialog.getSaveFileName(self, "保存文件", ".", "pdf file(*.pdf)")
            if toname:
                docxToPdf(filename, toname)
                self.open_file(toname)

    def inpic(self):
        dig = InPicDialog(self)
        dig.finishSignal.connect(self.open_file)
        dig.show()

    def tohtml(self):
        if not self.book_open:
            return
        toname, _ = QFileDialog.getSaveFileName(self, "保存文件", ".", "html file(*.html)")
        if toname:
            t = convertThread(pdfToHtmlorTxt, (self.file_path, toname, "html"))
            t.finishSignal.connect(lambda: self.openFileNote(toname))
            t.start()
            time.sleep(1)

    def totxt(self):
        if not self.book_open:
            return
        toname, _ = QFileDialog.getSaveFileName(self, "保存文件", ".", "txt file(*.txt)")
        if toname:
            t = convertThread(pdfToHtmlorTxt, (self.file_path, toname, "text"))
            t.finishSignal.connect(lambda: self.openFileNote(toname))
            t.start()
            time.sleep(1)

    def totoc(self):
        if not self.book_open:
            return
        toname, _ = QFileDialog.getSaveFileName(self, "保存文件", ".", "markdown file(*.md)")
        if toname:
            t = convertThread(tocToMd, (self.file_path, toname))
            t.finishSignal.connect(lambda: self.openFileNote(toname))
            t.start()
            time.sleep(1)

    def openFileNote(self, file_path):
        ret = QMessageBox.question(self, '提示', "转换成功，是否打开文件", QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
        if ret == QMessageBox.Yes:
            os.startfile(file_path)

    def topic(self):
        if not self.book_open:
            return
        file = QFileDialog.getExistingDirectory(self, "", ".")
        if file:
            y = convertThread(pdfToImg, (self.file_path, file))
            y.start()
            time.sleep(1)

    def todocx(self):
        if not self.book_open:
            return
        toname, _ = QFileDialog.getSaveFileName(self, "保存文件", ".", "docx file(*.docx)")
        if toname:
            t = convertThread(pdfToDocx, (self.file_path, toname))
            t.finishSignal.connect(lambda: self.openFileNote(toname))
            t.start()
            time.sleep(1)

    def tokindle(self):
        if not self.book_open:
            return
        emailList = self.db.getAllKindleMail()
        dig = EmailToKindleDialog(self, emailList)
        dig.addressSignal.connect(self.sendMail)
        dig.show()

    def sendMail(self, address):
        if not self.db.mailInDB(address):
            self.db.addKindleMail(address)
        t = EmailThread(email_to, (self.file_path, address))
        t.finishSignal.connect(self.finish_mail)
        t.start()
        time.sleep(0.8)

    def finish_mail(self, success):
        if success:
            QMessageBox.about(self, "提示", "发送成功")
        else:
            QMessageBox.about(self, "提示", "抱歉，发送失败，请检查邮箱后再次发送")

    def toqq(self):
        if not self.book_open:
            return
        copyFile(self.file_path)
        QMessageBox.about(self, "提示", "文件已复制到剪贴板")
        CtrlAltZ()

    def towechat(self):
        copyFile(self.file_path)
        QMessageBox.about(self, "提示", "文件已复制到剪贴板")
        CtrlAltW()

    def toemail(self):
        if not self.book_open:
            return
        dig = EmailToOthersDialog(self, self.file_path)
        dig.emailSignal.connect(self.onToemail)
        dig.show()

    def onToemail(self, suc, fail):
        QMessageBox.about(self, "提示", "{}个成功，{}个失败".format(suc, fail))


if __name__ == '__main__':
    app = QApplication(sys.argv)
    reader = PDFReader()
    sys.exit(app.exec_())
