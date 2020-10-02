import os
import re
import time

import fitz
from PyQt5.QtCore import pyqtSignal, QUrl
from PyQt5.QtGui import QDesktopServices, QFont
from PyQt5.QtWidgets import QDialog, QVBoxLayout, QPushButton, QFileDialog, QInputDialog, QLabel, QLineEdit, \
    QApplication, QMessageBox, QComboBox

from convert import picsToPdf, htmlToPdf
from myemail import sendMailByOutLook
from mythreads import outEmailThread, convertThread, EmailThread


class InsertDialog(QDialog):
    picSignal = pyqtSignal(list)
    pdfSignal = pyqtSignal(str, int, int)

    def __init__(self, parent=None):
        super(InsertDialog, self).__init__(parent)
        layout = QVBoxLayout()
        self.inpicBtn = QPushButton("插入单张图片")
        self.inpicsBtn = QPushButton("插入多张图片")
        self.infilepicBtn = QPushButton("插入目录下的所有图片")
        self.inpdfBtn = QPushButton("插入另一个pdf的页面")
        layout.addWidget(self.inpicBtn)
        layout.addWidget(self.inpicsBtn)
        layout.addWidget(self.infilepicBtn)
        layout.addWidget(self.inpdfBtn)
        self.setLayout(layout)
        self.setWindowTitle("插入页面")
        self.setMinimumSize(400, 200)
        self.inpicBtn.clicked.connect(self.inpic)
        self.inpicsBtn.clicked.connect(self.inpics)
        self.infilepicBtn.clicked.connect(self.infilepic)
        self.inpdfBtn.clicked.connect(self.inpdf)

    def inpic(self):
        fDialog = QFileDialog()
        pic_path, _ = fDialog.getOpenFileName(self, "选择单张图片", ".", 'Image file (*.png *.jpg *.jpeg)')
        if pic_path:
            pic_list = [pic_path]
            self.picSignal.emit(pic_list)

    def inpics(self):
        fDialog = QFileDialog()
        pic_paths, _ = fDialog.getOpenFileNames(self, "选择多张图片", ".", 'Image file (*.png *.jpg *.jpeg)')
        self.picSignal.emit(pic_paths)

    def infilepic(self):
        fDialog = QFileDialog()
        path = fDialog.getExistingDirectory(self, "选择文件夹", ".")
        if path:
            files = os.listdir(path)
            pic_list = [os.path.join(path, file) for file in files if file.endswith((".png", ".jpg", ".jpeg"))]
            self.picSignal.emit(pic_list)

    def inpdf(self):
        fDialog = QFileDialog()
        file_path, _ = fDialog.getOpenFileName(self, "选择pdf文件", ".", 'PDF file (*.pdf)')
        if file_path:
            doc = fitz.open(file_path)
            page_num = doc.pageCount
            start, ok = QInputDialog.getInt(self, "选择开始页面", "输入开始页面(1-{})".format(page_num), min=1, max=page_num)
            if ok:
                end, bok = QInputDialog.getInt(self, "选择结束页面", "输入结束页面({}-{})".format(start, page_num), min=start,
                                               max=page_num)
                if bok:
                    self.pdfSignal.emit(file_path, start, end)


class EmailToKindleDialog(QDialog):
    addressSignal = pyqtSignal(str)

    def __init__(self, parent=None, emailList=None):
        super(EmailToKindleDialog, self).__init__(parent)
        layout = QVBoxLayout()
        self.noteLabel = QLabel("<b style='color: red'>请确保您已将2587354021@qq.com加入信任邮箱</b>")
        self.linkButton = QPushButton("前往亚马逊进行设置")
        self.tsLabel = QLabel("输入您的kindle邮箱")
        self.tsLabel.setFont(QFont("", 14))
        self.sendBtn = QPushButton("发送")
        self.addressComboBox = QComboBox()
        self.addressComboBox.setEditable(True)
        self.addressComboBox.addItems(emailList)
        self.linkButton.clicked.connect(self.openLink)
        self.sendBtn.clicked.connect(self.sendAddr)
        # self.sendBtn.setFocusPolicy(Qt.StrongFocus)
        layout.addWidget(self.tsLabel)
        layout.addWidget(self.addressComboBox)
        layout.addWidget(self.sendBtn)
        layout.addWidget(self.noteLabel)
        layout.addWidget(self.linkButton)
        self.setLayout(layout)

    def openLink(self):
        clipboard = QApplication.clipboard()
        clipboard.setText("2587354021@qq.com")
        QMessageBox.about(self, "提示", "邮箱已复制到剪贴板")
        QDesktopServices.openUrl(QUrl('https://www.amazon.cn/hz/mycd/myx#/home/settings/payment'))

    def sendAddr(self):
        address = self.addressComboBox.currentText()
        if address:
            self.addressSignal.emit(address)


class EmailToOthersDialog(QDialog):
    emailSignal = pyqtSignal(int, int)

    def __init__(self, parent=None, file_path=""):
        super(EmailToOthersDialog, self).__init__(parent)
        self.file_path = file_path
        layout = QVBoxLayout()
        self.sendToOne = QPushButton("输入单个邮箱")
        self.sendToMany = QPushButton("发送给多个邮箱")
        self.noteLabel = QLabel("发送给多个邮箱请将邮箱号写入txt文件中，一行一个")
        self.sendToOne.clicked.connect(self.sendOne)
        self.sendToMany.clicked.connect(self.sendMany)
        layout.addWidget(self.sendToOne)
        layout.addWidget(self.sendToMany)
        layout.addWidget(self.noteLabel)
        self.setLayout(layout)

    def sendOne(self):
        address, ok = QInputDialog.getText(self, "输入邮箱", "请输入邮箱")
        if ok:
            t = outEmailThread(sendMailByOutLook, (self.file_path, [address]))
            t.finishSignal.connect(self.onFinishThread)
            t.start()
            time.sleep(1)  # 记住开线程就要睡一会

    def onFinishThread(self, suc, fail):
        self.emailSignal.emit(suc, fail)

    def sendMany(self):
        txtfile, _ = QFileDialog.getOpenFileName(self, "选择文件", ".", "txt file(*.txt)")
        if not txtfile:
            return
        f = open(txtfile, 'r')
        address_list = [line.strip() for line in f]
        t = outEmailThread(sendMailByOutLook, (self.file_path, address_list))
        t.finishSignal.connect(self.onFinishThread)
        t.start()
        time.sleep(1)


class InPicDialog(QDialog):
    finishSignal = pyqtSignal(str)

    def __init__(self, parent=None):
        super(InPicDialog, self).__init__(parent)
        self.picBtn = QPushButton("手动选择图片")
        self.fileBtn = QPushButton("选择目录下的所有图片")
        self.picBtn.clicked.connect(self.onPic)
        self.fileBtn.clicked.connect(self.onFile)
        layout = QVBoxLayout()
        layout.addWidget(self.picBtn)
        layout.addWidget(self.fileBtn)
        self.toname = None
        self.setLayout(layout)

    def onPic(self):
        filenames, _ = QFileDialog.getOpenFileNames(self, "选择文件", ".", "Image file(*.jpg *.png *.jpeg)")
        if filenames:
            self.toname, _ = QFileDialog.getSaveFileName(self, "保存文件", ".", "PDF file(*.pdf)")
            if self.toname:
                t = convertThread(picsToPdf, (filenames, self.toname))
                t.finishSignal.connect(self.handleSig)
                t.start()
                time.sleep(2)

    def handleSig(self):
        self.finishSignal.emit(self.toname)

    def onFile(self):
        path = QFileDialog.getExistingDirectory(self, "选择文件夹", ".")
        filenames = [os.path.join(path, filename) for filename in os.listdir(path) if
                     filename.endswith(('.png', '.jpg', 'jpeg'))]
        self.toname, _ = QFileDialog.getSaveFileName(self, "保存文件", ".", "PDF File(*.pdf)")
        if self.toname:
            t = convertThread(picsToPdf, (filenames, self.toname))
            t.finishSignal.connect(self.handleSig)
            t.start()
            time.sleep(2)


class pdfkitNoteDialog(QDialog):
    def __init__(self, parent=None):
        super(pdfkitNoteDialog, self).__init__(parent)
        self.label1 = QLabel("非常抱歉，转换失败")
        self.label2 = QLabel("请确保您已成功安装wkhtmltopdf，并将其路径添加到系统环境变量中")
        self.toWeb = QPushButton("点击前往下载wkhtmltopdf")
        self.toPath = QPushButton("点击前往添加系统环境变量")
        layout = QVBoxLayout()
        layout.addWidget(self.label1)
        layout.addWidget(self.label2)
        layout.addWidget(self.toWeb)
        layout.addWidget(self.toPath)
        self.toWeb.clicked.connect(self.onToWeb)
        self.toPath.clicked.connect(self.onToPath)
        self.setLayout(layout)

    def onToWeb(self):
        QDesktopServices.openUrl(QUrl('https://wkhtmltopdf.org/downloads.html'))

    def onToPath(self):
        os.system('sysdm.cpl')
        QMessageBox.about(self, "提醒", "请切换到‘高级’标签页，摁下键盘的‘N’键，找到系统变量的PATH栏，然后将wkhtmltopdf的安装路径添加进去")


class inHtmlDialog(QDialog):
    finishSignal = pyqtSignal(str)

    def __init__(self, parent=None):
        super(inHtmlDialog, self).__init__(parent)
        self.noteLabel = QLabel("请输入您想要转换的网址")
        self.urlInput = QLineEdit()
        self.conBtn = QPushButton("转换")
        self.filename = ""
        layout = QVBoxLayout()
        layout.addWidget(self.noteLabel)
        layout.addWidget(self.urlInput)
        layout.addWidget(self.conBtn)
        self.conBtn.clicked.connect(self.onConvert)
        self.setLayout(layout)

    def onConvert(self):
        url = self.urlInput.text()
        if not re.match(r'https?://(?:[a-zA-Z0-9$-_@.&+]|[!*\\,]|(?:%[0-9a-fA-F][0-9a-fA-F]))+', url):
            box = QMessageBox(QMessageBox.Warning, "警告", "请输入合法的网址")
            box.setStandardButtons(QMessageBox.Yes | QMessageBox.Retry)
            box.button(QMessageBox.Yes).setText("修改原网址")
            box.button(QMessageBox.Retry).setText("重新输入网址")
            box.setDefaultButton(QMessageBox.Yes)
            box.buttonClicked.connect(self.handleBut)
            box.exec_()
            return

        self.filename, _ = QFileDialog.getSaveFileName(self, "保存文件", ".", "PDF file(*.pdf)")
        if self.filename:
            t = EmailThread(htmlToPdf, (url, self.filename))
            t.finishSignal.connect(self.onFinish)
            t.start()
            time.sleep(1)

    def handleBut(self, btn):
        if btn.text() == "重新输入网址":
            self.urlInput.clear()

    def onFinish(self, success):
        if success:
            self.finishSignal.emit(self.filename)
        else:
            dig = pdfkitNoteDialog(self)
            dig.show()
