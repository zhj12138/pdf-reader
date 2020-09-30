from PyQt5.QtCore import pyqtSignal, QUrl, Qt, QMimeData
from PyQt5.QtGui import QDesktopServices, QFont
from PyQt5.QtWidgets import QDialog, QVBoxLayout, QPushButton, QFileDialog, QInputDialog, QLabel, QLineEdit, \
    QApplication, QMessageBox
import os
import fitz
from myemail import sendMailByOutLook
from mythreads import outEmailThread
import time


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
                end, bok = QInputDialog.getInt(self, "选择结束页面", "输入结束页面({}-{})".format(start, page_num), min=start, max=page_num)
                if bok:
                    self.pdfSignal.emit(file_path, start, end)


class EmailToKindleDialog(QDialog):
    addressSignal = pyqtSignal(str)

    def __init__(self, parent=None):
        super(EmailToKindleDialog, self).__init__(parent)
        layout = QVBoxLayout()
        self.noteLabel = QLabel("<b style='color: red'>请确保您已将2587354021@qq.com加入信任邮箱</b>")
        self.linkButton = QPushButton("前往亚马逊进行设置")
        self.tsLabel = QLabel("输入您的kindle邮箱")
        self.tsLabel.setFont(QFont("", 14))
        self.sendBtn = QPushButton("发送")
        self.addressInput = QLineEdit()
        self.linkButton.clicked.connect(self.openLink)
        self.sendBtn.clicked.connect(self.sendAddr)
        # self.sendBtn.setFocusPolicy(Qt.StrongFocus)
        layout.addWidget(self.tsLabel)
        layout.addWidget(self.addressInput)
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
        address = self.addressInput.text()
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


# class ToQQDialog(QDialog):
#     def __init__(self, parent=None, file_path=""):
#         super(ToQQDialog, self).__init__(parent)
#         self.file_path = file_path
#         layout = QVBoxLayout()
#         self.toPhone = QPushButton("发送到我的手机")
#         self.toOthers = QPushButton("前往QQ手动选择好友")
#         self.toPhone.clicked.connect(self.onToPhone)
#         self.toOthers.clicked.connect(self.onToOthers)
#         layout.addWidget(self.toPhone)
#         layout.addWidget(self.toOthers)
#         self.setLayout(layout)
#
#     def onToPhone(self):
#         # 模拟Ctrl+Alt+Z打开QQ窗口
#         CtrlAltZ()
#         # 设置剪贴板
#         # time.sleep(0.5)
#         setClipText("我的手机")
#         print('copied')
#         # time.sleep(0.5)
#         # 模拟Ctrl+V
#         # time.sleep(1)
#         CtrlV()
#         print("ctrl+v")
#         # 模拟enter
#         # time.sleep(3)
#         Enter()
#         print('enter')
#         # 设置剪贴板
#         # copyFile(self.file_path)
#         # time.sleep(1)
#         # print("copied")
#         # # 模拟Ctrl+V
#         # CtrlV()
#         # print("to")
#         # time.sleep(1)
#         # # 模拟enter键
#         # Enter()
#         # print("enter2")
#
#     def onToOthers(self):
#         # 设置剪贴板
#         copyFile(self.file_path)
#         QMessageBox.about(self, "提示", "文件已复制到剪贴板")
#         # 模拟Ctrl+Alt+Z打开QQ窗口
#         CtrlAltZ()








