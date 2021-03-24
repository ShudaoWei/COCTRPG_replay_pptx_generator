# import sys
 
# #这里我们提供必要的引用。基本控件位于pyqt5.qtwidgets模块中。
# from PyQt5.QtWidgets import (QWidget, QToolTip, 
#     QMessageBox, QPushButton, QApplication, QLabel, QLineEdit, 
#     QTextEdit, QGridLayout)
# from PyQt5.QtGui import QFont   
# from PyQt5.QtGui import QIcon
# from PyQt5.QtCore import QCoreApplication

# class Example(QWidget):
#     def __init__(self):
#         super().__init__()
        
#         self.initUI() #界面绘制交给InitUi方法
        
#         # 需要填写的内容：
#         # 玩家列表
#         # 角色列表存储fp
#         # 读取的文件txt fp
#         # 读取的txt 格式（qq\ldn\colored）
#         # 读取的txt 编码格式(尝试获取)
#         # 读取的ppt模板 fp
#         # 读取的ppt模板 文本框内容判断
#         # 生成文件存储路径 fp
#         # 生成文件存储名称
#         # 是否修改对话框字体 boolean
#         # 对话框字体hardcoded
#         #   -字体、 -字号、 -加粗
        
#     def initUI(self):

#         char_fp = QLabel('角色列表存储路径：')
#         txt_fp = QLabel('txt文件路径：')
#         review = QLabel('txt文件格式：')

#         #设置窗口的位置和大小
#         self.setGeometry(100, 100, 550, 450)  
#         #设置窗口的标题
#         self.setWindowTitle('Icon')
#         #设置窗口的图标，引用当前目录下的web.png图片
#         self.setWindowIcon(QIcon('web.png'))
        
#         #这种静态的方法设置一个用于显示工具提示的字体。我们使用10px滑体字体。
#         QToolTip.setFont(QFont('SansSerif', 10))
#         #创建一个提示，我们称之为settooltip()方法。我们可以使用丰富的文本格式
#         # self.setToolTip('This is a <b>QWidget</b> widget')

#          #创建一个PushButton并为他设置一个tooltip
#         qbtn = QPushButton('关闭窗口', self)
#         qbtn.clicked.connect(QCoreApplication.instance().quit)
        
#         #btn.sizeHint()显示默认尺寸
#         qbtn.resize(qbtn.sizeHint())
        
#         #移动窗口的位置
#         qbtn.move(50, 50)
#     def closeEvent(self, event):
        
#         # 最后那个QMessage.No是默认选项
#         reply = QMessageBox.question(self, '确认',
#             "确定要退出吗？", QMessageBox.Yes | 
#             QMessageBox.No, QMessageBox.No)
 
#         if reply == QMessageBox.Yes:
#             event.accept()
#         else:
#             event.ignore()   

# if __name__ == '__main__':
#     #每一pyqt5应用程序必须创建一个应用程序对象。sys.argv参数是一个列表，从命令行输入参数。
#     app = QApplication(sys.argv)
#     #QWidget部件是pyqt5所有用户界面对象的基类。他为QWidget提供默认构造函数。默认构造函数没有父类。
#     w = Example()
#     #设置窗口的标题
#     w.setWindowTitle('pptx文件log生成器')
#     #显示在屏幕上
#     w.show()
    
#     #系统exit()方法确保应用程序干净的退出
#     #的exec_()方法有下划线。因为执行是一个Python关键词。因此，exec_()代替
#     sys.exit(app.exec_())
import sys
import os

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from Core import pptxGenerate as pptxg
from Core import Log
from Core import TxtToLog as ttl
from PyQt5.QtWidgets import QApplication, QDialog, QFileDialog, QMainWindow, QMessageBox
from PyQt5 import QtWidgets, QtCore
from PyQt5.uic import loadUi

from pptx import Presentation

# 修改角色信息 副窗口
class EditCharDialog(QDialog):
    _signal = QtCore.pyqtSignal(Log.Character, int)
    def __init__(self, parent=None):
        super(EditCharDialog,self).__init__(parent)
        loadUi('./GUI/editCharacter.ui', self)
        # 从父窗口读取指定修改的角色序号
        self.index = getattr(parent, 'row_to_edit', 0)
        # 获取原角色参数值
        self.char = getattr(parent, 'char_list', [Log.KP('','')])[self.index]
        # 设置取磁盘位置按钮功能
        self.char_imgdir_btn.clicked.connect(self.browsefiles_char_imgdir)
        # 展示原角色参数值
        #   - 下拉菜单默认值
        self.char_identity_select.setCurrentIndex(self.identity_to_index())
        self.disable_inputs(self.char.get_identity())
        #   - 其他文本输入默认值
        self.char_name_input.setText(self.char.name)
        self.char_qnum_input.setText(getattr(self.char,'qNum','-'))
        self.char_PLid_input.setText(getattr(self.char,'PLid','-'))
        self.char_imgdir_input.setText(getattr(self.char,'imgdir','-'))
        # 下拉菜单无效填充
        self.char_identity_select.activated[str].connect(self.disable_inputs)
        # 保存按钮 功能
        self.save.clicked.connect(self.save_and_close)
        # 取消按钮 功能
        self.cancel.clicked.connect(self.cancel_and_close)
    
    # 取消 并关闭子窗口
    def cancel_and_close(self):
        self._signal.emit(self.char, self.index)
        self.close()

    # 保存 并关闭子窗口
    def save_and_close(self):
        self._signal.emit(self.get_new_char(), self.index)
        self.close()

    # 获取当前窗口填写的 新参数值 
    def get_new_char(self):
        if self.char_identity_select.currentText() == 'KP':
            name = self.char_PLid_input.text()
            qnum = str(self.char_qnum_input.text())
            char = Log.KP(name, qnum)
            char = self.get_imgpath(char)
            return char
        elif self.char_identity_select.currentText() == 'PL':
            name = self.char_name_input.text()
            qnum = str(self.char_qnum_input.text())
            plid = self.char_PLid_input.text()
            char = Log.PL(name, qnum, plid)
            char = self.get_imgpath(char)
            return char
        elif self.char_identity_select.currentText() == '骰子':
            name = self.char_name_input.text()
            char = Log.Dice(name)
            char = self.get_imgpath(char)
            return char
        elif self.char_identity_select.currentText() == 'NPC':
            name = self.char_name_input.text()
            char = Log.Character(name)
            char = self.get_imgpath(char)
            return char
        else:
            msgBox = QMessageBox(QMessageBox.NoIcon, '未实装', '当前功能未实装，敬请期待。')
            msgBox.exec()
            return None

    # 获取当前角色身份并转为下拉菜单列表的序号
    def identity_to_index(self):
        print(self.char)
        if self.char.get_identity() == 'KP':
            return 0
        elif self.char.get_identity() == 'PL':
            return 1
        elif self.char.get_identity() == '骰子':
            return 2
        elif self.char.get_identity() == 'NPC':
            return 3
        else:
            msgBox = QMessageBox(QMessageBox.NoIcon, '错误', '无法判断角色身份。')
            msgBox.exec()
            self.close()

    # 打开获取文件列表的窗口
    def browsefiles_char_imgdir(self):
        fname = QFileDialog.getExistingDirectory(None, "选择角色立绘文件夹", "")
        self.char_imgdir_input.setText(fname)

    # 控制填充值无效 并 删除无效输入值
    def disable_inputs(self, identity):
        if identity == 'KP':
            self.char_name_input.setEnabled(False)
            self.char_name_label.setEnabled(False)
            self.char_name_input.setText('KP')
            self.char_PLid_input.setEnabled(True)
            self.char_PLid_label.setEnabled(True)
            self.char_PLid_input.setText('')
            self.char_qnum_input.setEnabled(True)
            self.char_qnum_label.setEnabled(True)
            self.char_qnum_input.setText('')
        elif identity == '骰子' or identity == 'NPC':
            self.char_name_input.setEnabled(True)
            self.char_name_label.setEnabled(True)
            self.char_name_input.setText('')
            self.char_PLid_input.setEnabled(False)
            self.char_PLid_label.setEnabled(False)
            self.char_PLid_input.setText('')
            self.char_qnum_input.setEnabled(False)
            self.char_qnum_label.setEnabled(False)
            self.char_qnum_input.setText('')
        else:
            self.char_name_input.setEnabled(True)
            self.char_name_label.setEnabled(True)
            self.char_name_input.setText('')
            self.char_PLid_input.setEnabled(True)
            self.char_PLid_label.setEnabled(True)
            self.char_PLid_input.setText('')
            self.char_qnum_input.setEnabled(True)
            self.char_qnum_label.setEnabled(True)
            self.char_qnum_input.setText('')

    # 尝试获取imgpath，判断地址是否有效
    def get_imgpath(self, char):
        if self.char_imgdir_input.isEnabled():
                try:
                    char.set_imgpath(self.char_imgdir_input.text())
                except:
                    msgBox = QMessageBox(QMessageBox.NoIcon, '警告', '错误的立绘文件夹地址')
                    msgBox.exec()
        return char

# 主窗口
class MainWindow(QMainWindow):
    def __init__(self):
        super(MainWindow,self).__init__()
        loadUi('./GUI/replay.ui', self)
        self.char_list = []
        # 根据当前选择无效填充位置
        self.disable_inputs(self.char_identity_select.currentText())
        
        # 全部fp读取按钮功能安装
        self.char_fp_btn.clicked.connect(self.browsefiles_char_fp)
        self.char_imgdir_btn.clicked.connect(self.browsefiles_char_imgdir)
        self.model_fp_btn.clicked.connect(self.browsefiles_model_fp)
        self.txt_fp_btn.clicked.connect(self.browsefiles_txt_fp)
        self.save_fp_btn.clicked.connect(self.browsefiles_save_fp)
        self.add_Character_btn.clicked.connect(self.addlist_char)
        #------------------------------#
        # 角色部分：
        #   - 角色身份选择按钮效果
        self.char_identity_select.activated[str].connect(self.disable_inputs)
        #   - 读取角色列表按钮效果
        self.char_read_list_btn.clicked.connect(self.read_char_list)
        
        #   - 角色展示表内容
        #   -   -四列宽度设置
        self.char_table.horizontalHeader().resizeSection(0,80)
        self.char_table.horizontalHeader().resizeSection(1,159)
        self.char_table.horizontalHeader().resizeSection(2,110)
        self.char_table.horizontalHeader().resizeSection(3,140)
        #   -   -设定选定方式为整行
        self.char_table.setSelectionBehavior(QtWidgets.QTableView.SelectRows)
        #   -   -行表头不显示
        self.char_table.verticalHeader().setVisible(False)
        #   -   -开启右键副菜单
        self.char_table.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
        #   -   -右键副菜单功能安装
        self.char_table.customContextMenuRequested.connect(self.char_list_rightclick)
        #   -   -当前char_list填入表格显示
        self.fill_char_list()

        #   - 修改文字字体部分
        #   -   -确定文字修改勾选框功能
        self.font_change.stateChanged.connect(self.set_font_change)
        self.set_font_change(self.font_change.checkState())
        #   -   -字号范围，单步大小，默认值
        self.font_size_select.setRange(1, 99)
        self.font_size_select.setSingleStep(1)
        self.font_size_select.setValue(18)

        # - 开始生成 按钮 功能安装
        self.start_generation.clicked.connect(self.generate)

    # 根据勾选框判断是否无效填充字体部分
    def set_font_change(self, state):
        if state == QtCore.Qt.Checked:
            self.font_name_label.setEnabled(True)
            self.font_name_select.setEnabled(True)
            self.font_size_label.setEnabled(True)
            self.font_size_select.setEnabled(True)
            self.font_bold.setEnabled(True)
        else:
            self.font_name_label.setEnabled(False)
            self.font_name_select.setEnabled(False)
            self.font_size_label.setEnabled(False)
            self.font_size_select.setEnabled(False)
            self.font_bold.setEnabled(False)

    # 读取角色列表json文件
    def read_char_list(self):
        try:
            self.char_list = ttl.getCharacterList(self.char_fp_input.text())
        except:
            msgBox = QMessageBox(QMessageBox.NoIcon, '错误', '读取失败，请检查角色列表路径格式是否正确。')
            msgBox.exec()
        self.fill_char_list()
        
    # 角色表格 右键副菜单 实现
    def char_list_rightclick(self, pos):
        # 获取当前点击位置的行列值。 index.row() → 第几行  index.column() → 第几列
        index = self.char_table.indexAt(pos)
        if index.isValid():
            # 创建右键菜单可选功能列表
            menu = QtWidgets.QMenu()
            delete_row_action = menu.addAction('删除角色')
            edit_row_action = menu.addAction('修改角色')
            # 通过mapToGlobal判断当前鼠标位置的相对位置， 以确认之后点击选择的是哪一个功能
            action = menu.exec_(self.char_table.viewport().mapToGlobal(pos))
            
            # 选择 删除角色 时对应的操作
            if action == delete_row_action:
                # 弹出确认窗口
                reply = QMessageBox.question(self,'删除确认','确认删除该角色？', QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
                # 判断 已经点击确认后 进行删除操作
                if reply == QMessageBox.Yes:
                    row_num = index.row()
                    self.char_table.removeRow(row_num)
                    self.char_list.pop(row_num)
                    self.save_char_list()
            # 选择 编辑角色 时对应的操作
            if action == edit_row_action:
                self.row_to_edit= index.row()
                # 开启副窗口
                self.edit_char_dialog = EditCharDialog(self)
                self.edit_char_dialog.show()
                # 通过副窗口传递回的值修改当前角色
                self.edit_char_dialog._signal.connect(self.edit_char)

    # 根据输入的 行数 填入 角色新信息      
    def edit_char(self, char, row):
        print('row:', row)
        print('char:', char)
        self.char_list[row] = char
        self.save_char_list()
        self.fill_char_list()

    # 刷新填入当前角色列表
    def fill_char_list(self):
        self.char_table.clearContents()
        self.char_table.setRowCount(0)
        for rowPosition, c in enumerate(self.char_list):
            self.char_table.insertRow(rowPosition)
            self.char_table.setItem(rowPosition, 0, QtWidgets.QTableWidgetItem(c.get_identity()))
            self.char_table.setItem(rowPosition, 1, QtWidgets.QTableWidgetItem(c.name))
            self.char_table.setItem(rowPosition, 2, QtWidgets.QTableWidgetItem(getattr(c, 'PLid', '-')))
            self.char_table.setItem(rowPosition, 3, QtWidgets.QTableWidgetItem(getattr(c, 'qNum', '-')))

    # 无效化输入
    def disable_inputs(self, identity):
        if identity == 'KP':
            self.char_name_input.setEnabled(False)
            self.char_name_label.setEnabled(False)
            self.char_name_input.setText('KP')
            self.char_PLid_input.setEnabled(True)
            self.char_PLid_label.setEnabled(True)
            self.char_PLid_input.setText('')
            self.char_qnum_input.setEnabled(True)
            self.char_qnum_label.setEnabled(True)
            self.char_qnum_input.setText('')
        elif identity == '骰子' or identity == 'NPC':
            self.char_name_input.setEnabled(True)
            self.char_name_label.setEnabled(True)
            self.char_name_input.setText('')
            self.char_PLid_input.setEnabled(False)
            self.char_PLid_label.setEnabled(False)
            self.char_PLid_input.setText('')
            self.char_qnum_input.setEnabled(False)
            self.char_qnum_label.setEnabled(False)
            self.char_qnum_input.setText('')
        else:
            self.char_name_input.setEnabled(True)
            self.char_name_label.setEnabled(True)
            self.char_name_input.setText('')
            self.char_PLid_input.setEnabled(True)
            self.char_PLid_label.setEnabled(True)
            self.char_PLid_input.setText('')
            self.char_qnum_input.setEnabled(True)
            self.char_qnum_label.setEnabled(True)
            self.char_qnum_input.setText('')

    # 创建新角色并存入角色列表
    def addlist_char(self):
        if self.char_identity_select.currentText() == 'KP':
            name = self.char_PLid_input.text()
            qnum = str(self.char_qnum_input.text())
            # 创建角色
            char = Log.KP(name, qnum)
            # 尝试读取立绘列表
            char = self.get_imgpath(char)
            # 保存新角色
            self.char_list.append(char)
        elif self.char_identity_select.currentText() == 'PL':
            name = self.char_name_input.text()
            qnum = str(self.char_qnum_input.text())
            plid = self.char_PLid_input.text()
            # 创建角色
            char = Log.PL(name, qnum, plid)
            # 尝试读取立绘列表
            char = self.get_imgpath(char)
            # 保存新角色
            self.char_list.append(char)
        elif self.char_identity_select.currentText() == '骰子':
            name = self.char_name_input.text()
            # 创建角色
            char = Log.Dice(name)
            # 尝试读取立绘列表
            char = self.get_imgpath(char)
            # 保存新角色
            self.char_list.append(char)
        elif self.char_identity_select.currentText() == 'NPC':
            name = self.char_name_input.text()
            # 创建角色
            char = Log.Character(name)
            # 尝试读取立绘列表
            char = self.get_imgpath(char)
            # 保存新角色
            self.char_list.append(char)
        else:
            msgBox = QMessageBox(QMessageBox.NoIcon, '未实装', '当前功能未实装，敬请期待。')
            msgBox.exec()
        # 保存当前角色列表 （更新json）
        self.save_char_list()
        # 刷新当前表格内的角色列表 （更新显示）
        self.fill_char_list()
            
    # 尝试保存当前角色列表 含报错
    def save_char_list(self):
        try:
            ttl.saveCharacterList(self.char_list,self.char_fp_input.text())
        except:
            if self.char_list:
                self.char_list.pop(-1)
                msgBox = QMessageBox(QMessageBox.NoIcon, '保存失败','保存玩家列表失败。')
                msgBox.exec()

    # 尝试获取立绘列表 含报错
    def get_imgpath(self, char):
        if self.char_imgdir_input.isEnabled():
                try:
                    char.set_imgpath(self.char_imgdir_input.text())
                except:
                    msgBox = QMessageBox(QMessageBox.NoIcon, '警告', '错误的立绘文件夹地址')
                    msgBox.exec()
        return char

    # 按钮打开文件地址 实现
    def browsefiles_char_fp(self):
        fname = QFileDialog.getOpenFileName(self, "导入角色列表", "", 'JSON 文件 (*.json)')
        self.char_fp_input.setText(fname[0])

    def browsefiles_char_imgdir(self):
        fname = QFileDialog.getExistingDirectory(None, "选择角色立绘文件夹", "")
        self.char_imgdir_input.setText(fname)

    def browsefiles_model_fp(self):
        fname = QFileDialog.getOpenFileName(self, "选择pptx模板文件", "", 'PPTX 文件 (*.pptx)')
        self.model_fp_input.setText(fname[0])

    def browsefiles_txt_fp(self):
        fname = QFileDialog.getOpenFileName(self, "选择模组记录txt文件", "", 'txt 文件 (*.txt)')
        self.txt_fp_input.setText(fname[0])

    def browsefiles_save_fp(self):
        fname = QFileDialog.getExistingDirectory(None, "选择生成pptx文件存储目标位置", "")
        self.save_fp_input.setText(fname)

    # 检查数据完整性
    def check_valid(self):
        if not self.get_txt_format():
            return False
        try:
            f = open(self.txt_fp_input.text())
        except:
            msgBox = QMessageBox(QMessageBox.NoIcon, '错误', '请填写正确的txt文件路径。')
            msgBox.exec()
            return False
        try:
            f = open(self.model_fp_input.text())
        except:
            msgBox = QMessageBox(QMessageBox.NoIcon, '错误', '加载pptx模板失败，请填写正确的pptx模板路径。')
            msgBox.exec()
            return False
        if not self.model_default_replacer_input.text():
            msgBox = QMessageBox(QMessageBox.NoIcon, '错误', '填充内容识别不能为空。')
            msgBox.exec()
            return False
        try:
            f = open(self.char_fp_input.text())
        except:
            msgBox = QMessageBox(QMessageBox.NoIcon, '错误', '加载角色列表失败，请填写正确的角色列表路径或创建新的角色列表。')
            msgBox.exec()
            return False
        if not os.path.exists(self.save_fp_input.text()):
            msgBox = QMessageBox(QMessageBox.NoIcon, '错误', '不存在的文件保存地址。请确定输入正确。')
            msgBox.exec()
            return False
        fp = os.path.join(self.save_fp_input.text(), self.save_filename_input.text() + '.pptx') 
        if os.path.exists(fp):
            reply = QMessageBox.question(self, '文件已存在', '文件已存在，是否确定覆盖？', QMessageBox.Yes | QMessageBox.No , QMessageBox.Yes)
            if reply == QMessageBox.No:
                return False
        return True


    def get_txt_format(self):
        if self.txt_format_input.currentText() == '朗读女适配':
            return 'ldn'
        if self.txt_format_input.currentText() == 'QQ聊天记录导出':
            return 'qq'
        if self.txt_format_input.currentText() == '跑团记录染色器':
            return 'colored'
        
        msgBox = QMessageBox(QMessageBox.NoIcon, '错误', '格式选择功能出错，请联系程序作者。')
        msgBox.exec()
        return None


    # 开始生成
    def generate(self):
        if not self.check_valid():
            return None
        txt_fp = self.txt_fp_input.text()
        txt_format = self.get_txt_format()
        txt_encoding = self.txt_encoding_input.currentText()
        model_ppt_fp = self.model_fp_input.text()
        char_fp = self.char_fp_input.text()
        name = self.save_filename_input.text()
        save_fp = self.save_fp_input.text()
        font_change = self.font_change.isChecked()
        font_size = self.font_size_select.value()
        font_bold = self.font_bold.isChecked()
        font_name = self.font_name_select.currentText()
        d_replacer = self.model_default_replacer_input.text()

        try:
            lines = ttl.readFromTxt(txt_fp, self.char_list, char_fp, txt_format=txt_format, encoding=txt_encoding)
        except:
            msgBox = QMessageBox(QMessageBox.NoIcon, '警告', '读取txt文件失败，请检查格式是否正确，或更换编码。')
            msgBox.exec()
            return None

        scene = Log.Scene(lines, characters=self.char_list)
        prs = Presentation(model_ppt_fp)
        dictionary = pptxg.getSlideDictionary(prs.slides, [c.name for c in self.char_list], {})
        try:
            pptxg.generatePreFromLog(scene,dictionary,prs,name, save_fp, d_replacer, font_change=font_change, font_size=font_size, font_bold=font_bold, font_name=font_name)
        except:
            msgBox = QMessageBox(QMessageBox.NoIcon, '错误', '生成失败。')
            msgBox.exec()
            return None
        else:
            msgBox = QMessageBox(QMessageBox.NoIcon, '成功', '生成完毕。')
            msgBox.exec()
            return None




app = QApplication(sys.argv)
mainwindow = MainWindow()
widget = QtWidgets.QStackedWidget()
widget.addWidget(mainwindow)
widget.setFixedSize(510,930)
widget.show()
sys.exit(app.exec_())


