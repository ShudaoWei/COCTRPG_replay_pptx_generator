
import sys
import os
from Core import pptxGenerate as pptxg
from Core import Log
from Core import TxtToLog as ttl
# from GUI.UI_ECWindow import Ui_EditCharacter as ec_Dialog
from GUI.UI_MainWindow import Ui_MainWindow as rep_MW
from PyQt5.QtWidgets import QApplication, QDialog, QFileDialog, QMainWindow, QMessageBox
from PyQt5 import QtWidgets, QtCore
from PyQt5.uic import loadUi

from pptx import Presentation

# ---ver2.0 废弃弹出窗口---
# 修改角色信息 副窗口
#class EditCharDialog(QDialog, ec_Dialog):
#  class EditCharDialog(QtWidgets.QWidget):
#     _signal = QtCore.pyqtSignal(Log.Character, int)
#     def __init__(self, parent=None):
#         super(EditCharDialog,self).__init__(parent)
#         loadUi('./GUI/createCharacter.ui', self)
#         # self.setupUi(self)
#         # 从父窗口读取指定修改的角色序号
#         self.index = getattr(parent, 'row_to_edit', 0)
#         # 获取原角色参数值
#         self.char = getattr(parent, 'char_list', [Log.KP('','')])[self.index]
#         # 设置取磁盘位置按钮功能
#         self.char_imgdir_btn.clicked.connect(self.browsefiles_char_imgdir)
#         # 展示原角色参数值
#         #   - 下拉菜单默认值
#         self.char_identity_select.setCurrentIndex(self.identity_to_index())
#         self.disable_inputs(self.char.get_identity())
#         #   - 其他文本输入默认值
#         self.char_name_input.setText(self.char.name)
#         self.char_qnum_input.setText(getattr(self.char,'qNum','-'))
#         self.char_PLid_input.setText(getattr(self.char,'PLid','-'))
#         self.char_imgdir_input.setText(getattr(self.char,'imgdir','-'))
#         # 下拉菜单无效填充
#         self.char_identity_select.activated[str].connect(self.disable_inputs)
#         # 保存按钮 功能
#         self.save.clicked.connect(self.save_and_close)
#         # 取消按钮 功能
#         self.cancel.clicked.connect(self.cancel_and_close)
    
#     # 取消 并关闭子窗口
#     def cancel_and_close(self):
#         self._signal.emit(self.char, self.index)
#         self.close()

#     # 保存 并关闭子窗口
#     def save_and_close(self):
#         self._signal.emit(self.get_new_char(), self.index)
#         self.close()

#     # 获取当前窗口填写的 新参数值 
#     def get_new_char(self):
#         if self.char_identity_select.currentText() == 'KP':
#             name = self.char_PLid_input.text()
#             qnum = str(self.char_qnum_input.text())
#             char = Log.KP(name, qnum)
#             char = self.get_imgpath(char)
#             return char
#         elif self.char_identity_select.currentText() == 'PL':
#             name = self.char_name_input.text()
#             qnum = str(self.char_qnum_input.text())
#             plid = self.char_PLid_input.text()
#             char = Log.PL(name, qnum, plid)
#             char = self.get_imgpath(char)
#             return char
#         elif self.char_identity_select.currentText() == '骰子':
#             name = self.char_name_input.text()
#             char = Log.Dice(name)
#             char = self.get_imgpath(char)
#             return char
#         elif self.char_identity_select.currentText() == 'NPC':
#             name = self.char_name_input.text()
#             char = Log.Character(name)
#             char = self.get_imgpath(char)
#             return char
#         else:
#             msgBox = QMessageBox(QMessageBox.NoIcon, '未实装', '当前功能未实装，敬请期待。')
#             msgBox.exec()
#             return None

#     # 获取当前角色身份并转为下拉菜单列表的序号
#     def identity_to_index(self):
#         print(self.char)
#         if self.char.get_identity() == 'KP':
#             return 0
#         elif self.char.get_identity() == 'PL':
#             return 1
#         elif self.char.get_identity() == '骰子':
#             return 2
#         elif self.char.get_identity() == 'NPC':
#             return 3
#         else:
#             msgBox = QMessageBox(QMessageBox.NoIcon, '错误', '无法判断角色身份。')
#             msgBox.exec()
#             self.close()

#     # 打开获取文件列表的窗口
#     def browsefiles_char_imgdir(self):
#         fname = QFileDialog.getExistingDirectory(None, "选择角色立绘文件夹", "")
#         self.char_imgdir_input.setText(fname)

#     # 控制填充值无效 并 删除无效输入值
#     def disable_inputs(self, identity):
#         if identity == 'KP':
#             self.char_name_input.setEnabled(False)
#             self.char_name_label.setEnabled(False)
#             self.char_name_input.setText('KP')
#             self.char_PLid_input.setEnabled(True)
#             self.char_PLid_label.setEnabled(True)
#             self.char_PLid_input.setText('')
#             self.char_qnum_input.setEnabled(True)
#             self.char_qnum_label.setEnabled(True)
#             self.char_qnum_input.setText('')
#         elif identity == '骰子' or identity == 'NPC':
#             self.char_name_input.setEnabled(True)
#             self.char_name_label.setEnabled(True)
#             self.char_name_input.setText('')
#             self.char_PLid_input.setEnabled(False)
#             self.char_PLid_label.setEnabled(False)
#             self.char_PLid_input.setText('')
#             self.char_qnum_input.setEnabled(False)
#             self.char_qnum_label.setEnabled(False)
#             self.char_qnum_input.setText('')
#         else:
#             self.char_name_input.setEnabled(True)
#             self.char_name_label.setEnabled(True)
#             self.char_name_input.setText('')
#             self.char_PLid_input.setEnabled(True)
#             self.char_PLid_label.setEnabled(True)
#             self.char_PLid_input.setText('')
#             self.char_qnum_input.setEnabled(True)
#             self.char_qnum_label.setEnabled(True)
#             self.char_qnum_input.setText('')

#     # 尝试获取imgpath，判断地址是否有效
#     def get_imgpath(self, char):
#         if self.char_imgdir_input.isEnabled():
#                 try:
#                     char.set_imgpath(self.char_imgdir_input.text())
#                 except:
#                     msgBox = QMessageBox(QMessageBox.NoIcon, '警告', '错误的立绘文件夹地址')
#                     msgBox.exec()
#         return char

# 主窗口
class MainWindow(QMainWindow, rep_MW):
#class MainWindow(QMainWindow):
    def __init__(self, parent=None):
        super(MainWindow,self).__init__(parent)
        # loadUi('./GUI/replay_ver2.ui', self)
        self.setupUi(self)
        self.setFixedSize(510, 625)
        self.char_list = []
        
        # 全部fp读取按钮功能安装
        self.char_fp_btn.clicked.connect(self.browsefiles_char_fp)
        self.model_fp_btn.clicked.connect(self.browsefiles_model_fp)
        self.txt_fp_btn.clicked.connect(self.browsefiles_txt_fp)
        self.save_fp_btn.clicked.connect(self.browsefiles_save_fp)

        # 添加角色按钮实装
        self.add_Character_btn.clicked.connect(lambda state, aim='add': self.show_page_char(aim))
        #------------------------------#
        # 角色部分：
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

        # #------=-读取字体已修复，丢弃-=---------
        # #   - 修改文字字体部分 
        # #   -   -确定文字修改勾选框功能
        # self.font_change.stateChanged.connect(self.set_font_change)
        # self.set_font_change(self.font_change.checkState())
        # #   -   -字号范围，单步大小，默认值
        # self.font_size_select.setRange(1, 99)
        # self.font_size_select.setSingleStep(1)
        # self.font_size_select.setValue(18)


        # 初始化图像读取路径
        self.char_imgdir_btn.clicked.connect(self.browsefiles_char_imgdir)
        # 实装 角色身份选择 按钮
        self.char_identity_select.activated[str].connect(self.disable_inputs)
        # 实装 读取立绘差分列表 按钮
        self.pic_get_fps_btn.clicked.connect(self.read_imgdict_list)
        # 实装 启动立绘自动替换 按钮
        self.pic_enable.stateChanged.connect(self.disable_pic_autochange)
        # 在表格内容（关键词修改）时更新关键词个数
        self.pic_keyword_table.cellChanged.connect(self.refresh_keywordCount)
        # 在模板默认立绘下拉菜单修改时更新表格
        self.pic_default_input.activated.connect(self.change_default)
        # page_char页面取消按钮 实装
        self.char_cancel_btn.clicked.connect(self.close_page)

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
                # # 开启副窗口
                # self.edit_char_dialog = EditCharDialog(self)
                # self.edit_char_dialog.show()
                # 通过副窗口传递回的值修改当前角色
                # self.edit_char_dialog._signal.connect(self.edit_char)
                self.show_page_char('edit')

    # 根据输入的 行数 填入 角色新信息      
    def edit_char(self):
        char = self.get_char()
        self.char_list[self.row_to_edit] = char
        self.save_char_list()
        self.fill_char_list()
        self.close_page()
        self.char_save_btn.clicked.disconnect(self.edit_char)

    # 更新表格的默认立绘行
    def change_default(self, new_index):
        rows = self.pic_keyword_table.rowCount()
        for row in range(rows):
            if row == new_index:
                cell = QtWidgets.QTableWidgetItem('')
                cell.setFlags(QtCore.Qt.ItemIsEnabled)
                cell.setFlags(QtCore.Qt.ItemIsSelectable)
                self.pic_keyword_table.setItem(row, 2, cell)
                cell0 = self.pic_keyword_table.item(row,0)
                cell1 = self.pic_keyword_table.item(row,1)
                cell0.setFlags(QtCore.Qt.ItemIsEnabled)
                cell0.setFlags(QtCore.Qt.ItemIsSelectable)
                cell1.setFlags(QtCore.Qt.ItemIsEnabled)
                cell1.setFlags(QtCore.Qt.ItemIsSelectable)
            else:
                item = self.pic_keyword_table.item(row, 0)
                item.setFlags(QtCore.Qt.ItemIsEnabled)
                item = self.pic_keyword_table.item(row, 1)
                item.setFlags(QtCore.Qt.ItemIsEnabled)
                item = self.pic_keyword_table.item(row, 2)
                item.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEditable)

    # 刷新填入当前角色列表
    def fill_char_list(self):
        self.char_table.clearContents()
        self.char_table.setRowCount(0)
        for rowPosition, c in enumerate(self.char_list):
            self.char_table.insertRow(rowPosition)
            self.char_table.setItem(rowPosition, 0, QtWidgets.QTableWidgetItem(c.get_identity()))
            self.char_table.setItem(rowPosition, 1, QtWidgets.QTableWidgetItem(c.name))
            self.char_table.setItem(rowPosition, 2, QtWidgets.QTableWidgetItem(getattr(c, 'PLid', '-')))
            self.char_table.setItem(rowPosition, 3, QtWidgets.QTableWidgetItem(str(len(c.imges))))

    # 刷新填入关键词dictionary 列表
    def fill_keyword_list(self, dictionary):
        self.pic_keyword_table.blockSignals(True)
        self.pic_keyword_table.clearContents()
        self.pic_keyword_table.setRowCount(0)
        for row, key in enumerate(dictionary):
            self.pic_keyword_table.insertRow(row)
            filename = os.path.basename(key)
            cell0 = QtWidgets.QTableWidgetItem(filename)
            cell0.setFlags(QtCore.Qt.ItemIsSelectable)
            self.pic_keyword_table.setItem(row, 0, cell0)
            words = dictionary[key]
            count = str(len(words))
            cell1 = QtWidgets.QTableWidgetItem(count)
            cell1.setFlags(QtCore.Qt.ItemIsSelectable)
            self.pic_keyword_table.setItem(row, 1, cell1)
            cell2 = QtWidgets.QTableWidgetItem(';'.join(words))
            if row == self.pic_default_input.currentIndex():
                cell2 = QtWidgets.QTableWidgetItem('')
                cell2.setFlags(QtCore.Qt.ItemIsEnabled)
                cell2.setFlags(QtCore.Qt.ItemIsSelectable)
            else:
                cell0.setFlags(QtCore.Qt.ItemIsEnabled)
                cell1.setFlags(QtCore.Qt.ItemIsEnabled)
            self.pic_keyword_table.setItem(row, 2, cell2)
        self.pic_keyword_table.blockSignals(False)

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

    # 获取当前角色身份并转为下拉菜单列表的序号
    def identity_to_index(self, char):
        identity = char.get_identity()
        if identity == 'KP':
            return 0
        elif identity == 'PL':
            return 1
        elif identity == '骰子':
            return 2
        elif identity == 'NPC':
            return 3
        else:
            msgBox = QMessageBox(QMessageBox.NoIcon, '错误', '无法判断角色身份。')
            msgBox.exec()
            self.close()

    # 信号重连函数(https://www.coder.work/article/339536)
    def reconnect(self, signal, newhandler=None, oldhandler=None):        
        try:
            if oldhandler is not None:
                while True:
                    signal.disconnect(oldhandler)
            else:
                signal.disconnect()
        except TypeError:
            pass
        if newhandler is not None:
            print(signal)
            signal.connect(newhandler)

    # 角色修改界面开启函数
    def show_page_char(self, aim):
        # 清空界面
        self.clear_page_char()
        # 显示右侧page_char
        self.setFixedSize(1039, 625)
        self.stackedWidget.setCurrentWidget(self.page_char)
        # 新建角色界面
        if aim == 'add':
            self.char_createNewChar.setTitle('&新建玩家')
            self.char_save_btn.setText('添加角色')
            # 保存按钮
            self.reconnect(self.char_save_btn.clicked, self.add_list)
            # 根据当前选择无效填充位置
            self.disable_inputs(self.char_identity_select.currentText())

        # 修改角色界面
        else:
            self.char_createNewChar.setTitle('&玩家设置')
            self.char_save_btn.setText('保存')
            # 获取原角色参数值
            index = getattr(self, 'row_to_edit', 0)
            char = getattr(self, 'char_list', [Log.KP('','')])[index]
            print(char.imges_dict)
            print(char.original_img)
            #   - 下拉菜单默认值
            self.char_identity_select.setCurrentIndex(self.identity_to_index(char))
            # 根据当前选择无效填充位置
            self.disable_inputs(self.char_identity_select.currentText())
            #   - 其他文本输入默认值
            self.char_name_input.setText(char.name)
            self.char_qnum_input.setText(getattr(char,'qNum','-'))
            self.char_PLid_input.setText(getattr(char,'PLid','-'))
            self.char_imgdir_input.setText(getattr(char,'imgdir',''))
            if self.char_imgdir_input.text():
                self.pic_enable.setChecked(True)
                self.pic_default_input.clear()
                self.pic_default_input.addItems(char.get_filenames())
                self.pic_default_input.setCurrentIndex(char.original_img)
                self.change_default(char.original_img)
                self.fill_keyword_list(char.imges_dict)
            else:
                self.pic_enable.setChecked(False)
            # 保存按钮
            self.reconnect(self.char_save_btn.clicked, self.edit_char)

    # 刷新立绘关键词列表的关键词计数
    def refresh_keywordCount(self, row, col):
        if col == 2:
            self.pic_keyword_table.blockSignals(True)
            strs = self.pic_keyword_table.item(row, col).text()
            if strs:
                words = strs.split(';')
            else:
                words = []
            count = str(len(words))
            item = self.pic_keyword_table.item(row, 1)
            item.setText(count)
            self.pic_keyword_table.blockSignals(False)

    # 读取立绘路径并获取立绘文件名以及填入dictionary list
    def read_imgdict_list(self):
        temp = Log.Character("temp")            
        self.get_imgpath(temp)
        if temp.imges:
            self.pic_default_input.clear()
            self.pic_default_input.addItems(temp.get_filenames())
            self.fill_keyword_list(temp.imges_dict)

    # 清空 page_char 玩家设置页面
    def clear_page_char(self):
        # 下拉菜单默认值
        self.char_identity_select.setCurrentIndex(0)
        self.disable_inputs(self.char_identity_select.currentText())
        self.clear_pic_autochange()
        self.pic_enable.setChecked(False)

    # 关闭 page 右侧窗口
    def close_page(self):
        self.setFixedSize(510, 625)

    # 清空立绘自动替换模块
    def clear_pic_autochange(self):
        self.pic_autochange.setEnabled(False)
        self.char_imgdir_input.setText('')
        self.pic_default_input.clear()
        self.pic_keyword_table.clearContents()
        self.pic_keyword_table.setRowCount(0)

    # 自动替换模块按钮 功能
    def disable_pic_autochange(self, checkstate):
        if checkstate > 0:
            self.pic_autochange.setEnabled(True)
        else:
            self.clear_pic_autochange()

    # 获取当前窗口内角色设置
    def get_char(self):
        if self.char_identity_select.currentText() == 'KP':
            plid = self.char_PLid_input.text()
            qnum = str(self.char_qnum_input.text())
            # 创建角色
            char = Log.KP(plid, qnum)
        elif self.char_identity_select.currentText() == 'PL':
            name = self.char_name_input.text()
            qnum = str(self.char_qnum_input.text())
            plid = self.char_PLid_input.text()
            # 创建角色
            char = Log.PL(name, qnum, plid)
        elif self.char_identity_select.currentText() == '骰子':
            name = self.char_name_input.text()
            # 创建角色
            char = Log.Dice(name)
        elif self.char_identity_select.currentText() == 'NPC':
            name = self.char_name_input.text()
            # 创建角色
            char = Log.Character(name)
        else:
            msgBox = QMessageBox(QMessageBox.NoIcon, '未实装', '当前功能未实装，敬请期待。')
            msgBox.exec()
            return None
        if self.pic_enable.checkState():
            imgdir = self.char_imgdir_input.text()
            if not imgdir.replace(' ', ''):
                msgBox = QMessageBox(QMessageBox.NoIcon, '错误', '请输入立绘文件夹，或关闭立绘自动替换功能。')
                msgBox.exec()
            if not self.pic_default_input.currentText().replace(' ', ''):
                msgBox = QMessageBox(QMessageBox.NoIcon, '错误', '请读取立绘差分列表，选择模板内默认立绘。')
                msgBox.exec()
            char.original_img = self.pic_default_input.currentIndex()
            char.set_imgpath(imgdir)
            keyword_dict = self.get_keyword_dictionary()
            char.set_dictionary(keyword_dict)
        return char
    
    # 保存添加新角色
    def add_list(self):
        char = self.get_char()
        self.char_list.append(char)
        if self.char_fp_input.text().replace(' ','') == '':
            self.char_fp_input.setText('./characterlist.json')
        # 保存当前角色列表 （更新json）
        self.save_char_list()
        # 刷新当前表格内的角色列表 （更新显示）
        self.fill_char_list()
        self.close_page()
        self.char_save_btn.clicked.disconnect(self.add_list)

    # 从列表获取当前立绘对应关键词列表 的 dicitonary
    def get_keyword_dictionary(self):
        rows = self.pic_keyword_table.rowCount()
        keyword_dict = {}
        for i in range(rows):
            keyword_dict[self.pic_keyword_table.item(i,0).text()] = self.pic_keyword_table.item(i,2).text().split(';')
        return keyword_dict

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
                    self.clear_pic_autochange()
                    self.pic_autochange.setEnabled(True)
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

    # 获取txt格式转换为缩写
    def get_txt_format(self):
        if self.txt_format_input.currentText() == '朗读女适配':
            return 'ldn'
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
        print('新建Presentation成功。')
        try:
            pptxg.generatePreFromLog(scene,model_ppt_fp,name, save_fp, d_replacer, font_change=font_change, font_size=font_size, font_bold=font_bold, font_name=font_name)
        except Exception as e:
            msgBox = QMessageBox(QMessageBox.NoIcon, '错误', '生成失败。\n' + str(e))
            msgBox.exec()
            return None
        else:
            msgBox = QMessageBox(QMessageBox.NoIcon, '成功', '生成完毕。')
            msgBox.exec()
            return None






