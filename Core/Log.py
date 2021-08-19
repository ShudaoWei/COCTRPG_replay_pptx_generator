import os

# name: String (名称)
# imges: [Strings](表情/状态 地址列表)
# is_dice: 是否为骰子
# is_KP: 是否为KP
# is_NPC: 是否为NPC
# is_PL: 是否为玩家
class Character(object):
    def __init__(self, name, **kw):
        self.name = name
        self.imges = []
        self.imgdir = ''
        self.is_dice = False
        self.is_NPC = True
        self.is_KP = False
        self.is_PL = False
        self.imges_dict = {}
        self.original_img = 0
        for k, w in kw.items():
            setattr(self, k, w)
        print("创建角色：", name)

    def set_imgpath(self, imgdir):
        self.imgdir = imgdir
        files = os.listdir(imgdir)
        self.imges = [os.path.join(imgdir, f) for f in files 
                      if os.path.splitext(f)[-1] in ['.jpg', '.jpeg', '.png', '.bmp', '.gif']]
        for img in self.imges:
            self.imges_dict[img] = []
    
    def set_KP(self, qNum):
        self.is_KP = True
        self.is_NPC = False
        self.__class__ = KP
        self.qNum = qNum
        self.PLid = getattr(self, 'PLid', self.name)
        self.name = 'KP'
        print("玩家 " + self.PLid + " 为守密人。")

    def set_PL(self, qNum, PLid):
        self.is_PL = True
        self.is_NPC = False
        self.__class__ = PL
        self.qNum = qNum
        self.PLid = PLid
        print("设置 " + self.PLid + "(角色：" + self.name + ") 为PL。")

    def set_dice(self):
        self.is_dice = True
        self.is_NPC = False
        self.__class__ = Dice
        print("设置 " + self.name + " 为骰子。")
    
    def __repr__(self):
        result = '该角色名称为 :' + self.name + '。'
        if self.imges:
            result += '  有 ' + str(len(self.imges)) + ' 张立绘。'
            result += '\n立绘存储地址：' + str(self.imges) + '\n'
        else:
            result += '  没有立绘。'
        if self.is_NPC:
            result += '  是一名NPC。'
        return result
    
    def get_identity(self):
        if self.is_dice:
            return '骰子'
        if self.is_KP:
            return 'KP'
        if self.is_NPC:
            return 'NPC'
        if self.is_PL:
            return 'PL'
        return None

    def get_filenames(self):
        return [os.path.basename(fp) for fp in self.imges]
    
    def add_dictionary(self, pic_dict):
        for key in pic_dict:
            key_fp = os.path.join(self.imgdir, key)
            # 拼接两个
            new = getattr(self.imges_dict, key_fp, []) + [word for word in pic_dict[key] if word]
            # 去重
            self.imges_dict[key] = list(set(new))

    def set_dictionary(self, pic_dict):
        for key in pic_dict:
            key_fp = os.path.join(self.imgdir, key)
            # 只取重新定义的列表
            new = [word for word in pic_dict[key] if word]
            # 去重
            self.imges_dict[key_fp] = list(set(new))
    
    # def set_original(self, fn):
    #     for index, filepath in self.imges:
    #         if fn == os.path.basename(filepath):
    #             self.original_img = index
    #             return index
    #     return None


    


class Dice(Character):
    def __init__(self, name, **kw):
        Character.__init__(self, name)
        self.is_dice = True
        self.is_NPC = False
        self.results = DiceResult()
        for k, w in kw.items():
            setattr(self, k, w)
        print('这是个骰子。\n')
    def __repr__(self):
        result = Character.__repr__(self) + '  是骰娘。'
        return result

# qNum: String (QQ号)
class Player(Character):
    def __init__(self, name, qNum, **kw):
        Character.__init__(self, name)
        self.qNum = qNum
        self.is_NPC = False
        for k, w in kw.items():
            setattr(self, k, w)
        print("该角色为玩家。\nQQ：", qNum)
    def __repr__(self):
        result = Character.__repr__(self) + '  是玩家。'
        result += "  qq号：" + str(self.qNum)
        return result

# PCname: String (PC角色名)
class PL(Player):
    def __init__(self, name, qNum, PLid, **kw):
        Player.__init__(self, name, qNum)
        self.PLid = PLid
        self.is_PL = True
        for k, w in kw.items():
            setattr(self, k, w)
        print("玩家ID：" + PLid + '\n')
    def __repr__(self):
        result = Character.__repr__(self) + '  该角色是由PL：' + self.PLid + ' 使用。'
        result += "  qq号：" + str(self.qNum)
        return result

class KP(Player):
    def __init__(self, name, qNum, **kw):
        Player.__init__(self, 'KP', qNum)
        self.PLid = name
        self.is_KP = True
        for k, w in kw.items():
            setattr(self, k, w)
        print("玩家 " + name + " 为守密人。\n")
    def __repr__(self):
        result = Character.__repr__(self) + ' 玩家id： ' + self.PLid + '  该玩家是KP。'
        result += "  qq号：" + str(self.qNum)
        return result
    
    def set_PL(self, qNum, name):
        self.is_PL = True
        self.is_NPC = False
        self.__class__ = PL
        self.qNum = qNum
        self.name = name
        print("设置 " + self.PLid + "(角色：" + self.name + ") 为PL。")

# character: Character (发言人)
# content: String (发言内容)
# state: String(directory) (当前状态/表情)
class Line(object):
    def __init__(self, character, content, **kw):
        self.character = character
        self.content = content
        self.state = ''
        for k, w in kw.items():
            setattr(self, k, w)
        line = content.replace(' ','').replace('\n', '')
        self.pic_fp = self._find_char_img(line, character.imges_dict)

    def _find_char_img(self, line, imges_dict):
        for key in imges_dict:
            words = imges_dict[key]
            for word in words:
                if word in line:
                    return key
        return None

    def set_State(self, statenum):
        try:
            self.state = self.character.imges[statenum]
        except ValueError:
            print('立绘列表为空。')

    def __repr__(self):
        result = self.character.name + self.state
        result += ':\n'
        result += self.content + '\n'
        return result

# lines: [Line] (场景内对话)
# background: String(dir) (场景背景 地址)
# characters: [Character] (场景内人员列表)
class Scene(object):
    def __init__(self, lines=[], characters=[], **kw):
        self.lines = lines
        self.background = ''
        temp = []
        self.characters = characters 

        for k, w in kw.items():
            setattr(self, k, w)

    def countKP(self):
        count = 0
        for char in self.characters:
            if char.is_KP:
                count += 1
        return count


class DiceResult(object):
    def __init__(self, **kw):
        self.reg_success = '成功'
        self.hard_success = '困难成功'
        self.extreme_success = '极难成功'
        self.special_success = '大成功'
        self.reg_failure = '失败'
        self.special_failure = '大失败'
        for k, w in kw.items():
            setattr(self, k, w)