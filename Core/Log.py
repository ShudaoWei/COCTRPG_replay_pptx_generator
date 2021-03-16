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
        self.is_dice = False
        self.is_NPC = True
        self.is_KP = False
        self.is_PL = False
        for k, w in kw.items():
            setattr(self, k, w)
        print("创建角色：", name)

    def set_imgpath(self, imgdir):
        files = os.listdir(imgdir)
        self.imges = [os.path.join(imgdir, f) for f in files]
    
    def set_KP(self, qNum):
        self.is_KP = True
        self.is_NPC = False
        self.__class__ = KP
        self.qNum = qNum
        print("玩家 " + self.name + " 为守密人。")

    def set_PL(self, qNum):
        self.is_PL = True
        self.is_NPC = False
        self.__class__ = PL
        self.qNum = qNum
        print("设置 " + self.name + " 为PL。")

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

    


class Dice(Character):
    def __init__(self, name, **kw):
        Character.__init__(self, name)
        self.is_dice = True
        self.is_NPC = False
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
        Player.__init__(self, name, qNum)
        self.is_KP = True
        for k, w in kw.items():
            setattr(self, k, w)
        print("玩家 " + name + " 为守密人。\n")
    def __repr__(self):
        result = Character.__repr__(self) + '  该玩家是KP。'
        result += "  qq号：" + str(self.qNum)
        return result


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
    
    def set_State(self, statenum):
        try:
            self.state = self.character.imges[statenum]
        except ValueError:
            print('image set is empty.')

    def __repr__(self):
        result = self.character.name + self.state
        result += ':\n'
        result += self.content + '\n'
        return result

# lines: [Line] (场景内对话)
# background: String(dir) (场景背景 地址)
# characters: [Character] (场景内人员列表)
class Scene(object):
    def __init__(self, lines=[], **kw):
        self.lines = lines
        self.background = ''
        temp = []
        self.characters = [temp.append(l.character) for l in lines if not l.character in temp] 

        for k, w in kw.items():
            setattr(self, k, w)
