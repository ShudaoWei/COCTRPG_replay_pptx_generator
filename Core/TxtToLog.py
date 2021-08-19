import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
import re
import json
import Log

re_date = r'([0-9]{4}-[0-9]{2}-[0-9]{2})'
re_time = r'([0-9]*:[0-9]*:[0-9]*)'
re_title = r'(【.*】|)'
re_id = r'(.*)'
re_qnum = r'(\([1-9][0-9]*\)|<.*@.*\..*>)'
default_header = re_date + r' *' + re_time + r' *' + re_title + r' *' + re_id + r' *' + re_qnum
re_colored_id = r'<(.*)>'

re_ldn_id = r'\[(.*)\]'

re_dice = r'D[0-9]+=[0-9]+/[0-9]+'


# obj2dic的 code 来源：https://blog.csdn.net/u012410724/article/details/51259761
def object2dict(obj):
    #convert object to a dict
    d = {}
    d['__class__'] = obj.__class__.__name__
    d['__module__'] = obj.__module__
    d.update(obj.__dict__)
    return d

# obj2dic的 code 来源：https://blog.csdn.net/u012410724/article/details/51259761
def dict2object(d):
    #convert dict to object
    if'__class__' in d:
        class_name = d.pop('__class__')
        print('class_name:', class_name)
        module_name = d.pop('__module__')
        print('module_name:', module_name)
        module = __import__(module_name, fromlist=['None'])
        print('module:', module)
        class_ = getattr(module,class_name)
        args = dict((key, value) for key, value in d.items()) #get args
        inst = class_(**args) #create new instance
    else:
        inst = d
    return inst

def getCharacterList(path):
    f = open(path, 'r')
    result = json.loads(f.read(), object_hook = dict2object)
    return result

def saveCharacterList(l, path):
    jslist = json.dumps(l, default=object2dict)
    f = open(path, 'w')
    f.write(jslist)
    f.close()

def readFromTxt(path, charlist, char_fp, encoding='GBK', **kwargs):
    lines = []
    f = open(path, 'r', encoding=encoding)
    l = f.readline() 
    while l: 
        c, cline = findCharacter(l, charlist, char_fp, **kwargs)
        if c:
            if cline.replace(' ',''):
                add_line =True
                clines = cline
                while add_line and l:
                    if not cline:
                        clines += l
                    l = f.readline()
                    c0, cline = findCharacter(l, charlist, char_fp, **kwargs)
                    add_line = not c0
                line = Log.Line(c,clines)
                lines.append(line)
            else:
                l = f.readline()
                add_line = True
                clines = ''
                while add_line and l:
                    clines += l
                    l = f.readline()
                    c0, cline = findCharacter(l, charlist, char_fp, **kwargs)
                    add_line = not c0
                line = Log.Line(c, clines)
                lines.append(line)
        else:
            l = f.readline()
    return lines


# userline 指用户信息行 如： 2021-01-22 4:59:59 洛克 10 60<tseirp@qq.com>
# return内容为 tuple： （说话的角色， 角色发言）
# 如果是风羽染色格式，则tuple 均不为None
# 如果是QQ记录导出格式， 则tuple 第二个值为None
# 如果是朗读女适配格式，则tuple第二个值可能为 '',也被if作为false判断
# txt_format格式为： 
#   qq：qq聊天记录导出
#   colored： 风羽QQ跑团记录染色器 https://logpainter.kokona.tech/
#   ldn： 朗读女格式
def findCharacter(lineString, characters, char_fp, txt_format='qq', header=default_header):
    if txt_format=='qq':
        if header != default_header:
            print('无法识别的ID栏，请使用跑团记录着色器：https://logpainter.kokona.tech/ 更新格式。并使用txt_format="colored"')
            return (None, None)
        is_userline = re.search(header, lineString)
        if not is_userline:
            return (None, None)
        qnum = is_userline.group(5)
        qid = is_userline.group(4)
        for c in characters:
            if c.qNum == qnum or qid.replace(' ','').find(c.PLid) != -1:
                return (c,None)
    elif txt_format == 'colored':
        is_userline = re.search(re_colored_id + r'(.*)', lineString)
        if not is_userline:
            return (None, None)
        name = is_userline.group(1)
        cline = is_userline.group(2)
        for c in characters:
            if (c.is_KP and c.PLid == name) or c.name == name:
                return (c, cline)
        character = Log.Character(name)
        characters.append(character)
        saveCharacterList(characters,char_fp)
        return(character, cline)

    elif txt_format == 'ldn':
        is_userline = re.search(re_ldn_id + r'(.*)', lineString)
        if not is_userline:
            return (None, None)
        name = is_userline.group(1)
        cline = is_userline.group(2)
        for c in characters:
            if (c.is_KP and c.PLid == name) or c.name == name:
                return (c, cline)
        character = Log.Character(name)
        characters.append(character)
        saveCharacterList(characters,char_fp)
        return(character, cline)
    else:
        print("无法识别的txt_format，请从qq, colored 以及 ldn 中选择。其他格式暂时无法识别。")
        return (None, None)
