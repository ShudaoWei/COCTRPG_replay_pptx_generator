import TxtToLog as ttl
import Log
import re

# 内置 默认修改内容
dict_f_default = {'re':True, r'hhh*': '哈哈哈', r'233+': '二三三', r'\(*[xX]+':''}

re_dice = r'.*[0-9]*D[0-9]+=([0-9]+)\/*([0-9]*).*'

def linesToLDN(lines, text_fp, dict_voice, dict_filter):
    f = open(text_fp, 'w')
    for l in lines:
        if l.character.is_dice:
            text = getDiceText(l.content, dict_voice)
        else:
            text = filterCustomize(l.content, dict_f_default)
            if dict_filter:
                text = filterCustomize(text, dict_filter)
            text = '[' + l.character.name + ']' + text
        f.write(text)
    f.close()

def getDiceText(content, dict_voice):
    dVoice = dict_voice.get('dice', '')
    is_normal = re.search(re_dice, content)
    if is_normal and is_normal.group(2):
        dice_result = int(is_normal.group(1))
        dice_skill = int(is_normal.group(2))
        result = dVoice + '\n' + getSpecialVoice(dice_result, dice_skill, dict_voice) + '\n'
    elif is_normal or re.search(r'暗骰', content):
        result = dVoice + '\n'
    else:
        result = ''
        
    return result

def getSpecialVoice(dresult, dskill, dict_voice):
    div = dskill/dresult
    if dresult <= 5:
        return dict_voice.get('successSpecial', getSpecialVoice(6, 30, dict_voice))
    elif div >= 5:
        return dict_voice.get('successExtreme', getSpecialVoice(6, 12, dict_voice))
    elif div >= 2:
        return dict_voice.get('successDifficult', getSpecialVoice(6, 6, dict_voice))
    elif div >= 1:
        return dict_voice.get('success', '')
    elif dresult >= 96:
        return dict_voice.get('failSpecial', getSpecialVoice(6, 1, dict_voice))
    else:
        return dict_voice.get('fail', '')


def filterCustomize(content, dict_filter):
    model = dict_filter.get('re', False)
    if model:
        for key in dict_filter:
            if key == 're':
                continue
            content = re.sub(key, dict_filter[key], content)
    else:
        for key in dict_filter:
            if key == 're':
                continue
            content = content.replace(key, dict_filter[key])
    return content
