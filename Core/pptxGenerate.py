import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Pt
import TxtToLog as ttl
import Log
import copy
import six 

def checkrecursivelyfortext(shpthissetofshapes,textrun,
                            replace=False,default_replacer='默认填充文字', content='', 
                            font_change=False,font_name='',font_bold=False,font_size=19):
    # print(pre, '[start shape]\n')
    # pre += '    '
    for shape in shpthissetofshapes:
        # print(pre , shape.shape_type)
        # print(pre,shape)
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            # pre += '    '
            textrun=checkrecursivelyfortext(shape.shapes,textrun,
                                            replace, default_replacer, content,
                                            font_change, font_name, font_bold, font_size)
        else:
            if hasattr(shape, "text"):
                # print(pre, shape.text)
                if replace and shape.text.find(default_replacer) !=-1:
                    shape.text = content
                    if font_change:
                        for p in shape.text_frame.paragraphs:
                            p.font.name = font_name
                            p.font.bold = font_bold
                            p.font.size = Pt(font_size)
                    return []
                textrun.append(shape.text)
    # print(pre, '[end shape]\n')
    return textrun

def copy_slide(pres,pres1,index):
    source = pres.slides[index]
    blank_slide_layout = pres.slide_layouts[6]
    dest = pres1.slides.add_slide(blank_slide_layout)

    for shp in source.shapes:
        el = shp.element
        newel = copy.deepcopy(el)
        dest.shapes._spTree.insert_element_before(newel, 'p:extLst')



        for key, value in six.iteritems(source.part.rels):
                    # Make sure we don't copy a notesSlide relation as that won't exist
                if not "notesSlide" in value.reltype:
                        dest.part.rels.add_relationship(value.reltype, value._target, value.rId)

        return dest

def findModelSlide(slides, name):
    for i, slide in enumerate(slides):
        if is_in_shapes(slide.shapes, name):
            return i
    return 0

def is_in_shapes(shapes, name):
    result = False
    texts = checkrecursivelyfortext(shapes,[])
    for text in texts:
        text = text.replace(' ','')
        result = result or text.find(name) != -1 
    return result

def getSlideDictionary(slides,names,dictionary):
    for name in names:
        dictionary[name] = findModelSlide(slides, name)
    return dictionary

def generatePreFromLog(scene, dictionary, mods, filename, filepath, default_replacer,**args):
    if scene.lines:
        new_pre = Presentation()
        new_pre.slide_width = mods.slide_width
        for l in scene.lines:
            name = l.character.name
            index = dictionary[name]
            checkrecursivelyfortext(mods.slides[index].shapes, [], default_replacer=default_replacer, content=l.content, replace=True, **args)
            #print(checkrecursivelyfortext(mods.slides[index].shapes,[]))
            copy_slide(mods, new_pre, index)
            checkrecursivelyfortext(mods.slides[index].shapes, [], default_replacer=l.content, content=default_replacer, replace=True)
        fp = os.path.join(filepath, filename + '.pptx')
        new_pre.save(fp)
    else:
        print("台词列表为空。") 

# prs = Presentation('model.pptx')

# pre = ''
# for s in prs.slides:
#     for shape in s.shapes:
#         if hasattr(shape, "text"):
#             a = shape.text
#             shape.text = 'what???'
#             a = 'abcde'
#             print('yes')

# d_replacer = '這一周也不例外，你们刚刚从家里，神社，公园来到集会地点，现在距离集会开始還有一段时間，你们可以自由交流一下'
# kp = Log.KP('KP','')
# zhima = Log.PL('芝麻','', '阿恬')
# qiuqiu = Log.PL('球球','', '蓝蓝')
# momo = Log.PL('momo','','Syin')
# touzi = Log.Dice('骰子')

# path_characters = 'D://TRPG/TRPG_replay_video_generator/characters.json'
# ttl.saveCharacterList([kp, zhima, qiuqiu, momo, touzi], path_characters)
# charlist = ttl.getCharacterList(path_characters)
# txt_fp = 'C:/Users/44418/Desktop/TRPG/猫幽灵所寻之物[后日谈部分编的].txt'
# txt_format = 'ldn'
# txt_encoding = 'UTF-8'
# lines = ttl.readFromTxt(txt_fp, charlist,path_characters, txt_format=txt_format, encoding=txt_encoding)
# scene = Log.Scene(lines, characters=charlist)
# model_ppt_fp = 'cats.pptx'
# prs = Presentation(model_ppt_fp)
# dictionary = getSlideDictionary(prs.slides, [c.name for c in charlist], {})
# name = '后日谈'
# fp = 'C:/Users/44418/Desktop/TRPG'
# font_change = True
# font_size = 20
# font_bold = True
# font_name = '清松手寫體1'
# generatePreFromLog(scene,dictionary,prs,name, fp, d_replacer, font_change=font_change, font_size=font_size, font_bold=font_bold, font_name=font_name)


# for slide in prs.slides:
#     print(pre + ' 【start slide】\n')
#     print(checkrecursivelyfortext(slide.shapes,[], pre))
#     print(pre + '【end slide】')
#     print(' ')