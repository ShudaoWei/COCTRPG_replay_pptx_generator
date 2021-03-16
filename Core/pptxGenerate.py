from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
import TxtToLog as ttl
import Log
import copy
import six 

d_replacer = '這一周也不例外，你们刚刚从家里，神社，公园来到集会地点，现在距离集会开始還有一段时間，你们可以自由交流一下'

def checkrecursivelyfortext(shpthissetofshapes,textrun,default_replacer=d_replacer, content='', replace=False):
    # print(pre, '[start shape]\n')
    # pre += '    '
    for shape in shpthissetofshapes:
        # print(pre , shape.shape_type)
        # print(pre,shape)
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            # pre += '    '
            textrun=checkrecursivelyfortext(shape.shapes,textrun,default_replacer, content, replace)
        else:
            if hasattr(shape, "text"):
                # print(pre, shape.text)
                if replace and shape.text.find(default_replacer) !=-1:
                    shape.text = content
                    return []
                textrun.append(shape.text)
    # print(pre, '[end shape]\n')
    return textrun

def copy_slide(pres,pres1,index, content='abcd what?'):
    source = pres.slides[index]
    blank_slide_layout = pres.slide_layouts[0]
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

def generatePreFromLog(scene, dictionary, mods):
    if scene.lines:
        new_pre = Presentation()
        new_pre.slide_width = mods.slide_width
        for l in scene.lines:
            name = l.character.name
            index = dictionary[name]
            checkrecursivelyfortext(mods.slides[index].shapes, [], content=l.content, replace=True)
            #print(checkrecursivelyfortext(mods.slides[index].shapes,[]))
            copy_slide(mods, new_pre, index)
            checkrecursivelyfortext(mods.slides[index].shapes, [],default_replacer=l.content, content=d_replacer, replace=True)
        new_pre.save('test-01.pptx')
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

 
kp = Log.KP('KP','')
zhima = Log.PL('芝麻','', '阿恬')
qiuqiu = Log.PL('球球','', '蓝蓝')
momo = Log.PL('momo','','Syin')
touzi = Log.Dice('骰子')

ttl.saveCharacterList([kp, zhima, qiuqiu, momo, touzi])
charlist = ttl.getCharacterList()
filepath = 'C:/Users/44418/Desktop/TRPG/猫幽灵所寻之物[整理第四部分][正片完].txt'
lines = ttl.readFromTxt(filepath, charlist, txt_format='ldn')
scene = Log.Scene(lines, characters = charlist)
prs = Presentation('cats.pptx')
dictionary = getSlideDictionary(prs.slides, [c.name for c in charlist], {})
generatePreFromLog(scene,dictionary,prs)


# for slide in prs.slides:
#     print(pre + ' 【start slide】\n')
#     print(checkrecursivelyfortext(slide.shapes,[], pre))
#     print(pre + '【end slide】')
#     print(' ')