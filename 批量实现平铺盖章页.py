# -*- coding: utf-8 -*-
"""
@Time ： 2022/8/6 11:40
@Auth ： 唐成
@File ：批量实现平铺盖章页.py
@IDE ：PyCharm

"""

import os
from docx import Document
from docx.shared import Inches, Pt, Cm
from add_float_picture import add_float_picture

'''需要修改签章页和说明性文件所在文件夹位置，文档尽量保存为docx'''
path_picture="C:\\Users\\唐成\\Desktop\\图片"
path_word="C:\\Users\\唐成\\Desktop\\word"
pics = os.listdir(path_picture)
words = os.listdir(path_word)
tc=len(pics)

for i in range(tc):

    if __name__ == '__main__':
        document = Document(path_word+"\\"+words[i])

        # add a floating picture
        p = document.paragraphs[-1] ##在最后一段插入

        add_float_picture(p, path_picture+"\\"+pics[i], width=Cm(21), pos_x=Cm(0), pos_y=Cm(0))

        # add text

        document.save(path_word+"\\"+words[i])

