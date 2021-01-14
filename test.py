import os

from docx import Document
from docx.shared import Cm, Inches

from paper import Paper

test_docx = 'test1.docx'
document = Document()

# 文档页边距设置
sec = document.sections
# 获取、设置页面边距
sec0 = sec[0]  # 获取章节对象
# 获取页面边距值：（单位为像素）
print('左边距：', sec0.left_margin)
# 左边距： 1143000
print('右边距：', sec0.right_margin)
# 右边距： 1143000
print('上边距：', sec0.top_margin)
# 上边距： 914400
# print('下边距：', sec0.bottom_margin)
# 下边距： 914400

def temp1():
    global sec0
    print()
    sec0.left_margin = Inches(1.2)
    sec0.right_margin = Inches(1.2)
    sec0.top_margin = Inches(1)
    # sec0.bottom_margin = Inches(1)
    print('左边距：', sec0.left_margin)
    # 左边距： 1143000
    print('右边距：', sec0.right_margin)
    # 右边距： 1143000
    print('上边距：', sec0.top_margin)
    # 上边距： 914400
    # print('下边距：', sec0.bottom_margin)
    # document.add_picture('image-filename.png', width=Inches(4.0))
    # document.add_picture('image-filename.png', width=Inches(5.0))
    w = 15.3
    h = 11
    png = 'image-filename.png'
    p1 = Paper(png, h, w)
    document.add_picture(p1.path, width=Cm(p1.w), height=Cm(p1.h))
    p1 = Paper(png, h, w)
    document.add_picture(p1.path, width=Cm(p1.w), height=Cm(p1.h))


temp1()

if os.path.exists(test_docx):
    os.remove(test_docx)
document.save(test_docx)
