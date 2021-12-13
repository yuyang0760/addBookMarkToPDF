#!/usr/bin/python
# -*- coding: UTF-8 -*-

import time
import pdfplumber
# from tools import
import re
import time
import xml.dom.minidom as minidom
import fitz

# 页面偏移量,真正的第一页在PDF中是第几页
pageOffset = 9
pdfName = '类型题改.pdf'
bookMarkXmlName = '书签.xml'
bookMarkTxtName = '书签.txt'
bookMarkNotDoTxtName = '书签未整理.txt'


# 处理标题
def dealTitle(title, jishu):
    """处理标题"""
    title = title.strip()
    # 提取前面数字
    n = len(title.split(' ', 2)[0])
    # 如果是2级以上就只用数字
    if jishu >= 2:
        title = title[:n]
    # 如果是1级就用全部
    title = title.replace(' ', '')
    # 当然所有的空格都要去掉
    return title


# 处理文本文档到书签类,最多处理7级书签
def fromTextToBookMark():
    """处理文本文档到书签字典,最多处理7级书签"""
    # 读取文本文件,放入数组
    with open(bookMarkTxtName, 'r', encoding='UTF-8') as f:
        yy_book_mark_zong = {'childBookMark': [], 'title': '', 'startPageNumber': 0, 'endPageNumber': 0, 'pointX': 0,
                             'pointY': 0, 'jishu': -1}
        for line in f.readlines():
            line = line.strip('\n')  # 去掉列表中每一个元素的换行符
            if line.count('\t') == 1:  # ==1 代表一级书签
                line_split = line.split('\t')
                yy_book_mark_zong['childBookMark'].append(
                    {'childBookMark': [], 'title': '', 'startPageNumber': 0, 'endPageNumber': 0, 'pointX': 0,
                     'pointY': 0, 'jishu': 0})
                yy_book_mark_zong['childBookMark'][-1]['title'] = line_split[-2]
                yy_book_mark_zong['childBookMark'][-1]['startPageNumber'] = line_split[-1]
                yy_book_mark_zong['childBookMark'][-1]['jishu'] = line.count('\t')
            if line.count('\t') == 2:  # ==2 代表二级书签
                line_split = line.split('\t')
                yy_book_mark_zong['childBookMark'][-1]['childBookMark'].append({
                    'childBookMark': [], 'title': '', 'startPageNumber': 0, 'endPageNumber': 0, 'pointX': 0,
                    'pointY': 0, 'jishu': 0})
                yy_book_mark_zong['childBookMark'][-1]['childBookMark'][-1]['title'] = line_split[-2]
                yy_book_mark_zong['childBookMark'][-1]['childBookMark'][-1]['startPageNumber'] = line_split[-1]
                yy_book_mark_zong['childBookMark'][-1]['childBookMark'][-1]['jishu'] = line.count('\t')
            if line.count('\t') == 3:  # ==3 代表3级书签
                line_split = line.split('\t')
                yy_book_mark_zong['childBookMark'][-1]['childBookMark'][-1]['childBookMark'].append({
                    'childBookMark': [], 'title': '', 'startPageNumber': 0, 'endPageNumber': 0, 'pointX': 0,
                    'pointY': 0, 'jishu': 0})
                yy_book_mark_zong['childBookMark'][-1]['childBookMark'][-1]['childBookMark'][-1]['title'] = line_split[
                    -2]
                yy_book_mark_zong['childBookMark'][-1]['childBookMark'][-1]['childBookMark'][-1]['startPageNumber'] = \
                    line_split[-1]
                yy_book_mark_zong['childBookMark'][-1]['childBookMark'][-1]['childBookMark'][-1]['jishu'] = line.count(
                    '\t')
            if line.count('\t') == 4:  # ==4 代表4级书签
                line_split = line.split('\t')
                yy_book_mark_zong['childBookMark'][-1]['childBookMark'][-1]['childBookMark'][-1][
                    'childBookMark'].append({
                    'childBookMark': [], 'title': '', 'startPageNumber': 0, 'endPageNumber': 0, 'pointX': 0,
                    'pointY': 0, 'jishu': 0})
                yy_book_mark_zong['childBookMark'][-1]['childBookMark'][-1]['childBookMark'][-1]['childBookMark'][-1][
                    'title'] = line_split[-2]
                yy_book_mark_zong['childBookMark'][-1]['childBookMark'][-1]['childBookMark'][-1]['childBookMark'][-1][
                    'startPageNumber'] = line_split[-1]
                yy_book_mark_zong['childBookMark'][-1]['childBookMark'][-1]['childBookMark'][-1]['childBookMark'][-1][
                    'jishu'] = line.count('\t')
            if line.count('\t') == 5:  # ==5 代表5级书签
                line_split = line.split('\t')
                yy_book_mark_zong['childBookMark'][-1]['childBookMark'][-1]['childBookMark'][-1][
                    'childBookMark'].append({
                    'childBookMark': [], 'title': '', 'startPageNumber': 0, 'endPageNumber': 0, 'pointX': 0,
                    'pointY': 0, 'jishu': 0})
                yy_book_mark_zong['childBookMark'][-1]['childBookMark'][-1]['childBookMark'][-1]['childBookMark'][-1][
                    'title'] = line_split[-2]
                yy_book_mark_zong['childBookMark'][-1]['childBookMark'][-1]['childBookMark'][-1]['childBookMark'][-1][
                    'startPageNumber'] = line_split[-1]
                yy_book_mark_zong['childBookMark'][-1]['childBookMark'][-1]['childBookMark'][-1]['childBookMark'][-1][
                    'jishu'] = line.count('\t')
            if line.count('\t') == 6:  # ==6 代表6级书签
                line_split = line.split('\t')
                yy_book_mark_zong['childBookMark'][-1]['childBookMark'][-1]['childBookMark'][-1]['childBookMark'][-1][
                    'childBookMark'].append({
                    'childBookMark': [], 'title': '', 'startPageNumber': 0, 'endPageNumber': 0, 'pointX': 0,
                    'pointY': 0, 'jishu': 0})
                yy_book_mark_zong['childBookMark'][-1]['childBookMark'][-1]['childBookMark'][-1]['childBookMark'][-1][
                    'childBookMark'][-1]['title'] = line_split[-2]
                yy_book_mark_zong['childBookMark'][-1]['childBookMark'][-1]['childBookMark'][-1]['childBookMark'][-1][
                    'childBookMark'][-1]['startPageNumber'] = line_split[-1]
                yy_book_mark_zong['childBookMark'][-1]['childBookMark'][-1]['childBookMark'][-1]['childBookMark'][-1][
                    'childBookMark'][-1]['jishu'] = line.count('\t')
            if line.count('\t') == 7:  # ==7 代表7级书签
                line_split = line.split('\t')
                yy_book_mark_zong['childBookMark'][-1]['childBookMark'][-1]['childBookMark'][-1]['childBookMark'][-1][
                    'childBookMark'][-1][
                    'childBookMark'].append({
                    'childBookMark': [], 'title': '', 'startPageNumber': 0, 'endPageNumber': 0, 'pointX': 0,
                    'pointY': 0, 'jishu': 0})
                yy_book_mark_zong['childBookMark'][-1]['childBookMark'][-1]['childBookMark'][-1]['childBookMark'][-1][
                    'childBookMark'][-1][
                    'childBookMark'][-1]['title'] = line_split[-2]
                yy_book_mark_zong['childBookMark'][-1]['childBookMark'][-1]['childBookMark'][-1]['childBookMark'][-1][
                    'childBookMark'][-1][
                    'childBookMark'][-1]['startPageNumber'] = line_split[-1]
                yy_book_mark_zong['childBookMark'][-1]['childBookMark'][-1]['childBookMark'][-1]['childBookMark'][-1][
                    'childBookMark'][-1][
                    'childBookMark'][-1]['jishu'] = line.count('\t')
    return yy_book_mark_zong


# 打印书签
def printBookMark(book_mark):
    """打印书签"""
    for p in book_mark:
        # 打印书签
        print('\t' * (p['jishu'] - 1), p['title'], p['startPageNumber'], p['jishu'])
        # 递归打印子书签
        if len(p['childBookMark']) >= 1:
            printBookMark(p['childBookMark'])


# 在pdf中找到书签位置
def findBookMarkInPDF(book_mark):
    """在pdf中找到书签位置"""
    # 读取pdf并选择对应的页数
    with pdfplumber.open("类型题改带书签.pdf") as pdf:

        for p in book_mark:
            # 打印书签
            print('\t' * (p['jishu'] - 1), p['title'], p['startPageNumber'], p['jishu'])
            # 提取这一页的单词
            page_height = pdf.pages[int(p['startPageNumber']) + pageOffset - 2]
            words = pdf.pages[int(p['startPageNumber']) + pageOffset - 2].extract_words(x_tolerance=5, y_tolerance=3,
                                                                                        keep_blank_chars=False)
            # 定义是否找到
            is_find = False
            # 提取文本
            for word in words:
                if dealTitle(p['title'], p['jishu']) in word['text'].replace(" ", ""):
                    # 如果找到,标记并跳出循环
                    is_find = True
                    # print('')
                    p['pointX'] = word['x0']
                    p['pointY'] = word['top']
                    break
            if not is_find:
                print(p['title'].replace(" ", "") + "没找到")
            # 递归子书签
            if len(p['childBookMark']) >= 1:
                findBookMarkInPDF(p['childBookMark'])
    return book_mark


# 格式化书签
def formatBookMark():
    """格式化书签"""
    book_mark_lines = []
    with open(bookMarkNotDoTxtName, 'r', encoding='UTF-8') as f:
        for line in f.readlines():
            line_arr = line.strip().split('\t')  # 把末尾的'\n'删掉 再用\t分割
            pattern = re.compile(r'^[0-9]{1,2}(\.[0-9]{1,2})?|^第[一二三四五六七八九十]{1,2}章')
            top_number = pattern.search(line_arr[0]).string
            top_title = line_arr[1][len(top_number):]
            book_mark_lines.append(
                '%s %s%s\t%s%s' % ('\t' * (top_number.count(r'.')), top_number, top_title, line_arr[1], '\n'))
    with open(bookMarkTxtName, 'w', encoding='UTF-8') as f:
        f.writelines(book_mark_lines)


# 递归创建文档xml书签
def createBookMarkXml(book_mark, page_offset, book_mark_xml_name):
    """创建标签xml文件"""
    now = time.strftime("%Y-%m-%d %H:%M:%S")

    # 根目录
    dom = minidom.getDOMImplementation().createDocument(None, 'PDF信息', None)
    root = dom.documentElement
    root.setAttribute('程序名称', 'PDFPatcher')
    root.setAttribute('程序版本', '0.0.3')
    root.setAttribute('导出时间', now)
    root.setAttribute('PDF文件位置', '--')

    # 度量单位
    du_liang_element = dom.createElement('度量单位')
    du_liang_element.setAttribute('单位', '点')
    root.appendChild(du_liang_element)
    # 文档书签
    wen_dang_element = dom.createElement('文档书签')
    # 递归创建文档子书签
    createWenDangBookMarkXml(dom, book_mark, wen_dang_element, page_offset)
    root.appendChild(wen_dang_element)
    # 页码样式
    ye_ma_element = dom.createElement('页码样式')
    root.appendChild(ye_ma_element)
    # for i in range(5):
    #     element = dom.createElement('Name')
    #     element.appendChild(dom.createTextNode('default'))
    #     element.setAttribute('age', str(i))
    #     root.appendChild(element)
    # 保存文件
    with open(book_mark_xml_name, 'w', encoding='utf-8') as f:
        dom.writexml(f, addindent='\t', newl='\n', encoding='utf-8')


# 递归创建文档子书签
def createWenDangBookMarkXml(dom, book_mark, wen_dang_element, page_offset, pageHeight):
    """递归创建子书签xml文件"""

    # for i in range(5):
    #     element = dom.createElement('Name')
    #     element.appendChild(dom.createTextNode('default'))
    #     element.setAttribute('age', str(i))
    #     root.appendChild(element)
    # 保存文件
    for p in book_mark:
        # 打印书签
        print('\t' * (p['jishu'] - 1), p['title'], p['startPageNumber'], p['jishu'])

        element = dom.createElement('书签')

        element.setAttribute('文本', p['title'])
        element.setAttribute('默认打开', '是')
        element.setAttribute('动作', '转到页面')
        element.setAttribute('页码', str(int(p['startPageNumber']) + page_offset - 1))
        element.setAttribute('显示方式', '坐标缩放')
        element.setAttribute('左', str(p['pointX'] + 40))
        element.setAttribute('上', str(849 - p['pointY']))
        element.setAttribute('比例', "0")
        wen_dang_element.appendChild(element)

        # 递归打印子书签
        if len(p['childBookMark']) >= 1:
            createWenDangBookMarkXml(dom, p['childBookMark'], element, page_offset)
    return wen_dang_element


if __name__ == '__main__':
    # # 把书签未整理,整理成有一定格式的书签到书签.txt
    formatBookMark()
    # # 处理文本文档到书签字典,最多处理7级书签
    book_mark_class = fromTextToBookMark()['childBookMark']
    # # 在pdf中找到书签位置,完善字典,yy_book_mark是一个列表组数
    yy_book_mark = findBookMarkInPDF(book_mark_class)
    # # 将书签字典数组导出为xml文件,然后用PDFPatcher软件处理
    createBookMarkXml(yy_book_mark, pageOffset, bookMarkXmlName)
    # 在PDF中添加书签
    print('完成')
