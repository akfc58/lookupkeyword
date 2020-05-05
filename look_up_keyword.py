# -*- coding: utf-8 -*-
# Author: 董亮
import os
import time
import docx.document
import re
from win32com import client as w32
print(
    '--------------脚本将临时转换doc文件为原文件名_tmp_新版格式.docx,查找完毕后可选删除临时文件--------------'
)
print('--------------递归查找脚本所在目录&子目录下的所有docx文件包含的所输入的关键词并展示--------------',
      )
print('--------------仅支持一般文本内容如标题、段落查找，不支持图片、表格内容查找--------------',
      end='')


def each_line(file_name):  # 读取一个文件并返回行数和行内容
    line_counter = 0
    each_text = ''
    each_line_with_number = []
    # docx模块返回的paragraphs是一个对象，包含内容和格式等属性
    paras = (docx.Document(file_name)).paragraphs
    for x in range(0, len(paras)):
        # 读取文字属性内容并分段存储，此处通过切片分段而不是分行，所以没有断行问题
        each_text = paras[x].text
        line_counter += 1
        # 生成附带段落编号的段落内容
        i = [line_counter, each_text]
        each_line_with_number.append(i)
    return each_line_with_number


# 遍历脚本所在文件夹的所有指定后缀名文件并返回文件绝对路径的列表
def file_path(suffix='docx'):
    #os.walk遍历并返回一个特殊元组，具体查看os模块文档
    walk = [walk for walk in os.walk(os.getcwd())]  # 使用了工作路径（cwd）
    file_list = []
    # 寻找每个子文件夹的文件并拼接绝对路径
    for scolder_number in range(0, len(walk)):
        for file_name in walk[scolder_number][-1]:
            if file_name.find('.') != -1:
                (_, file_ext_name) = file_name.rsplit('.', 1)
            # 若打开word文档生成了~!开头的隐藏文件，此处代码会出错。
            if file_ext_name == suffix:
                file_list.append(walk[scolder_number][0] + '\\' + file_name)
    return file_list


# 查找文字在单一字符串中的数量和位置并返回
def look_up_str(str_in, look_up_word):
    list_real_loc = []
    list_loc = [i.start() for i in re.finditer(look_up_word, str_in)]
    for loc in list_loc:
        loc = loc + 1
        list_real_loc.append(loc)
    return list_real_loc


# 衔接函数，将所有文本文件处理成单个带行(段)数的字符串
def link_up(look_up_word):
    for each_file_path in file_path():

        for each_line_with_number in each_line(each_file_path):
            # 提取某一文件某一行(段)的内容本身
            each_line_wo_number = each_line_with_number[1]
            word_locs = look_up_str(each_line_wo_number, look_up_word)
            if word_locs:
                print('查找到word文件%s：' % each_file_path)
                print('第%d行有关键词\'%s\'，其位置是:' %
                      (each_line_with_number[0], look_up_word) +
                      str(word_locs))
        print(
            '================================================================')


# 转换doc文档为towhat格式：txt=4, html=10, docx=16， pdf=17
# 若设置convert_or_delete = false,则删除转换后的文件
def convert(towhat=16):
    doc_file_list = file_path('doc')
    for each_doc_file in doc_file_list:
        (doc_path, doc_name) = each_doc_file.rsplit('.', 1)
        # DispatchEx代表独立打开
        word_client = w32.Dispatch('Word.Application')
        word_client.visible = 1
        doc03 = word_client.Documents.Open(each_doc_file)
        new_name = doc_path + '_tmp_新版格式.' + doc_name + 'x'
        try:
            doc03.SaveAs(new_name, towhat)
        except OSError:
            print('转换doc为docx时，保存文件时出现某种错误，请打开程序调试')
        else:
            print('doc转docx保存成功！,保存的文件为%s' % new_name)
        # 每次都打开和关闭
        doc03.Close()
    # 完成后再退出word
    print('转换完毕，开始查找关键字：%s' % user_input_keyword)
    word_client.Quit()

# 删除 转换以后的docx文件，判断依据是是否包含‘_tmp_新版格式.docx’
def del_converted_doc():
    choice = input('请确认是否删除转换后的docx文件！删除操作不可恢复！（输入Y确认，输入其余键退出）\n')
    docx_file_list = file_path()  # file_path()默认查找docx文件
    for each_doc_file in docx_file_list:
        print(each_doc_file)
        if (each_doc_file.rfind('_tmp_新版格式.docx') != -1) and (choice == 'Y'):
            os.remove(each_doc_file)
            print('删除临时转换文件程序执行完毕！')
        else:
            print('上述文件未被删除，请确认。')


def look_up_keyword(keyword, choice=True):  # 选择是否转换文件，默认转换
    if choice:
        convert(16)  # 16代表将doc转换为docx
    else:
        pass
    link_up(keyword)  # 查找所有docx文件
    if choice:
        del_converted_doc()  # 删除所有转换后的临时docx文件

# 输入与输出部分
user_input_keyword = ''
while user_input_keyword != 'QQQ' or user_input_keyword != '':
    user_input_keyword = input(
        '\n输入QQQ退出\n请确认搜索前已关闭所有文档，否则会发生错误！\n请输入脚本所在目录下要查找的关键词：')
    if user_input_keyword == 'QQQ' or user_input_keyword == '':
        break
    # 防止用户输入1和0之外的指令
    try:
        user_choice = int(input('输入1转换doc为docx临时文件，输入0不转换：'))
    except ValueError:
        print('请键入正确数字命令!')
    # 调用主函数
    look_up_keyword(user_input_keyword, user_choice)

print('程序结束')
time.sleep(3)  # 优雅的延时关闭。
