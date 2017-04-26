# -*- coding: utf-8 -*-
##考勤数据处理脚本
##第一步输入原始数据，运行pre_process(filename1)
##第二步，对原始数据进行颜色标注，运行data_process(filename1,filename2)
##第三步，对照请假邮件，研一课表，对步骤二输出的文件进行手动修正
##第四部，结果统计，运行result = statistic(filename2,score)
##返回结果 result是一个n*4列表，第一列是score，第二列姓名，第三列缺卡次数
##第四列迟到次数，按照score逆序排列，score＝缺卡次数＋迟到次数/2
##score越大，考勤越差。
import sys
import xlrd
import xlwt
from xlutils.copy import copy
import json

reload(sys)
sys.setdefaultencoding('utf8')
data_begin = 5 #正常打卡数据开始列号
data_end = 10 #正常打卡数据结束列号
data_cols = 11 #正常打卡数据共有11列
#寻找数据的正确列号
def search_col(value,time):
    n = len(time)
    for i in range(n):
        if (value <= time[i]):
            break
    return (data_begin + i - 1)
#数据对齐
def pre_process(filename):
    #读取表格
    exl_file = xlrd.open_workbook(filename)
    wb = copy(exl_file)
    #通过get_sheet()获取的sheet有write()方法
    ws = wb.get_sheet(0)  #1代表是写到第几个工作表里，从0开始算是第一个。

    sheet = exl_file.sheet_by_index(0)
    time = ['00:00:00','10:30:00','12:30:00','16:30:00',
    '18:30:00','20:30:00','23:59:00']
    cols = sheet.ncols
    rows = sheet.nrows
    k = len(time)-1
    #从后向前遍历列
    #将数据对齐到应该在的列
    for col in range(cols-1, data_begin-1, -1):
        for row in range (1, rows):
            value = sheet.cell(row,col).value.encode('utf-8')
            if value:
                if col > data_cols-1:
                    col_should = search_col(value,time)
                    ws.write(row,col_should,value)
                    ws.write(row,col,'')

                else:
                    if ((time[k-1] <= value) & (value <= time[k])):
                        continue
                    else:
                        col_should = search_col(value,time)
                        #print col_should
                        ws.write(row,col_should,value)
                        ws.write(row,col,'')
                        #print col, col_should
        if col <= data_cols-1:
            if k > 1:
                k -= 1
    wb.save(filename)
#标记迟到早退以及缺勤
#处理结果写入新的表格中
def data_process(filename1,filename2):
    #建立写入文件
    file=xlwt.Workbook()
    sheet=file.add_sheet('result',cell_overwrite_ok=True)

    #读取预处理过的文件
    data = xlrd.open_workbook(filename1)
    table = data.sheet_by_index(0)
    cols,rows = table.ncols, table.nrows
    time = [['09:10:10','10:00:00'],['11:00:00','11:25:00'],
    ['14:10:00','16:00:00'],['17:00:00','17:25:00'],
    ['19:00:00','20:30:00'],['20:00:00','21:25:00']]
    #黄色样式
    yellow = xlwt.XFStyle()
    pattern1 = xlwt.Pattern()
    pattern1.pattern = 1
    pattern1.pattern_fore_colour = 5 # 黄色
    yellow.pattern = pattern1
    #红色样式
    red = xlwt.XFStyle()
    pattern2 = xlwt.Pattern()
    pattern2.pattern = 1
    pattern2.pattern_fore_colour = 2 # 红色
    red.pattern = pattern2
    #先将姓名、打卡时间、标号写入新表格
    for col in range(2,5):
        for row in range(rows):
            value = table.cell(row,col).value
            sheet.write(row,col-2,value)
    #边写入打卡时间，边标记迟到、缺勤
    for col in range(data_begin, data_end+1):
        index = col % 5
        if col == data_end:
            index = 5
        area = time[index]
        for row in range(rows):
            value = table.cell(row,col).value.encode('utf-8')
            if value:
                #迟到，黄色
                if ((area[0] < value) & (value < area[1])):
                    sheet.write(row,col-2,value,yellow)
                else:
                    sheet.write(row,col-2,value)
            #没有打卡，红色
            else:
                    sheet.write(row,col-2,value,red)
    #file1.save('result.xls')
    file.save(filename2)
#判断单元格前景颜色，计数迟到和缺勤
def count(book,sheet, row, cols, red, yellow):
    for col in range(3,cols):
        #当前单元格前景色
        xfx = sheet.cell_xf_index(row , col)
        xf = book.xf_list[xfx]
        bgx = xf.background.pattern_colour_index
        #相邻单元格前景色
        xfx_pre = sheet.cell_xf_index(row , col-1)
        xf_pre = book.xf_list[xfx_pre]
        bgx_pre = xf_pre.background.pattern_colour_index
        if ((bgx ==2) & (bgx_pre == 2)):
            red += 1
        elif ((bgx == 2) & (bgx_pre != 2)):
            red += 1
        elif (bgx == 5):
            yellow += 1
    return [red,yellow]
#统计每个人的出勤
#score＝缺勤次数＋迟到次数＊0.5
#score越大表示考勤越不合格
def statistic(filename,score_list):
    data = xlrd.open_workbook(filename, formatting_info=1)
    sheet = data.sheet_by_index(0)
    red_list, yellow_list, name_list = [],[],[]
    wb = copy(data)
    ws = wb.get_sheet(0)
    cols, rows = sheet.ncols, sheet.nrows
    person = sheet.cell(1,0).value.encode('utf-8')
    color = [0,0] #color = [red, yellow]
    start = 1
    ws.write(0,cols,'absent')
    ws.write(0,cols+1,'late')
    ws.write(0,cols+2,'score')
    #统计缺勤和迟到并打分
    for row in range(1,rows):
        person_now = sheet.cell(row,0).value.encode('utf-8')
        if (person == person_now):
            color = count(data,sheet,row,cols,color[0],color[1])
        else:
            name_list.append(person)
            red_list.append(color[0])
            yellow_list.append(color[1])
            score = color[0] + color[1]/2
            score_list.append(score)
            ws.write_merge(start,row-1,cols,cols,color[0])
            ws.write_merge(start,row-1,cols+1,cols+1,color[1])
            ws.write_merge(start,row-1,cols+2,cols+2,float(score))
            start = row
            person = person_now
            color = count(data,sheet, row, cols, 0, 0)
        if (row == rows-1):
            name_list.append(person)
            red_list.append(color[0])
            yellow_list.append(color[1])
            score = color[0] + color[1]/2
            score_list.append(score)
            ws.write_merge(start,row-1,cols,cols,color[0])
            ws.write_merge(start,row-1,cols+1,cols+1,color[1])
            ws.write_merge(start,row-1,cols+2,cols+2,float(score))
    #将所有信息整合到一个list中
    result_list = []
    for i in range (len(name_list)):
        temp = [score_list[i],name_list[i],red_list[i],yellow_list[i]]
        result_list.append(temp)
    #降序
    result = [result_list[0]]
    flag = False

    for j in range(1,len(name_list)):
        temp1 = result_list[j][0]
        for i in range(len(result)):

            if (result[i][0] < temp1):
                flag = True
                break
            else:
                flag = False
        if flag:
            result.insert(i,result_list[j])
        else:
            result.append(result_list[j])
    wb.save(filename)
    return result
if __name__ == '__main__':
    #原始文件
    filename1 = 'Daily Log(20170423_1651).xls'
    #处理结果
    filename2 = '2017_4.xls'
    score = []
    #预处理，数据对齐
    pre_process(filename1)
    #颜色标注
    data_process(filename1,filename2)
    ###result = [name_list, score_list, red_list,yellow_list]
    #结果统计，在统计结果之前需要对照请假、研一上课课表，对数据进行修正
    result = statistic(filename2,score)
    print ['score', 'name', 'absent','late']
    print "note: score = absent + late/2"
    for i in range(len(result)):
        print(repr(result[i]).decode('string-escape'))
