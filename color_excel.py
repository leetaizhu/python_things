
# -*- coding: utf-8 -*-
import xlrd
import xlwt
#新建一个excel文件
file=xlwt.Workbook()
#新建一个sheet
table=file.add_sheet('sheet name',cell_overwrite_ok=True)
for i in range(0,256):
        stylei= xlwt.XFStyle()            #初始化样式
        patterni= xlwt.Pattern()          #为样式创建图案
        patterni.pattern=1                #设置底纹的图案索引，1为实心，2为50%灰色，对                                            应为excel文件单元格格式中填充中的图案样式
        patterni.pattern_fore_colour=i    #设置底纹的前景色，对应为excel文件单元格格式                                            中填充中的背景色
        #patterni.pattern_back_colour=35   #设置底纹的背景色，对应为excel文件单元格格式                                            中填充中的图案颜色
        stylei.pattern=patterni           #为样式设置图案
        table.write(i,0,i,stylei)         #使用样式

file.save('/Users/lee/Desktop/colour.xls')

book = xlrd.open_workbook("colour.xls", formatting_info=1)
sheet = book.sheet_by_index(0)

rows, cols = sheet.nrows, sheet.ncols
#print "Number of rows: %s Number of cols: %s" % (rows, cols)
for row in range(3):
    for col in range(cols):
        xfx = sheet.cell_xf_index(row, col)
        xf = book.xf_list[xfx]
        bgx = xf.background.pattern_colour_index
        if bgx == 5:
            print sheet.cell(row,col).value, bgx,row,col
xfx = sheet.cell_xf_index(3 , 7)
xf = book.xf_list[xfx]
bgx = xf.background.pattern_colour_index
print sheet.cell(row,col).value, bgx,row,col
xfx = sheet.cell_xf_index(3 , 5)
xf = book.xf_list[xfx]
bgx = xf.background.pattern_colour_index
print sheet.cell(row,col).value, bgx,row,col
