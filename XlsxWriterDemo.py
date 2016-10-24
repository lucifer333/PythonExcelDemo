#encoding:utf_8
#只能写文件，不能修改文件,且不支持xls
import xlsxwriter

workbook=xlsxwriter.Workbook(r"F:\demo1.xls")       #创建一个excel文件
worksheet=workbook.add_worksheet()                  #创建一个工作表对象
worksheet.set_column("A:A",20)                      #定义第一列（A）宽度为20像素
bold=workbook.add_format({"bold":True})             #定义一个加粗的格式对象

worksheet.write("A1", "Hello")                      #"A1"单元格填入"Hello"
worksheet.write("A2","World", bold)
worksheet.write("B2", u"中文测试", bold)
worksheet.write(2, 0, 32)                           #用行列表示法写入数字"32"和"35.5"
worksheet.write(3, 0, 35.5)                         #行列表示法的单元格下标以0作为起始值，'3，0'等价于'A3'
worksheet.write(4, 0, "=SUM(A3:A4)")                #求A3、A4的和，并将结果写入"4,0"，即A5
#worksheet.insert_image("B7", r"F:\IMG_3010.JPG")                    #在"B5"单元格插入图片

workbook.close()                                    #关闭excel文件
