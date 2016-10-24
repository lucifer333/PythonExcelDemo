#可以修改excel,支持xls,貌似不支持xlsx
import xlrd;
from xlutils.copy import copy;

#open existed xls file,文件不存在会报错
oldWb=xlrd.open_workbook(r"F:\demo1.xls",formatting_info=True);
newWb=copy(oldWb);
 
newWs=newWb.get_sheet(0);
newWs.write(31,0,"value31");
newWs.write(29,1,"value1");
newWs.write(30,2,"value1");
 
newWb.save("F:\\demo1.xls");

#xlsx格式,文件不存在会报错,不能使用formatting_info=True,下面这段代码修改的excel会没办法打开
# oldWb2=xlrd.open_workbook(r"F:\demo1.xlsx");
# newWb2=copy(oldWb2);
#   
# newWs2=newWb2.get_sheet(0);
# newWs2.write(31,0,"value31");
# newWs2.write(29,1,"value1");
# newWs2.write(30,2,"value1");
#   
# newWb2.save("F:\\demo1.xlsx");
