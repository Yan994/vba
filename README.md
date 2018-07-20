# vba

可以录制宏，录制完可以执行

“开发工具”->“插入”->“按钮” 可将指定宏附于按钮

对象、属性、方法、事件
对象是被处理的内容，包括工作簿、工作表、工作表上的单元格区域、图表等等

对象的引用：
Application.Workbooks("mybook.xls").Worksheets("mysheet").Range("a1:d10")
如果引用的单元range是单个的单元格，可用cells(1,1)引用。

属性，对象的各种特征，例如名称、格式。引用对象的属性也要用点.来分割
Worksheet.name
Worksheets(1) 表示Worksheets集合里的第一个工作表
Worksheets("sheet1") 表示Worksheets集合里名为"sheet1"的工作表

msgbox 弹出窗口提示信息

方法，是在对象上执行的某个动作，例如select
立即窗口中选中工作表区域：range("a1:d10").select

事件，即代码要完成的动作

赋值：文本用双引号引起来，日期用#号引起来

定义变量：dim 变量名 as 数据类型
    常量：dim 变量名 as 数据类型 = 变量的值
声明数组：dim/public 数组名 (a to b) as 数据类型
          dim myarr(5) as Integer       '得出数组：0~5
     	  dim myarr(1 to 5,1 to 10) as Integer	  
	
