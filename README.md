# vba

可以录制宏，录制完可以执行

“开发工具”->“插入”->“按钮” 可将指定宏附于按钮

----------------------------------------------------------------------------
对象、属性、方法、事件

对象是被处理的内容，包括工作簿、工作表、工作表上的单元格区域、图表等等

对象的引用：
Application.Workbooks("mybook.xls").Worksheets("mysheet").Range("a1:d10")
如果引用的单元range是单个的单元格，可用cells(1,1)引用。
如果mysheet当前是激活的，引用可以简化为[a1:d10]

属性，对象的各种特征，例如名称、格式。引用对象的属性也要用点.来分割
Worksheet.name
Worksheets(1) 表示Worksheets集合里的第一个工作表
Worksheets("sheet1") 表示Worksheets集合里名为"sheet1"的工作表

msgbox 弹出窗口提示信息

方法，是在对象上执行的某个动作，例如select
立即窗口中选中工作表区域：range("a1:d10").select

事件，即代码要完成的动作

----------------------------------------------------------------------------
数据类型              存储空间大小       范围
Byte	                1 个字节        0 到 255
Boolean	                2 个字节        True 或 False
Integer                 2 个字节        -32,768 到 32,767
Long(长整型)	        4 个字节        -2,147,483,648 到 2,147,483,647
Single (单精度浮点型)	4 个字节	    负数时从 -3.402823E38 到 -1.401298E-45；正数时从 1.401298E-45 到 3.402823E38
Double (双精度浮点型)	8 个字节	    负数时从 -1.79769313486232E308 到-4.94065645841247E-324；正数时从4.94065645841247E-324 到 1.79769313486232E308
Currency (变比整型)     8 个字节	    从 -922,337,203,685,477.5808 到 922,337,203,685,477.5807
Decimal	                14 个字节	    没有小数点时为 +/-79,228,162,514,264,337,593,543,950,335，而小数点右边有 28 位数时为 +/-7.9228162514264337593543950335；最小的非零值为 +/-0.0000000000000000000000000001
Date	                8 个字节	    100 年 1 月 1 日 到 9999 年 12 月 31 日
Object	                4 个字节	    任何 Object 引用
String(变长)       10 字节加字符串长度	0 到大约 20 亿
String(定长)	       字符串长度   	1 到大约 65,400
Variant(数字)	        16 个字节	    任何数字值，最大可达 Double 的范围
Variant(字符)	22 个字节加字符串长度	与变长 String 有相同的范围
用户自定义          所有元素所需数目	每个元素的范围与它本身的数据类型的范围相同。
（利用 Type）

--------------------------------------------------------------------------------------

赋值：文本用双引号引起来，日期用#号引起来

定义变量：dim 变量名 as 数据类型
    常量：dim 变量名 as 数据类型 = 变量的值
声明数组：dim/public 数组名 (a to b) as 数据类型
          dim myarr(5) as Integer       '得出数组：0~5
     	  dim myarr(1 to 5,1 to 10) as Integer	  
常用还有static语句、Private语句、Public语句，不同语句定义的变量不同的是它们的作用域不同，具体为：
（1）若在一个过程中包含了一个Dim或Static语句，此时声明的变量作用域为此过程，即本地变量
（2）如果在一个模块的第一个过程之前包含了Dim或Private语句，此时声明的变量作用域为此模块里所有的过程，即模块作用域下的变量
（3）若在一个模块的第一个过程之前包含了Public语句，此时声明的变量作用域为所有模块，即公有变量。
