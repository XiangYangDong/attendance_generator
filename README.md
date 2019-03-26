# attendance_generator v1.0

It is one home work that use python libs to generate attendance reports of our department.


Used libs:
--------------------------------------------

pyodbc : fetch data from Access db.

xlwt : write xls file.


Results may be:
--------------------------------------------

餐补	日期	星期一		    ...		星期日		  情况说明

部门	姓名	2019-03-18		...	   2019-03-24		

技术部	Name1	08:45	21:19	...	   休息	 休息	 周一共1次晚餐补贴

技术部	Name2	09:12	21:07	...	   休息	休息	周一，周二共2次晚餐补贴

技术部	Name3	08:34	20:25	...	   休息	休息	周三共1次晚餐补贴

技术部	Name4	08:34	19:55	...	   休息	休息	周五共1次晚餐补贴

