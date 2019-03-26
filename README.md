# attendance_generator v1.0

It is one home work that use python libs to generate attendance reports of our department.

Used libs:
pyodbc : fetch data from Access db.
xlwt : write xls file.

Results may be:

餐补	日期	星期一		     星期二		     星期三		     星期四		     星期五		      星期六		   星期日		  情况说明
部门	姓名	2019-03-18		2019-03-19		2019-03-20		2019-03-21		2019-03-22		2019-03-23	   2019-03-24		
技术部	Name1	08:45	21:19	请假	请假	 08:41	20:38	 08:49	20:34	 08:39	 20:36	 休息	 休息	 休息	 休息	 周一共1次晚餐补贴
技术部	Name2	09:12	21:07	09:13	21:08	09:05	19:05	08:53	19:40	09:10	20:48	休息	休息	休息	休息	周一，周二共2次晚餐补贴
技术部	Name3	08:34	20:25	08:59	20:09	08:35	21:02	09:06	19:42	08:39	18:58	休息	休息	休息	休息	周三共1次晚餐补贴
技术部	Name4	08:34	19:55	08:46	19:30	08:35	19:29	08:43	19:41	08:53	21:17	休息	休息	休息	休息	周五共1次晚餐补贴
