#!/usr/bin/python3

from datetime import datetime, timedelta, date

import pyodbc
import xlwt
import numpy

def get_query_datetimes(date, query_last_week = True):
    if query_last_week:
        last_week_start = date - timedelta(days = date.weekday() + 7)
        last_week_end = date - timedelta(days = date.weekday() + 1)
    else:
        last_week_start = date - timedelta(days = date.weekday())
        last_week_end = date + timedelta(days = 6 - date.weekday())
    
    start_datetime = datetime(last_week_start.year, last_week_start.month, last_week_start.day, 0, 0, 0)
    end_datetime = datetime(last_week_end.year, last_week_end.month, last_week_end.day, 23, 59, 59)
    
    return start_datetime, end_datetime

def get_connection():
    mdb_path = 'D:/Program Files (x86)/ZKTeco/att2000.mdb'  #replace with your attendance db path
    driver = '{Microsoft Access Driver (*.mdb, *.accdb)}'
    return pyodbc.connect('DRIVER={};DBQ={}'.format(driver,mdb_path))

def get_attendances(start_datetime, end_datetime, target_names):
    connection = get_connection()
    cursor = connection.cursor()
    joined_names = "'{}'".format("','".join(target_names))
    get_checks_sql = "SELECT b.Name,FormatDateTime(CHECKTIME,2),FormatDateTime(MIN(CHECKTIME),4),FormatDateTime(MAX(CHECKTIME),4) \
        FROM CHECKINOUT a INNER JOIN USERINFO b ON a.USERID=b.USERID \
        WHERE CHECKTIME >=#{}# AND CHECKTIME <=#{}# AND b.Name IN ({}) \
        GROUP BY b.Name,FormatDateTime(CHECKTIME,2)".format(start_datetime, end_datetime, joined_names)
    return list(cursor.execute(get_checks_sql))

def get_cell_style(row_index, font_color = 'black'):
    style_template = 'align: wrap yes,vert centre, horiz center; border: left thin,right thin,top thin,bottom thin;'
    if row_index % 2 == 1:
        style_template = style_template + 'pattern: pattern solid, fore_colour gray25;'
    
    style_template = style_template + 'font: name 等线, height 240, color-index {};'.format(font_color)
    return xlwt.easyxf(style_template)

def write_shared_parts(sheet, start_datetime):
    header_style = xlwt.easyxf('align: wrap yes,vert centre, horiz center; \
        pattern: pattern solid, fore_colour light_orange; \
        font: name 等线, height 240, bold on; \
        border: left thin,right thin,top thin,bottom thin') #font_height = font_size * 20

    sheet.write(0, 0, '餐补', header_style)
    sheet.write(0, 1, '日期', header_style)
    sheet.write(1, 0, '部门', header_style)
    sheet.write(1, 1, '姓名', header_style)
    weekdays = ('星期一', '星期二', '星期三', '星期四', '星期五', '星期六', '星期日')
    week_length = len(weekdays)

    for i in range(week_length):
        sheet.write_merge(0, 0, (i + 1) * 2, (i + 1) * 2 + 1, weekdays[i], header_style)
        sheet.write_merge(1, 1, (i + 1) * 2, (i + 1) * 2 + 1, (start_datetime.date() + timedelta(days = i)).isoformat(), header_style)

    sheet.write_merge(0,1,16,16,'情况说明', header_style)
    sheet.col(16).width = 40 * 256  #cell_width = char_count * 256

def write_department_name_parts(sheet, data_start_row, data_row_count, target_names):
    for i in range(data_row_count):
        cell_style = get_cell_style(i)

        row = data_start_row + i
        sheet.write(row, 0, '技术部', cell_style)
        sheet.write(row, 1, target_names[i], cell_style)
        sheet.write(row, 16,'', cell_style)

def get_attendance_time_style(row_index, attendance_datetime, checkin_time_string, checkout_time_string):
    time_format = '%H:%M'
    checkin_warning_time = datetime.strptime('10:00', time_format)
    checkin_time = datetime.strptime(checkin_time_string, time_format)
    checkout_time = datetime.strptime(checkout_time_string, time_format)
    
    warning_style = get_cell_style(row_index, 'red')
    normal_style = get_cell_style(row_index)
    warning_timedelta = timedelta(hours = 10) # 19:00 - 9:00
    bonus_timedelta = timedelta(hours = 12) # 21:30 - 9:00 - dinner_time(30, we have no dinner time now) 
    
    if checkout_time < (checkin_time + warning_timedelta):
        return warning_style, warning_style
    elif checkout_time >= (checkin_time + bonus_timedelta):
        return normal_style, warning_style
    elif checkin_time > checkin_warning_time:
        return warning_style, normal_style
    else:
        return normal_style, normal_style
    
def write_attendance_parts(sheet, data_start_row, data_start_col, start_datetime, target_names, attendances, data_filled_flag):
    for attendance in attendances:
        attendance_name = attendance[0]
        data_row_relative_index = target_names.index(attendance_name)
        row_index = data_start_row + data_row_relative_index

        attendance_datetime = datetime.strptime(attendance[1], '%Y/%m/%d')
        data_col_relative_index = (attendance_datetime.date() - start_datetime.date()).days * 2
        col_index = data_start_col + data_col_relative_index
        
        checkin_time_string = attendance[2]
        checkout_time_string = attendance[3]
        
        checkin_time_style, checkout_time_style = get_attendance_time_style(row_index, attendance_datetime, 
                                                                            checkin_time_string, checkout_time_string)

        sheet.write(row_index, col_index, checkin_time_string, checkin_time_style)
        data_filled_flag[data_row_relative_index, data_col_relative_index] = 1

        sheet.write(row_index, col_index + 1, checkout_time_string, checkout_time_style)
        data_filled_flag[data_row_relative_index, data_col_relative_index + 1] = 1
    
def write_rest_parts(sheet, data_start_row, data_start_col, data_row_count, data_col_count, data_filled_flag):
    for i in range(data_row_count):
        for j in range(data_col_count):
            if data_filled_flag[i,j] == 0:
                cell_style = get_cell_style(i, 'green')
                sheet.write(data_start_row + i, data_start_col + j, '休息', cell_style)
    
def write_data_parts(sheet, start_datetime, target_names, attendances):
    data_start_row = 2
    data_start_col = 2
    data_row_count = len(target_names)
    data_col_count = 14 # weekday * 2
    data_filled_flag = numpy.ones((data_row_count, data_col_count)) * 0

    write_department_name_parts(sheet, data_start_row, data_row_count, target_names)
    write_attendance_parts(sheet, data_start_row, data_start_col, start_datetime, target_names, attendances, data_filled_flag)
    write_rest_parts(sheet, data_start_row, data_start_col, data_row_count, data_col_count, data_filled_flag)

def write_xls_file(start_datetime, end_datetime, target_names, attendances):
    workbook = xlwt.Workbook()
    sheet1 = workbook.add_sheet('Sheet1',cell_overwrite_ok = True)

    write_shared_parts(sheet1, start_datetime)
    write_data_parts(sheet1, start_datetime, target_names, attendances)

    save_attendance_file(workbook, start_datetime, end_datetime)
    
def save_attendance_file(workbook, start_datetime, end_datetime):
    datetime_format = '%Y%m%d'
    start_datetime_format = start_datetime.strftime(datetime_format)
    end_datetime_format =  end_datetime.strftime(datetime_format)
    file_path = 'D:/Docs/技术部/考勤记录/{}-{}技术部考勤情况.xls'.format(start_datetime_format, end_datetime_format)
    workbook.save(file_path)

def main(query_last_week = True):
    start_datetime, end_datetime = get_query_datetimes(date.today(), query_last_week)
    target_names = ('Name1', 'Name2', 'Name3', 'Name4')
    attendances = get_attendances(start_datetime, end_datetime, target_names)
    write_xls_file(start_datetime, end_datetime, target_names, attendances)

main()