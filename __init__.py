from code import Excel
import win32com.client as win32

excel=Excel(r'C:\Users\Balu\win32\sample_xlsm\dataval_navheadings.xlsm')

excel.app=win32.gencache.EnsureDispatch('Excel.Application')

excel.open_workbook()

excel.open_worksheet(sheet_name='List')

excel.display()

excel.turn_off_display_alerts()

excel.set_column_width(start_column=3,size=30)

excel.set_font_format(start_column=3,start_row=15,size=10,style='Bahnschrift SemiCondensed',vertical_alignment=win32.constants.xlTop,horizantal_alignment=win32.constants.xlCenter)

excel.insert_image(start_row=10,start_column=10,img_path=r'C:\Users\Balu\win32\sample_xlsm\dmitry-ratushny-xsGApcVbojU-unsplash.jpg')

excel.set_auto_column_width()

excel.clear_content(start_row=1,start_column=1,end_row=10,end_column=25)

excel.replace_column_content(start_row=10,start_column=1,content=[i for i in range(20)])

excel.replace_row_content(start_row=11,start_column=1,content=[i for i in range(50)])

excel.save_workbook()

list=[[1,2,3,23,4,],[34,4,34,34],[1,2,3,24,34],[1,2,3,34,34],[1,2,3,23,434]]
excel.replace_content(4,2,True,True,list)

excel.color_single_cell(start_row=6,start_column=6,rgb_to_int_value=6556260)

excel.close()