class Excel:
    def display(self):
        self.app.Visible=True
        
    def turn_off_display_alerts(self):
        self.app.DisplayAlerts=False
        
    def turn_on_display_alerts(self):
        self.app.DisplayAlerts=True
        
    def display_alerts(self, action=False):
        self.app.DisplayAlerts=action
    
    def hide(self):
        self.app.Visible=False
        
    def open_workbook(self):
        self.workbook=self.app.Workbooks.Open(self.path)
        
    def open_worksheet(self,sheet_name=''):
        self.worksheet=self.workbook.Sheets[sheet_name]
        
    def add_workbook(self):
        self.workbook=self.app.Workbooks.Add()
        
    def add_worksheet(self,sheet_name=''):
        self.worksheet=self.workbook.Worksheets.Add()
        self.worksheet.Name=sheet_name
    
    def replace_content(self,start_row=1,start_column=1,row_wise=False,column_wise=False,content=[]):
        if row_wise==True:
            if column_wise==True:
                for i in range(len(content)):
                    for j in range(len(content[i])):
                        self.worksheet.Cells(start_row+i,start_column+j).Value=content[i][j]
            else:
                for i in range(len(content)):
                    self.worksheet.Cells(start_row,start_column+i).Value=content[i]    
        elif column_wise==True:
            for i in range(len(content)):
                self.worksheet.Cells(start_row+i,start_column).Value=content[i]    
        else:
            self.worksheet.Cells(start_row,start_column).Value=content
            
    def insert_image(self, start_row=1,start_column=1,img_path='',height=0,width=0,left=0,top=0):
        cell=self.worksheet.Cells(start_row,start_column)
        pic=self.worksheet.Pictures().Insert(img_path)
        if left>0:
            pic.Left = cell.Left + left
        if top>0:
            pic.Top = cell.Top + top
        if height>0:
            pic.Height=height
        if width>0:
            pic.Width=width

    def replace_row_content(self,start_row=1,start_column=1,content=[]):
        for i in range(len(content)):
            self.worksheet.Cells(start_row,start_column+i).Value=content[i]
    def replace_column_content(self,start_row=1,start_column=1,content=[]):
        for i in range(len(content)):
            self.worksheet.Cells(start_row+i,start_column).Value=content[i]
    
    def replace_single_cell_content(self,start_row=1,start_column=1,content=None):
        self.worksheet.Cells(start_row,start_column).Value=content
        
    def set_font_format(self,start_row=1,start_column=1,style='Calibri',size=12,vertical_alignment='',horizantal_alignment=''):
        self.worksheet.Cells(start_row,start_column).Font.Name=style
        self.worksheet.Cells(start_row,start_column).Font.Size=size
        self.worksheet.Cells(start_row,start_column).HorizontalAlignment = horizantal_alignment
        self.worksheet.Cells(start_row,start_column).VerticalAlignment = vertical_alignment
        
    def clear_content(self,start_row=1,start_column=1,end_row=1,end_column=1):
        for i in range(start_row,end_row+1):
            for j in range(start_column,end_column+1):
                self.worksheet.Cells(i,j).Value=None
                
    def color_single_cell(self,start_row=1,start_column=1,rgb_to_int_value=0):
        self.worksheet.Cells(start_row,start_column).Interior.Color = int(rgb_to_int_value)
    
    def set_column_width(self, start_column=1,size=1):
        self.worksheet.Columns(start_column).ColumnWidth=size
        
    def set_auto_column_width(self):
        self.worksheet.Columns.AutoFit()
        
    def set_row_height(self, start_row=1,size=1):
        self.worksheet.Rows(start_row).RowHeight=size
        
    def set_auto_row_height(self):
        self.worksheet.Rows.AutoFit()
        
    def save_workbook(self):
        self.workbook.Close(SaveChanges=1)
    
    def close(self):
        self.app.Quit()
    
    def __init__(self, path):
        self.worksheet=''
        self.workbook=''
        self.path=path
        self.app=''