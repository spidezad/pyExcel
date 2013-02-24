'''
###########################################

 pyExcel module
 Author: Tan Kok Hua (Guohua tan)
 Email: Guohua.tan@seagate.com/kokhua81@gmail.com
 Revised date: Feb 24 2012

############################################

 Notes:
     Derived/modified from original ExcelLink.py found on internet

 Changes:
    Feb 17 2013: Remove the get xls avaliable function
               : Allow copy data to excel function to have transpose function 
    Dec 25 2012: Add in chart formatter function
    Dec 17 2012: Fix insert column function to take in both integer and alphabet
    Dec 12 2012: Amend get_shtname_and_lastdata_row_col_str function to have selective output
    Dec 11 2012: Add in cell replace function
    Dec 08 2012: Amend paste_format_fr_ext_xls function to cater for selected range copy (not working)
    Dec 07 2012: Add in paste_format_fr_ext_xls function
    Oct 20 2012: Add in copy_data_fr_csv function
    May 09 2012: Add in convert cell to percent function
    May 01 2012: Amend the highlight function to include highlight range 
    Apr 30 2012: Add in cell merge function
    Oct 30 2011: Comment out print statment in add_sheet and search_filename function
               : Remove version_history function
    Sep 30 2011: Add in paste_format function --> pending checks on function
    Jul 12 2011: resolve bug by removing version in repr
    Apr 17 2011: Allow check of version (2003,07) when saving. For 2007 version, save as xl97/95
    Feb 25 2011: Remove setcellalign function (can use setrangealign function)
               : Add in check_close_excel_wb (close an open workbook)
    Feb 23 2011: Add in autofit_column_width function
    Feb 19 2011: Solve bug in find_keyword function
    Feb 17 2011: Solve bug in the add_sheet function
    Feb 16 2011: Amend activate_sel function
    Feb 12 2011: Amend setrangealign function, set alignment= 'center'
               : Rename openfile_GUI to openxlfile_GUI, remove easygui dependency
    Feb 10 2011: Amend save_as_txt/save as  csv, textformat should be integer
    Feb 07 2011: Remove Sel_coordinates, active_address, active selection function, get_selectadd 
               : Replaced all with get_active_select_add (underconstruction for mulitple select)
    Feb 04 2011: Modify Wrap text function
               : Rename Editrange function to editrange
    Feb 03 2011: Modify the clear range function
               : Modify Editrange function (underconstruction: check can cater for single cell)
    Feb 02 2011: Reform getcell,getrange,setcell function. (convert all tuple to str for simplicity)
               : Modify setcellalign function
               : Replace addnewsheetafter and addnewsheetbefore with addnewsheet function
               : Rename copyrange function to copy_paste_range 
    Feb 01 2011: Add in convert_alphabet, convert_tuple_to_cell_str function
    Jan 10 2011: Remove chart_scatter function, improve chart_type function to default scatter
    Jan 09 2011: Amend module name to pyExcel
               : Remove Excel pass
               : Remove excelapp function


Changes to ExcelLInk:
    Ver 2.4
    Aug 21 2010 : Add in function to import data from text file
    July 17 2010: Add in function to get data from  multiple selection
    July 14 2010: Amend the paste select command so it can paste with no formatting.
                : Edit the paste special command (allows select, copy and paste in one function )
                : Add in the show_warning function and the delete sheet function
    Ver 2.3
    Aug 23 09 --> able to save CSV format to xls format
    Aug 24 09 --> able to bypass the x-axis for certain graph series
              --> add the series charttype function
    Aug 25 09 --> add the openfile GUI and filename convert

    Ver 2.2
    Jul 28 09--> add del column function,get_selectadd --> if num format False, may use for multi select

    Ver 2.1
    Jun 29 09--> make changes to chart_series_add --> include str for series label
    May 31  09 --> edit get selectadd function --> results in tuple or in str, chart type function
    May 30  09 --> add convert cell to pts function (required pass) for charts
    May 16-09 --> add search file, search and add sheet function
    Apr 29 2009 --> debug for active address function
    Apr 22 2009 --> Required Excel Pass for certain function

    Ver 1.0
    UPdate 31Mar 2009 --> Chart functions
    Update 27Mar 2009 --> Get active address, chart functions
    Update 22Mar 2009 --> graphical functions --> format scale, legend, add series, delete series
    Update 03Mar 2009 --> create the series count function
    Update 22Feb 2009 --> add, delete series for chart functions
    Update 04Feb 2009 --> add in convert number to column function
    Update 24Jan 2009 --> add the activate selection function,copy sel and paste sel function
    Update 21Jan 2009 -> add in the selection coordinates function
    Update 17Jan09 -> add the border function and add column function
    Update 11Jan 09 -> add the set range alignment functions
    Update Jan 2009 -> add the copy rows functions, add the edit range functions
    Update Dec 2008 --> add freeze panel, selection address function
    Update Dec 2008 --> copy function
    Update Oct 2008 --> Wrap text

 Underconstruction:
    to add in autofit column Columns("C:C").EntireColumn.AutoFit
    remove setcellalign, editcell if no use.
    sorting
    transpose
    read cell attributes as template
    improve on one touch compile
    get cell dec pl --> have to access cell by cell --> long access time??
    using format painter??--> use of paste special formatSelection.PasteSpecial Paste:=xlPasteFormats
    check paste special and paset_format to
    check activate_sel to see if any problem
    to check on the paste format only??
    search multiple of same keyword
    #excel --> have transpose list to column based?
    #function to columnize the list of columns input
    #format also include failure pareto
    #excel --> this does not work??xlFile.print_avaliable_excel_function()

 Learning:
     xlapp.GetSaveAsFilename() --> can open save dialog
     office 2007 version 12
     get last cell with data
     ww = xlbook.ActiveSheet.Cells.SpecialCells(11), ww.Row, ww.Column
     http://www.yogeshguptaonline.com/2009/09/excel-macros-find-last-row.html
     
'''

import win32com.client.dynamic 
import glob
try:
    import pyET_tools.win_program_manipulate as win
except:
    print 'Win function not installed'
    print "Some function will be disabled if module is not present"

class UseExcel(object): 
    '''
        Python Excel Interface. It provides methods for accessing the 
        basic functionality of MS Excel 97/2000 from Python.
        
        This interface uses dynamic dispatch objects. All necessary constants 
        are embedded in the code. There is no need to run makepy.py. 
    '''
    
    
    def __init__(self, fileName=None):
        
        '''
            if FileName == None, will open a new workbook
            e.g. xlFile = useExcel("e:\\python23\myfiles\\testExcel1.xls") 
        '''
        
        self.xlapp = win32com.client.dynamic.Dispatch("Excel.Application") 
        if fileName:                 
            self.xlbook = self.xlapp.Workbooks.Open(fileName) 
        else: 
            self.xlbook = self.xlapp.Workbooks.Add()

        #parameters for pivot table used
        self.pivot_table_name = ''
        self.rowfield_list = [] #in seq
        self.colfield_list = []
        self.datafield = ''

        
    def __repr__(self):
        return 'ExcelLink module'

    def save(self, newfileName=None):
        
        if newfileName:
            if float(self.xlapp.Version) < 12:  
                self.xlbook.SaveAs(newfileName)
            else:
                self.xlbook.SaveAs(newfileName,FileFormat = -4143)#for ms 2007 save format as 97/95
        else: 
            self.xlbook.Save()

            
    def save_fr_csv(self,newfileName):
        '''Save csv file as xls file'''
        self.xlapp.ActiveWorkbook.SaveAs(Filename =newfileName,FileFormat= 43 )#43 means Excel97/95

    def save_as_txt(self,newfileName):
        '''Save xls file as txt file'''
        self.xlapp.ActiveWorkbook.SaveAs(Filename =newfileName,FileFormat= -4158 )#-4158 txt file

    def close(self): 
        self.xlbook.Close(SaveChanges=False) 
        del self.xlapp 

    def show(self): 
        self.xlapp.Visible = True 

    def hide(self): 
        self.xlapp.Visible = False 

    def convert_alphabet(self,col_num):
        '''Convert xls column number to Column Alphabet'''
        if col_num > 26:
            return chr(int((col_num - 1) / 26) + 64) + chr(((col_num - 1) % 26) + 65)
        else:
            return chr(col_num + 64)

    def convert_tuple_to_cell_str(self,temp_tuple):#can be 2 or 4 tuple
        if len(temp_tuple) ==2:
            return self.convert_alphabet(temp_tuple[1])+ str(temp_tuple[0])

        elif len(temp_tuple) ==4:
            return self.convert_alphabet(temp_tuple[1])+ str(temp_tuple[0]) + ':' + self.convert_alphabet(temp_tuple[3])+ str(temp_tuple[2])

        else:
            print 'unable to process data'#raise error?


    def getcell(self, sheet, cellAddress): 
        '''
            Get value of one cell. 
            sheet       -   name of the excel worksheet 
            cellAddress -   tuple of integers (row, cloumn) or string "ColumnRow" 
                            e.g. (3,4) or "D3" 
        '''
        
        sht = self.xlbook.Worksheets(sheet)
        if (isinstance(cellAddress,tuple)):
            cellAddress =self.convert_tuple_to_cell_str(cellAddress)

        return sht.Range(cellAddress).Value 
        
    def setcell(self, sheet, value, cellAddress, fontStyle=("Regular",), fontName="Arial", fontSize=8, fontColor=1): 
        '''
            Set value of one cell. 

            sheet       -   name of the excel worksheet 
            value       -   The cell value. it can be a number, string etc. 
            cellAddress -   tuple of integers (row, cloumn) or string "ColumnRow" 
                            e.g. (3,4) or "D3" 
            fontStyle   -   tuple. Combination of Regular, Bold, Italic, Underline 
                            e.g. ("Regular", "Bold", "Italic") 
            fontColor   -   ColorIndex. Refer ColorIndex property in Microsoft 
            Excel Visual Basic Reference 
        '''         
        sht = self.xlbook.Worksheets(sheet)
        if (isinstance(cellAddress,tuple)):
            cellAddress =self.convert_tuple_to_cell_str(cellAddress)

        sht.Range(cellAddress).Value = value 
        sht.Range(cellAddress).Font.Size = fontSize 
        sht.Range(cellAddress).Font.ColorIndex = fontColor 
        for i, item in enumerate(fontStyle): 
            if (item.lower() == "bold"): 
                sht.Range(cellAddress).Font.Bold = True 
            elif (item.lower() == "italic"):     
                sht.Range(cellAddress).Font.Italic = True 
            elif (item.lower() == "underline"):     
                sht.Range(cellAddress).Font.Underline = True 
            elif (item.lower() == "regular"): 
                sht.Range(cellAddress).Font.FontStyle = "Regular" 
        sht.Range(cellAddress).Font.Name = fontName

    def getrange(self, sheet, rangeAddress): 
        '''
            Returns a tuple of tuples from a range of cells. Each tuple 
            corresponds to a row in excel sheet. 
            sheet           -   name of the excel worksheet 
            rangeAddress    -   tuple of integers (row1,col1,row2,col2) or "cell1Address:cell2Address" 
                                row1,col1 refers to first cell 
                                row2,col2 refers to second cell 
                                e.g. (1,2,5,7) or "B1:G5" 
        '''
        
        sht = self.xlbook.Worksheets(sheet)
        if (isinstance(rangeAddress,tuple)):
            rangeAddress =self.convert_tuple_to_cell_str(rangeAddress)
        
        return sht.Range(rangeAddress).Value 

    #what is specify whole address
    def setrange(self, sheet, topRow, leftCol, data): 
        '''
            Sets range of cells with values from data. data is a tuple 
            of tuples. 
            Each tuple corresponds to a row in excel sheet. 

            sheet   -   name of the excel worksheet 
            topRow  -   row number (integer data type) 
            leftCol -   column number (integer data type) 
        ''' 
        bottomRow = topRow + len(data) - 1 
        rightCol = leftCol + len(data[0]) - 1 
        sht = self.xlbook.Worksheets(sheet) 
        sht.Range(sht.Cells(topRow, leftCol), sht.Cells(bottomRow, rightCol)).Value = data 
        return (bottomRow, rightCol) 

    def clearrange(self, sheet, rangeAddress, clear_contents= 1, clear_format= 1): 
        '''
            Clear the contents of a range of cells. 

            sheet           -   name of the excel worksheet 
            rangeAddress    -   tuple of integers (row1,col1,row2,col2) or "cell1Address:cell2Address" 
                                    row1,col1 refers to first cell 
                                    row2,col2 refers to second cell 
                                    e.g. (1,2,5,7) or "B1:G5" 
            clear_contents  -   1 or 0. If 1 clears the formulas from the range 
            clear_format    -   1 or 0. If 1 clears the formatting of the object 
        '''
        
        sht = self.xlbook.Worksheets(sheet)
        if (isinstance(rangeAddress,tuple)):
            rangeAddress =self.convert_tuple_to_cell_str(rangeAddress)
        
        if clear_contents: 
            sht.Range(rangeAddress).ClearFormats() 
        if clear_format: 
            sht.Range(rangeAddress).ClearContents() 



    def copy_paste_range(self,source, destination): 
        '''
            Copy range of data from source range in a sheet to 
            destination range in same sheet or different sheet.
            Will copy all format


            Description of parameters (self explanatory parameters are not 
            described): 
            source      -   tuple (sheet, rangeAddress) 
                            sheet           -   name of the excel sheet 
                            rangeAddress    -   "cell1Address:cell2Address" 
            destination -   tuple (sheet, destinationCellAddress) 
                            destinationCellAddress - string "ColumnRow"
                            (NB: destination can be bigger than the source
        ''' 
        sourceSht = self.xlbook.Worksheets(source[0]) 
        destinationSht = self.xlbook.Worksheets(destination[0]) 
        sourceSht.Range(source[1]).Copy(destinationSht.Range(destination[1])) 

    def addnewworksheet(self, oldSheet, newSheetName, after =1): 
        '''
            Adds a new excel sheet before/after the given excel sheet.
          
            if after =1 : add sheet after, if after =0, add sheet before.
            
            Description of parameters (self explanatory parameters are not 
            described): 
            oldSheet        -   Name of the sheet before which a new sheet should 
            be inserted 
            newSheetName    -   Name of the new sheet
            
        ''' 
        sht = self.xlbook.Worksheets(oldSheet) 
        if not after:
            self.xlbook.Worksheets.Add(sht).Name = newSheetName
        else:
            self.xlbook.Worksheets.Add(None,sht).Name = newSheetName 
        
        
    def insertchart(self, sheet, left, top, width =240 , height =160): 
        '''
            Creates a new embedded chart. Returns a ChartObject object. 
            Refer Add Method(ChartObjects Collection) in Microsoft Excel Visual 
            Basic Reference. 


            Description of parameters (self explanatory parameters are not 
            described): 
            sheet           -   name of the excel worksheet 
            left, top       -   The initial coordinates of the new object (in 
            points), relative to the upper-left corner of cell A1 on a worksheet 
            or to the upper-left corner of a chart. 
            width, height   -   The initial size of the new object, in points. 
                                point   -   A unit of measurement equal to 1/72 inch. 
        '''             
        sht = self.xlbook.Worksheets(sheet) 
        return sht.ChartObjects().Add(left, top, width, height) 


    def plotdata(self,  sheet, dataRanges, chartObject, chartType, plotBy=None,categoryLabels=1, seriesLabels=0, hasLegend=None, title=None, categoryTitle=None, valueTitle=None, extraTitle=None): 
        """Plots data using ChartWizard. For details refer ChartWizard 
            method in Microsoft Excel Visual Basic Reference. 
            Before using PlotData method InsertChart method should be used. 


            Description of parameters:         
            sheet       -   name of the excel worksheet. This name should be same 
            as that in InsertChart method 
            dataRanges  -   tuple of tuples ((topRow, leftCol, bottomRow, 
            rightCol),). Range of data in excel worksheet to be plotted. 
            chartObject -   Embedded chart object returned by InsertChart method. 
            chartType   -   Refer plotType variable for available options. 
            For remaining parameters refer ChartWizard method in Microsoft Excel 
            Visual Basic Reference. 
        """ 
        sht = self.xlbook.Worksheets(sheet) 
        if (len(dataRanges) == 1): 
            topRow, leftCol, bottomRow, rightCol = dataRanges[0] 
            source = sht.Range(sht.Cells(topRow, leftCol), sht.Cells(bottomRow, rightCol)) 
        elif (len(dataRanges) > 1):     
            topRow, leftCol, bottomRow, rightCol = dataRanges[0] 
            source = sht.Range(sht.Cells(topRow, leftCol), sht.Cells(bottomRow, rightCol)) 
            for count in range(len(dataRanges[1:])):             
                topRow, leftCol, bottomRow, rightCol = dataRanges[count+1] 
                tempSource = sht.Range(sht.Cells(topRow, leftCol), sht.Cells(bottomRow, rightCol)) 
                source = self.xlapp.Union(source, tempSource) 
        plotType = { 
                            "Area" : 1, 
                            "Bar" : 2, 
                            "Column" : 3, 
                            "Line" : 4, 
                            "Pie" : 5, 
                            "Radar" : -4151, 
                            "Scatter" : -4169,
                            "XYScatter": 72,#Smooth
                            "XYScatterLines": 74,
                            "Combination" : -4111, 
                            "3DArea" : -4098, 
                            "3DBar" : -4099, 
                            "3DColumn" : -4100, 
                            "3DPie" : -4101, 
                            "3DSurface" : -4103, 
                            "Doughnut" : -4120, 
                            "Radar" : -4151, 
                            "Bubble" : 15, 
                            "Surface" : 83, 
                            "Cone" : 3, 
                            "3DAreaStacked" : 78, 
                            "3DColumnStacked" : 55                          
                            } 
        gallery = plotType[chartType] 
        format = None 
        chartObject.Chart.ChartWizard(source, gallery, format, plotBy,categoryLabels, seriesLabels, hasLegend, title, categoryTitle, valueTitle, extraTitle) 





    def copychart(self, sourceChartObject, destination,delete="N"): 
        """Copy chart from source range in a sheet to destination 
            range in same sheet or different sheet 


            Description of parameters (self explanatory parameters are not 
            described): 
            sourceChartObject   -   Chart object returned by InsertChart method. 
            destination         -   tuple (sheet, destinationCellAddress) 
                                    sheet                  - name of the excel 
            worksheet. 
                                    destinationCellAddress - string "ColumnRow" 
                                    if sheet is omitted and only 
            destinationCellAddress is available as string data then same sheet is 
            assumed. 
            delete              -   "Y" or "N". If "Y" the source chart object is 
            deleted after copy. 
                                    So if "Y" copy chart is equivalent to move 
            chart. 
    """         
        if (isinstance(destination,tuple)): 
            sourceChartObject.Copy()             
            sht = self.xlbook.Worksheets(destination[0]) 
            sht.Paste(sht.Range(destination[1])) 
        else: 
            sourceChartObject.Chart.ChartArea.Copy() 
            destination.Chart.Paste() 
        if (delete.upper() =="Y"):             
            sourceChartObject.Delete() 


    def hidecolumn(self, sheet, col): #can combine hide column and hide row, multiple column??
        """Hide a column. 

            Description of parameters (self explanatory parameters are not 
            described): 
            sheet   -   name of the excel worksheet. 
            col     -   column number (integer data) 
        """         
        sht = self.xlbook.Worksheets(sheet) 
        sht.Columns(col).Hidden = True 


    def hiderow(self, sheet, row): #mulitple row??
        """ Hide a row. 


        Description of parameters (self explanatory parameters are not 
        described): 
        sheet   -   name of the excel worksheet. 
        row     -   row number (integer data) 
        """ 
        sht = self.xlbook.Worksheets(sheet) 
        sht.Rows(row).Hidden = True 


    def excelfunction(self, sheet, range, function): 
        """Access Microsoft Excel worksheet functions. Refer 
            WorksheetFunction Object in Microsoft Excel Visual Basic Reference 


            Description of parameters (self explanatory parameters are not 
            described): 
            sheet   -   name of the excel worksheet 
            range   -   tuple of integers (row1,col1,row2,col2) or 
            "cell1Address:cell2Address" 
                                row1,col1 refers to first cell 
                                row2,col2 refers to second cell 
                                e.g. (1,2,5,7) or "B1:G5" 
                        For list of functions refer List of Worksheet Functions 
            Available to Visual Basic in Microsoft Excel Visual Basic Reference 
        """         
        sht = self.xlbook.Worksheets(sheet) 
        if isinstance(range,str): 
            xlRange = "(sht.Range(" + "'" + range + "'" + "))" 
        elif isinstance(range,tuple): 
            topRow = range[0] 
            leftColumn = range[1] 
            bottomRow = range[2] 
            rightColumn = range[3] 
            xlRange = "(sht.Range(sht.Cells(topRow, leftColumn), sht.Cells(bottomRow, rightColumn)))" 
        xlFunction = "self.xlapp.WorksheetFunction." + function + xlRange 
        return eval(xlFunction, globals(), locals()) 

    def addcomment(self, sheet, cellAddress, comment=""): 
        """Add or delete comment to a cell. If parameter comment is 
            None, delete the comments 


            Description of parameters (self explanatory parameters are not 
            described): 
            sheet       -   name of the excel worksheet 
            cellAddress -   tuple of integers (row, cloumn) or string "ColumnRow" 
                            e.g. (3,4) or "D3" 
            comment     -   String data. Comment to be added. If None, delete 
            comments 
        """ 
        sht = self.xlbook.Worksheets(sheet) 
        if (isinstance(cellAddress,str)): 
            if (comment != None): 
                sht.Range(cellAddress).AddComment(comment) 
            else: 
                sht.Range(cellAddress).ClearComments()     
        elif (isinstance(cellAddress,tuple)):           
            row1 = cellAddress[0] 
            col1 = cellAddress[1] 
            if (comment != None): 
                sht.Cells(row1, col1).AddComment(comment) 
            else: 
                sht.Cells(row1, col1).ClearComments()


    def show_warning(self,Toggle = True):
        '''
            Whether to display alerts. Toggle = True means display
        '''
        self.xlapp.DisplayAlerts = Toggle

    def delete_sheet(self,selected_sheet):
        '''
            Delete sheet. No warning given.
            
        '''        
        self.show_warning(Toggle =False)
        sht = self.xlbook.Worksheets(selected_sheet)
        sht.Activate()
        self.xlapp.ActiveSheet.Delete()
        
    def setrangealign(self, sheet, rangeAddress, alignment= 'center'): 
        """Aligns the contents of a range. 
            Description of parameters (self explanatory parameters are not 
            described): 
            sheet       -   name of the excel worksheet 
            rangeAddress    -   tuple of integers (row1,col1,row2,col2) or 
            "cell1Address:cell2Address" 
                                row1,col1 refers to first cell 
                                row2,col2 refers to second cell 
                                e.g. (1,2,5,7) or "B1:G5"       
            alignment   -   "Left", "Right" or "center" 
        """ 
        alignment_dict = {'left':2, 'center': 3, 'centre': 3, 'right': 4 }  
        sht = self.xlbook.Worksheets(sheet)
        if (isinstance(rangeAddress,tuple)):
            rangeAddress =self.convert_tuple_to_cell_str(rangeAddress)
           
        sht.Range(rangeAddress).HorizontalAlignment = alignment_dict[alignment.lower()]


    #set font size alignment, dec place,can cater for single cell?
    def editrange(self, sheet, rangeAddress,dec_place ='0.00', font_size = 8,BOLD =False): 
        """
            sheet           -   name of the excel worksheet 
            rangeAddress    -   tuple of integers (row1,col1,row2,col2) or 
            "cell1Address:cell2Address" 
                                row1,col1 refers to first cell 
                                row2,col2 refers to second cell 
                                e.g. (1,2,5,7) or "B1:G5"
                                
            NB: editrange bold function may be affected by setrange function                                 
        """

        sht = self.xlbook.Worksheets(sheet)
        if (isinstance(rangeAddress,tuple)):
            rangeAddress =self.convert_tuple_to_cell_str(rangeAddress)

        #sht.Range(rangeAddress).Font.Bold = BOLD
        sht.Range(rangeAddress).Font.Size = font_size
        sht.Range(rangeAddress).NumberFormat= dec_place
        sht.Range(rangeAddress).Font.Bold = BOLD
            
    def sheets_name(self):
        """ Returns a list of the name of the sheets found in the document.
        """
        No_of_workbook = self.xlbook.Worksheets.Count
        Name_sheet =list()
        for n in range(1,int(No_of_workbook)+1,1):
            Name_sheet.append(self.xlbook.Worksheets.Item(n).Name)

        return Name_sheet

    def paste_special(self, sheet, rangeAddress,dest_sheet, topRow, leftCol,ValuesOnly = 0):
        """
        Copy and paste in one function
        Paste value from one sheet to another if ValuesOnly =1
        """
        if not ValuesOnly:
            
            temp = self.getrange(sheet, rangeAddress)
            self.setrange(dest_sheet,topRow,leftCol,temp)
        else:
            '''
                Paste special values only
            '''
            #print type(rangeAddress)
            self.activate_sel(sheet, rangeAddress)#activate selection--> problem with this
            #self.xlbook.Worksheets(sheet).Activate()
            self.copy_sel()
            self.paste_select(dest_sheet,(topRow, leftCol))#dest cell in tuple (row, column), by default here is paste values only command

    def paste_format(self, sheet, ref_range, dst_sheet, dst_range):
        '''
            Use paste special function with format only .
            Duplicate use of format painter
            Underconstruction
            within sheets
        '''
        self.activate_sel(sheet, ref_range)
        self.copy_sel()
        self.activate_sel(dst_sheet, dst_range)
        self.xlapp.Selection.PasteSpecial(Paste = -4122)#paste format or combined with paste select??
        #required deactivate??

    def delete_row(self,sheet, topRow,bottomRow):# combine delele or hide
        '''
            Delete entire row
        '''       
        sht = self.xlbook.Worksheets(sheet) 
        sht.Range(sht.Cells(topRow, 1), sht.Cells(bottomRow, 1)).EntireRow.Delete()

    def delete_column(self,sheet, topCol, bottomCol):
        '''
        Delete Entire Column
        '''
        sht = self.xlbook.Worksheets(sheet) 
        sht.Range(sht.Cells(1, topCol), sht.Cells(1, bottomCol)).EntireColumn.Delete()
        
    #highllight cell
    def highlight(self,sheet,rangeAddress, colorindex = 36):
        '''
            Highlight selected cell
        '''
        sht = self.xlbook.Worksheets(sheet)
        if (isinstance(rangeAddress,tuple)):
            rangeAddress =self.convert_tuple_to_cell_str(rangeAddress)
            
        sht.Range(rangeAddress).Interior.ColorIndex = colorindex

    #Wrap Text
    def wrap_text(self, sheet, rangeAddress, value = 'True'):
        sht = self.xlbook.Worksheets(sheet)
        if (isinstance(rangeAddress,tuple)):
            rangeAddress =self.convert_tuple_to_cell_str(rangeAddress)
        
        sht.Range(rangeAddress).WrapText = value


    def cells_in_percent(self, sheet, rangeAddress):
        sht = self.xlbook.Worksheets(sheet)
        if (isinstance(rangeAddress,tuple)):
            rangeAddress =self.convert_tuple_to_cell_str(rangeAddress)
        
        sht.Range(rangeAddress).Style = "Percent"


    def merge_cells(self, sheet, rangeAddress, value = 'True'):
        '''Merging Cells'''
        sht = self.xlbook.Worksheets(sheet)
        if (isinstance(rangeAddress,tuple)):
            rangeAddress =self.convert_tuple_to_cell_str(rangeAddress)
        
        sht.Range(rangeAddress).MergeCells = value

    #clear all formating in a cell
    def clearformat(self, sheet, rangeAddress):
        sht = self.xlbook.Worksheets(sheet) 
        if (isinstance(rangeAddress,str)): 
            sht.Range(rangeAddress).ClearFormats()
        elif (isinstance(rangeAddress,tuple)): 
            row1 = rangeAddress[0] 
            col1 = rangeAddress[1] 
            row2 = rangeAddress[2] 
            col2 = rangeAddress[3]             
            sht.Range(sht.Cells(row1, col1), sht.Cells(row2, col2)).ClearFormats()

    #have to check sheet is active, 
    def freezePanes(self, sheetname, freeze =True):
        #sht = self.xlbook.Worksheets(sheetname)
        self.xlbook.Worksheets(sheetname).Activate()
        self.xlapp.ActiveWindow.FreezePanes = freeze

    #insert row, if range is null, insert according to Selected rows, else insert after row stated --. need to clarify the parameters used
    def insert_row(self, sheetname = '', Row ='', Selection =True):
        if Selection:#if selection true, insert to selected row
            self.xlapp.Selection.Insert(Shift = 1)
        else:
            temp_string = str(Row) + ':' + str(Row)
            self.xlbook.Worksheets(sheetname).Range(temp_string).Insert(Shift = 1)

    #insert column
    def insert_column(self, sheetname = '', column ='', Selection =True):
        if Selection:#if selection true, insert to selected column
            self.xlapp.Selection.Insert(Shift = 1)
        else:
            if type(column) is int:
                column = self.convert_alphabet(column)
            temp_string = str(column) + ':' + str(column)
            self.xlbook.Worksheets(sheetname).Range(temp_string).Insert(Shift = 1)
            
            
    #insert and copy rows
    def copy_rows(self,sheetname, source_row, dest_row):
        self.xlbook.Worksheets(sheetname).Activate()#focus on the current sheet
        self.xlapp.Rows(str(source_row) + ':' + str(source_row)).Select()
        self.xlapp.Selection.Copy()
        self.xlapp.Rows(str(dest_row) + ':' + str(dest_row)).Select()
        self.xlapp.Selection.Insert( Shift=0)#insert copy cell



    #Set borders --all around
    def set_border(self, sheet, rangeAddress):
        sht = self.xlbook.Worksheets(sheet) 
        if (isinstance(rangeAddress,str)): 
            Selection = sht.Range(rangeAddress)
            Selection.Borders(5).LineStyle = -4142 #xlDiagonalDown = 5,xlNone = -4142 
            Selection.Borders(6).LineStyle = -4142 #xlDiagonalup = 5,xlNone = -4142
            for n in range(7,13,1):#for n ranging from 7 to 12
                Selection.Borders(n).LineStyle = 1 #xlcontinuous
                Selection.Borders(n).Weight = 2 # xlThin
                Selection.Borders(n).ColorIndex =  -4105 #xlAutomatic

            
        elif (isinstance(rangeAddress,tuple)): 
            row1 = rangeAddress[0] 
            col1 = rangeAddress[1] 
            row2 = rangeAddress[2] 
            col2 = rangeAddress[3]             
            Selection = sht.Range(sht.Cells(row1, col1), sht.Cells(row2, col2))
            Selection.Borders(5).LineStyle = -4142 #xlDiagonalDown = 5,xlNone = -4142 
            Selection.Borders(6).LineStyle = -4142 #xlDiagonalup = 5,xlNone = -4142
            for n in range(7,13,1):#for n ranging from 7 to 12
                Selection.Borders(n).LineStyle = 1 #xlcontinuous
                Selection.Borders(n).Weight = 2 # xlThin
                Selection.Borders(n).ColorIndex =  -4105 #xlAutomatic

    #multiple selection???
    def get_active_select_add(self):
        '''
            Return active selection address in fomat:
            (Address str, number coordinate (row ,column),sheetname)    

            Currently not supported for multiple selection
        '''
        
        def eval_cell_address(cell_string):
            return (cell_string.replace('$',''),(int(cell_string[1:].split('$')[1]),self.xlapp.Selection.Column),self.xlapp.Selection.Worksheet.Name.encode("ascii"))
        
        temp_add = self.xlapp.Selection.Address.encode('ascii')
        if temp_add.find(':') == -1:#only cell address
            return eval_cell_address(temp_add)
        else:
            #for a range (no mulitple selection yet)
            raw = self.xlapp.Selection.Address.encode("ascii")
            output_string = ['',list(),self.xlapp.Selection.Worksheet.Name.encode("ascii")]
            for n in temp_add.split(':'):
                self.xlapp.Range(n).Select()
                temp = eval_cell_address(n)
                output_string[0] = output_string[0] + ':'+ temp[0]
                output_string[1].append(temp[1][0])
                output_string[1].append(temp[1][1])
            
            self.xlapp.Range(raw).Select()
            output_string[0] = output_string[0][1:]#remove the first colon
            output_string[1] = tuple(output_string[1])#convert from list to tuple
            output_string = tuple(output_string)
            print output_string
            return output_string
                

    #creating a selected range
    def activate_sel(self,sheet, rangeAddress):
        sht = self.xlbook.Worksheets(sheet)
        sht.Activate()
        if (isinstance(rangeAddress,tuple)):
            rangeAddress =self.convert_tuple_to_cell_str(rangeAddress)

        sht.Range(rangeAddress).Activate()



    def get_values_from_all_selections(self):
        '''
            Function complement to the getRange function. Use where the selection is not continuous.
            Format will be different when selection is column-wise vs row  
            
        '''
        sheetname_list=self.xlapp.Selection.Worksheet.Name
        address_list = self.xlapp.Selection.Address.encode().split(',')
        data= list()
        for n in address_list:
            data.extend(self.xlbook.Worksheets(sheetname_list).Range(n).Value)

        return data            

    #make sure selection is already there
    def copy_sel(self):
        '''
            Item must be selected first before copied.
            Pls refer to activate_sel function to select a group of values
        '''
        self.xlapp.Selection.Copy()


    #paste selected --> make sure have already selected a row and is preparing to copy--> or can just include in
    #dest cell in tuple (row, column)
    def paste_select(self, sheet,dest_cell,values_only = 1):

        '''
            Item must be copied before this function can be used. Copy function can refer to copy_sel.
            if values_only = 1 selected, only values (no formatting) will be pasted
        
        '''
        
        sht = self.xlbook.Worksheets(sheet)
        sht.Activate()
        sht.Cells(dest_cell[0],dest_cell[1]).Activate()#to modify
        if values_only:#paste special
            self.xlapp.Selection.PasteSpecial(Paste = -4163 )#-4163
        else:
            self.xlapp.Selection.PasteSpecial()#-4163

    #underconstruction
    def find_keyword(self,sheetname,item):
        '''
            Return address of cell in tuple of (Address str, number coordinate (row ,column),sheetname)

            Item must in double quote -->  "????"            
                       
        '''

        #? change active cell to other
        self.activate_sel(sheetname,'A1:A1')
        Search = self.xlbook.Worksheets(sheetname).Cells.Find(What = item,MatchCase = True)
        activecell = Search.Activate()
        return self.get_active_select_add()

    def replace_cell_contents(self,shtname, target,replacement):
        self.xlbook.Worksheets(shtname).Cells.Replace(What = target,Replacement = replacement,
                                                      MatchCase = True, LookAt = 1)#xlwhole =1


    def import_ext_data(self,filename,sheet,Dest_cell ="A1"):
        '''
            Import data from a text file.
            Underconstruction: make sure there is no dependancy
                             : Check for file present
        '''
        sht = self.xlbook.Worksheets(sheet)
        sht.Activate()
        temp = self.xlapp.ActiveSheet.QueryTables
        filename = 'TEXT;' + filename
        temp_ext_data = temp.Add(Connection = filename, Destination= self.xlapp.Range(Dest_cell))
        temp_ext_data.Refresh()
        temp_ext_data.Delete()

    def copy_data_fr_csv(self, filename, sheet_to_paste, do_transpose = 0):
        '''
            Assume copy all text
            assume sheet is already in place
            underconstruct: 
            
        '''
        #check if the sheetname is present, if not, add the sheet

        if sheet_to_paste not in self.sheets_name():
            self.addnewworksheet(self.sheets_name()[0],sheet_to_paste, after=0 )

        try:
            src_xlbook = self.xlapp.Workbooks.Open(filename)
            self.xlapp.Application.DisplayAlerts =False
            src_xlbook.ActiveSheet.Cells.Copy()
            self.activate_sel(sheet_to_paste, "A1:A1")
            #self.xlbook.ActiveSheet.Paste() #can use paste special here
            if do_transpose:
                self.xlapp.Selection.PasteSpecial(Paste = -4163 , Transpose = True)#value -4163#format is -4122
            else:
                self.xlapp.Selection.PasteSpecial(Paste = -4163)
            print 'Finish exporting'
        finally:
            src_xlbook.Close()
            self.xlapp.Application.DisplayAlerts =True

    def paste_format_fr_ext_xls(self, src_filename, src_sheet, target_sheet,target_range = "A1:A1"):
        '''
            Able to format target range, if format full area, use target_range = "A1:A1"
            Present selecting range does not work correctly
        '''
        self.xlapp.Application.DisplayAlerts =False
        try:
            src_xlbook = self.xlapp.Workbooks.Open(src_filename)
            src_sht = src_xlbook.Worksheets(src_sheet)
            src_sht.Activate()
            src_sht.Range(target_range).Activate()
            src_xlbook.ActiveSheet.Cells.Copy()
            
            self.activate_sel(target_sheet, target_range)
            self.xlapp.Selection.PasteSpecial(Paste = -4122 )#value -4163#format is -4122

        finally:
            src_xlbook.Close()
            self.xlapp.Application.DisplayAlerts =True

    def autofit_column_width(self,sheet,address):
        '''
            auto fit column width

            Col Address in str (presently only applicable to one column) --> eg:'C:C'
        '''
        sht = self.xlbook.Worksheets(sheet) 
        sht.Columns(address).EntireColumn.AutoFit()

    #count number of sheet
    def count_sheet(self):
        return self.xlbook.Worksheets.Count

    def get_shtname_and_lastdata_row_col_str(self, return_as_tuple = 0):
        '''If return_as_tuple == 1, return tuple (sheetname, row, column) '''
        lastcell = self.xlbook.ActiveSheet.Cells.SpecialCells(11)#xllastcell
        active_sheet_name = self.xlbook.ActiveSheet.Name
        output_str = '%s!R1C1:R%sC%s' %(active_sheet_name,lastcell.Row,lastcell.Column)
        if not return_as_tuple:
            return output_str
        else:
            return (active_sheet_name,lastcell.Row,lastcell.Column)

    def pivot_chart_generation(self):
        data_sheet_and_lastcell_info = self.get_shtname_and_lastdata_row_col_str()
        newsheet = self.xlbook.Sheets.Add()
        pc = self.xlbook.PivotCaches().Add(SourceType=1,SourceData = data_sheet_and_lastcell_info)#win32c.xlDatabase
        pt = pc.CreatePivotTable(TableDestination="%s!R4C1"%newsheet.Name,
                             TableName= self.pivot_table_name,
                             DefaultVersion=1)#win32c.xlPivotTableVersion10
        #create the pivot table then thechart
        #may have to add the row column then the data column??
        for pos, n in enumerate(self.rowfield_list):
            
            self.xlbook.ActiveSheet.PivotTables(self.pivot_table_name).PivotFields(n).Orientation = 1#xlrow field
            self.xlbook.ActiveSheet.PivotTables(self.pivot_table_name).PivotFields(n).Position = pos+1

        for pos, n in enumerate(self.colfield_list):
            
            self.xlbook.ActiveSheet.PivotTables(self.pivot_table_name).PivotFields(n).Orientation = 2#xlcol field
            self.xlbook.ActiveSheet.PivotTables(self.pivot_table_name).PivotFields(n).Position = pos+1

        c = pt.AddDataField(pt.PivotFields(self.datafield),"Sum of %s" %self.datafield,-4157) #sunm

        self.xlbook.ActiveSheet.Cells(5,1).Select()
        ss = self.xlapp.Charts.Add()#change the chart type later
    
    ##################################### Below mainly on chart function #####################################
    #count number of chart in sheet
    def count_chart(self,sheetname):
        return self.xlbook.Worksheets(sheetname).ChartObjects().Count
    

    #delete a series in a chart
    def chart_series_del(self,sheetname,chartname,series_no):
        #count object
        self.xlbook.Worksheets(sheetname).ChartObjects().Count

        
        chart_obj = self.xlbook.Worksheets(sheetname).ChartObjects(chartname)
        #count number of series
        chart_obj.Chart.SeriesCollection().Count
        
        chart_obj.Chart.SeriesCollection(series_no).Delete()

    #for adding series
    def chart_series_add(self,sheetname,chartname,X_tuple,Y_tuple,series_name):#should put the name here, then if name empty 
        '''
        tuple in form of start row, start col, end row, end column
        The x_string and Y_string are to be in this format '=Compile!R9C8:R12C8','=Compile!R9C11:R12C11'
        x_string/ y-string in tuple --> use conversion
        present series name must be a tuple, should change to include string later
        '''
        #conversion
        if X_tuple ==():
            print 'no formating for x-axis'
            X_string =()
        else:
            X_string = format_ch_series_data(sheetname,X_tuple)
            
        Y_string = format_ch_series_data(sheetname,Y_tuple)
        try:
            if type(series_name) == tuple:
                series_legend = format_ch_series_data(sheetname,series_name)
            else:#if not equal none
                series_legend = series_name
        except:
            print'series legend is None'
            pass
        chart_obj = self.xlbook.Worksheets(sheetname).ChartObjects(chartname)
        w =chart_obj.Chart.SeriesCollection()
        w.NewSeries()#create a new series
        #check the total number of series and use the last one as added
        new_series =chart_obj.Chart.SeriesCollection().Count
        if X_tuple !=():
            chart_obj.Chart.SeriesCollection(new_series).XValues = X_string
        chart_obj.Chart.SeriesCollection(new_series).Values = Y_string
        chart_obj.Chart.SeriesCollection(new_series).Name = series_legend

    def chart_series_count(self,sheetname,chartname):
        '''Return the number of series in a chart'''
        #count object
        self.xlbook.Worksheets(sheetname).ChartObjects().Count
        
        chart_obj = self.xlbook.Worksheets(sheetname).ChartObjects(chartname)
        #count number of series
        return chart_obj.Chart.SeriesCollection().Count

    #change series chart type of particular chart
    def series_charttype(self,sheetname,chartname,series_no,chart_type):
        '''
        For different chart series within the same chart
        '''
        chart_obj = self.xlbook.Worksheets(sheetname).ChartObjects(chartname)
        chart_obj.Chart.SeriesCollection(series_no).ChartType = chart_type
    

    #changin legend background color
    def format_legend(self,sheetname,chartname,colourindex =15, font_size = 8):#15 is grey
        '''Colour and font size'''
        chart_obj = self.xlbook.Worksheets(sheetname).ChartObjects(chartname)
        chart_obj.Chart.Legend.Interior.ColorIndex =colourindex
        chart_obj.Chart.Legend.Font.Size=font_size

    #not sure if don turn on, still can change style??
    def format_gridlines(self,sheetname,chartname,X_axis_on = True,Y_axis_on = False,Faint =False ):
        '''X_axis/Y_axis on allow enabling major gridlines on X,Y axis respectively
            Faint toggle means changing the gridlines to either dotted or continuous line
            
        '''
        chart_obj = self.xlbook.Worksheets(sheetname).ChartObjects(chartname)       
        #X-axis first
        X_object = chart_obj.Chart.Axes(1)
        X_object.HasMajorGridlines = X_axis_on
        Y_object = chart_obj.Chart.Axes(2)
        Y_object.HasMajorGridlines = Y_axis_on       
        
        if Faint:
            #make the line faint
            X_object.MajorGridlines.Border.ColorIndex =16
            X_object.MajorGridlines.Border.LineStyle = -4118
            Y_object.MajorGridlines.Border.ColorIndex =16
            Y_object.MajorGridlines.Border.LineStyle = -4118

            
    #under construction
    def scale_change(self,sheetname,chartname,Axis, max, min, interval, cutoff,font_size = 8, dec_pl = '0.00'):
        '''
        Scale change --> enter 1 for x axis and 2 for y axis
        Include number format, dec place in string eg '0.00'
        '''
        chart_obj = self.xlbook.Worksheets(sheetname).ChartObjects(chartname)
        Scale_obj = chart_obj.Chart.Axes(Axis)
        Scale_obj.MaximumScale =max
        Scale_obj.MinimumScale =min
        Scale_obj.MajorUnit = interval
        Scale_obj.CrossesAt = cutoff
        Scale_obj.TickLabels.Font.Size = font_size

        #change the axis number format
        chart_obj.Chart.Axes(Axis).TickLabels.NumberFormat = dec_pl


    def label_chart(self,sheetname,chartname,title,x_label, y_label,title_font = 10,x_font = 8,y_font = 8):
        '''
        use '' if do not want to include label
        '''
        chart_obj = self.xlbook.Worksheets(sheetname).ChartObjects(chartname)
        chart_obj.Chart.HasTitle =True
        chart_obj.Chart.ChartTitle.Characters.Text= title
        chart_obj.Chart.ChartTitle.Font.Size =title_font
        if not x_label == '':
            chart_obj.Chart.Axes(1).HasTitle =True
            chart_obj.Chart.Axes(1).AxisTitle.Characters.Text = x_label
            chart_obj.Chart.Axes(1).AxisTitle.Font.Size= x_font
            
        if not y_label == '':
            chart_obj.Chart.Axes(2).HasTitle =True
            chart_obj.Chart.Axes(2).AxisTitle.Characters.Text = y_label
            chart_obj.Chart.Axes(2).AxisTitle.Font.Size= y_font

    #prevent auto font resizing when scaling graph
    def freeze_scalefont(self,sheetname,chartname,AutoScale =False):
        chart_obj = self.xlbook.Worksheets(sheetname).ChartObjects(chartname)
        chart_obj.Chart.ChartArea.AutoScaleFont = AutoScale 

    #Under Construction        
    def add_trendline(self,sheetname,chartname,series_no):
        #need to check whether the data initially have trendline?? --> else remove all trendline then add'''
        chart_obj = self.xlbook.Worksheets(sheetname).ChartObjects(chartname)
        series_obj = chart_obj.Chart.SeriesCollection(series_no)
        #add trendine
        trendline_obj = series_obj.Trendlines().Add()#add with forecast
        #because automatic always produce value of 57 --> that is why difficult to add colour
##        >>> c.Border.ColorIndex = 50
##>>> c.Border.ColorIndex = 57
##>>> a.MarkerForegroundColor


    #more chart function
    def chart_formatter_fr_ext_xls(self, src_filename, src_sheet, src_chartname, target_sheet, target_chartname):
        '''
            Copy chart format from one chart in an excel file  to another
            Able to copy mulitple charts at one go

        '''
        
        self.xlapp.Application.DisplayAlerts =False
        try:
            src_xlbook = self.xlapp.Workbooks.Open(src_filename)
            src_sht = src_xlbook.Worksheets(src_sheet)
            src_sht.Activate()
            no_of_charts = src_xlbook.Worksheets(src_sheet).ChartObjects().Count
            print 'src xlbook charts count:', no_of_charts
            print 'src xlbook charts namelist', [src_xlbook.ActiveSheet.ChartObjects(n).Name for n in range(1,no_of_charts+1,1)]

            chart_obj = src_xlbook.Worksheets(src_sheet).ChartObjects(src_chartname)    
            chart_obj.Chart.ChartArea.Copy()

            trc_sht = self.xlbook.Worksheets(target_sheet)
            trc_sht.Activate()
            tar_chart_namelist = [self.xlbook.ActiveSheet.ChartObjects(n).Name for n in range(1,no_of_charts+1,1)]
            print 'tar xlbook charts namelist', tar_chart_namelist
            
            tar_chart_obj = self.xlbook.ActiveSheet.ChartObjects(target_chartname)
            tar_chart_obj.Activate()
            tar_chart_obj.Chart.ChartArea.Select()
            self.xlbook.ActiveChart.Paste(Type = -4122)#format only

        finally:
            src_xlbook.Close()
            self.xlapp.Application.DisplayAlerts =True
# ######################### More Functions  ##############################################################

#convert the number of squares to points(left,top)-->left means how far from x -axis (each cell ard 48pts), y means y axis (each cell ht is 15pts)
#mainly for chart purpose
def convert_sq_to_pts(left_num_cell, top_num_cell):

    return left_num_cell*48, top_num_cell*15

#under construction --> making len equal for every row to be used in set range
def equal_rows(data1):
    len_data = list()
    for n in range(len(data1)):
        len_data.append(len(data1[n]))

    #make every rows to be the max len
    for n in range(len(data1)):
        if len(data1[n]) < max(len_data):
            for n1 in range(max(len_data) - len(data1[n])):#add space till the len is same
                data1[n].append('')
    return data1


def format_ch_series_data(sheet,tuple_coord):
    '''Return the string format of series =Compile!R9C8:R12C8
        tuple in form of start row, start col, end row, end column
    '''
    return '='+ '\'' + str(sheet)+ '\'' + '!' + 'R' + str(tuple_coord[0]) + 'C' + str(tuple_coord[1]) \
           + ':' + 'R' + str(tuple_coord[2]) + 'C' + str(tuple_coord[3])


#converting column integer to alphabet --> note first row start from 1
def Convert_alphabet(col_num):
    if col_num > 26:
        return chr(int((col_num - 1) / 26) + 64) + chr(((col_num - 1) % 26) + 65)
    else:
        return chr(col_num + 64)

#copy sheet from different workbook or from same workbook    
def copysheet(select_workbook,select_worksheet,dest_workbook, dest_worksheet, sheetname):
    '''make sure create two workbook using open function if need to copy betw two workbook
        use xlfile.xlbook'''
    sht = select_workbook.Worksheets(select_worksheet)
    sht1 = dest_workbook.Worksheets(dest_worksheet)
    sht.Copy(sht1)
    copied_sheet = dest_workbook.ActiveSheet
    if  not sheetname  == '':            
        copied_sheet.Name=sheetname
    return copied_sheet.Name


def chart_type(chartObject,type = 72):
    '''
       Modified chart type parameters.

       Chart scatter = 72, scatter without line = '-4169'       
    '''
    chartObject.Chart.ChartType = str(type)


def search_sheet(xlFile, sheetname):

    sheets_data = xlFile.sheets_name()
    for n in sheets_data:
        if sheetname == n:
            return True
    return False
    
def add_sheet(xlFile, sheetname,previous_sheet):
    #print 'Searching and adding sheet'
    counter =1
    while(True):
        if search_sheet(xlFile, sheetname):
            if counter ==1: #if there is no appending
                sheetname = str(sheetname) + '_'+ str(counter)
                counter = counter +1
            else:
                #remove the previous counter
                reverse_name = sheetname[::-1]#reverse name to get the index out
                reverse_index = reverse_name.index('_')
                sheetname = reverse_name[reverse_index+1:][::-1]#reverse back and add the counter              
                sheetname =sheetname + '_'+ str(counter)
                counter = counter + 1
        else:
            break
    xlFile.addnewworksheet(previous_sheet, sheetname)
    return sheetname

def search_filename(filename,ldir = [ 'C:\\data\\' ]):
    #print 'Searching file name.....'
    for d in ldir:
        for f in glob.glob(d+'\\*.xls'):
            if f == filename:
                return True
        else: return False

def openxlfile_GUI(message='select data file',file_title='Data file'):
    import pyET_tools.support_gui
    import sys
    filename = pyET_tools.support_gui.dialog.file_open_dialog(msg = message, title=file_title)
    if filename ==None:
        sys.exit(0)
    filename=convert_filename(filename)
    xlFile = UseExcel(filename)
    return xlFile

def convert_filename(filename):
    import re
    return re.sub(r'/',r'\\',filename,0)

def get_filename(xlFile):
    '''note the path is not given'''
    import re
    mobj = re.search('Microsoft Excel \- (.*)',xlFile.xlapp.Caption )
    return mobj.group(1)

def check_close_excel_wb(filename):
    win.close_a_running_win_program(filename)

def testing():
    #for testing purpose
    test = 4
    
    
    if test ==1:    
        xlapp = win32com.client.dynamic.Dispatch("Excel.Application")
        #xlbook = xlapp.Workbooks.Add()
        xlbook =xlapp.Workbooks.Open(r'c:\data\temp\cda_format_trial.xls') 
        xlapp.Visible=1
        
        import sys
        sys.exit()
        #try out pivot table
        pivot_table_name = ''
        rowfield_list = [] #in seq
        colfield_list = []
        datafield = ''
        source_data_str = ''#how to auto see number of row column and 
        

        newsheet = xlbook.Sheets.Add()
        pc = xlbook.PivotCaches().Add(SourceType=1,SourceData="P061_OW_MEASUREMENT!R1C1:R2513C37")#win32c.xlDatabase
        pt = pc.CreatePivotTable(TableDestination="%s!R4C1"%newsheet.Name,
                             TableName="ggg",
                             DefaultVersion=1)#win32c.xlPivotTableVersion10
        #create the pivot table then thechart
        #may have to add the row column then the data column??
        for pos, n in enumerate(rowfield_list):
            
            xlbook.ActiveSheet.PivotTables("ggg").PivotFields(n).Orientation = 1#xlrow field
            xlbook.ActiveSheet.PivotTables("ggg").PivotFields(n).Position = pos


        xlbook.ActiveSheet.PivotTables("ggg").PivotFields("HD_PHYS_PSN").Orientation = 1#xlrow field
        xlbook.ActiveSheet.PivotTables("ggg").PivotFields("HD_PHYS_PSN").Position = 2
        
        xlbook.ActiveSheet.PivotTables("ggg").PivotFields("DATA_ZONE").Orientation = 1#xlrow field
        xlbook.ActiveSheet.PivotTables("ggg").PivotFields("DATA_ZONE").Position = 3
        xlbook.ActiveSheet.PivotTables("ggg").PivotFields("TRK_NUM").Orientation = 1#xlrow field
        xlbook.ActiveSheet.PivotTables("ggg").PivotFields("TRK_NUM").Position = 4

        xlbook.ActiveSheet.PivotTables("ggg").PivotFields("SBR").Orientation = 2#xlcol field
        xlbook.ActiveSheet.PivotTables("ggg").PivotFields("SBR").Position = 1

        xlbook.ActiveSheet.PivotTables("ggg").PivotFields("SERIAL_NUM").Orientation = 2#xlcol field
        xlbook.ActiveSheet.PivotTables("ggg").PivotFields("SERIAL_NUM").Position = 2

        c = pt.AddDataField(pt.PivotFields("OW_MEASUREMENT"),"Sum of OW_MEASUREMENT",-4157) #sunm

        xlbook.ActiveSheet.Cells(5,1).Select()
        ss =xlapp.Charts.Add()

##        xlbook.ActiveSheet.Select()
##        target = xlbook.Charts.Add()
##        target.Location(1)
##        target.PivotLayout
        
    if test ==2:
        #filename = r'c:\data\temp\Yarra 2D Scpk request 40 prs HGAs with coherence data 03Dec2010.xls'
        #xlFile = UseExcel(filename)
        xlFile = UseExcel(r'c:\data\temp\P061_OW_MEASUREMENT_try_pivot1.xls')
        xlFile.show()
        
        
        xlFile.pivot_table_name = 'ggg'
        xlFile.rowfield_list = ['OPERATION','HD_PHYS_PSN','DATA_ZONE','TRK_NUM'] #in seq
        xlFile.colfield_list = ['SBR','SERIAL_NUM']
        xlFile.datafield = 'OW_MEASUREMENT'

        xlFile.pivot_chart_generation()
        #next need to create only one visible data

        #xlFile.highlight('sheet1', (1,1,5,5),35)
        #cells_in_percent(self, sheet, rangeAddress)
        #xlFile.paste_format(sheet, ref_range, dst_sheet, dst_range)
        #raw_input()
        #print xlFile.get_active_selection()#for range
##        print xlFile.get_selectadd(number_format =True)#only for two
##        print xlFile.get_selectadd(number_format =False)
##        print xlFile.active_address()# only for two
##        print xlFile.active_selection()#only for sheet
        
##        raw_input('Select')
##        data = xlFile.get_values_from_all_selections()

        #date function to settle        
    if test ==3:
        
        xlapp = win32com.client.dynamic.Dispatch("Excel.Application")
        #xlbook = xlapp.Workbooks.Add()
        xlbook =xlapp.Workbooks.Open(r'c:\data\temp\ans.csv') 
        xlapp.Visible=1
        xlapp.Application.DisplayAlerts =False
        xlbook1 = xlapp.Workbooks.Add()
        xlbook.ActiveSheet.Cells.Copy()
        xlbook.ActiveSheet.Range("A1:A1").Activate()
        xlbook.ActiveSheet.Paste()

        #"P061_OW_MEASUREMENT!R1C1:R2513C37"


    if test ==4:
        #filename = r'c:\data\temp\Yarra 2D Scpk request 40 prs HGAs with coherence data 03Dec2010.xls'
        #xlFile = UseExcel(filename)
        xlFile = UseExcel(r'C:\data\temp\cda_format_trial.xls')
        xlFile.show()
        xlFile.copy_data_fr_csv(r'c:\data\temp\ans.csv', 'Sheet1', 1)
        #xlFile.close()
        
    if test ==5:
        
        xlapp = win32com.client.dynamic.Dispatch("Excel.Application")
        #xlbook = xlapp.Workbooks.Add()
        xlbook =xlapp.Workbooks.Open(r'c:\data\temp\sample_target.xls') 
        xlapp.Visible=1
        xlapp.Application.DisplayAlerts =False
        #target also need a sheetname


        try:
            src_xlbook = xlapp.Workbooks.Open(r'c:\data\temp\cda_format_trial.xls')
            #need to activate the active sheet in src_xlbook
            src_sht = src_xlbook.Worksheets('ADC_SUM_QB1h')
            src_sht.Activate()
            src_sht.Range("A1:A1").Activate()
            #sheet name to use ADC_SUM_QB1h
            src_xlbook.ActiveSheet.Cells.Copy()
            trc_sht = xlbook.Worksheets('Sheet2')
            trc_sht.Activate()
            trc_sht.Range("A1:A1").Activate()
            xlapp.Selection.PasteSpecial(Paste = -4163 )#-4163#format is -4122
            #xlbook.ActiveSheet.Paste()

        finally:
            src_xlbook.Close()
            xlapp.Application.DisplayAlerts =True

    if test == 6:
        xlFile = UseExcel(r'c:\data\temp\sample_target.xls')
        xlFile.show()
        xlFile.paste_format_fr_ext_xls(r'c:\data\temp\cda_format_trial.xls','ADC_SUM_QB1h','Sheet1', target_range = "B2:C45")
        
        #enable paste certain range
        #may not work as paste has to be of certain range


    if test ==7:
        print 'to deal with replace value by full case'
        xlFile =  UseExcel(r'c:\data\temp\sample_target.xls')
        xlFile.show()
        xlFile.replace_cell_contents('Sheet1','65535','')

    if test ==8:
        xlapp = win32com.client.dynamic.Dispatch("Excel.Application")
        #xlbook = xlapp.Workbooks.Add()
        xlbook =xlapp.Workbooks.Open(r'c:\data\temp\sample_target.xls') 
        xlapp.Visible=1
        #xlbook.Worksheets('Sheet1').Cells.Replace(What = '65535',Replacement = '', MatchCase = True, LookAt = 1)#xlwhole =1
        lastcell = xlbook.ActiveSheet.Cells.SpecialCells(11)#xllastcell, xlfirst = 0
        print lastcell.Row,lastcell.Column
        #cannot get the first cell???

        #need to get the range of data --> detect where is the data range??? --> have the last cell how to get first cell
        #function to search and replace

    if test ==9:
        #use for developing chart formatter
        #function to save the charts as well
        
        xlapp = win32com.client.dynamic.Dispatch("Excel.Application")
        #xlbook = xlapp.Workbooks.Add()
        xlbook =xlapp.Workbooks.Open(r'c:\data\temp\cda_format_trial.xls') 
        xlapp.Visible=1
##        xlapp.Application.DisplayAlerts =False
        #target also need a sheetname

        src_xlbook = xlapp.Workbooks.Open(r'c:\data\temp\xls_fmt_template.xls')
        src_sht = src_xlbook.Worksheets('GEN_SUM_bySBR')
        src_sht.Activate()        
        no_of_charts = src_xlbook.Worksheets('GEN_SUM_bySBR').ChartObjects().Count

        #get all chart name
        chart_namelist = [src_xlbook.ActiveSheet.ChartObjects(n).Name for n in range(1,no_of_charts+1,1)]
        #for n in no_of_charts:
        chart_obj = src_xlbook.Worksheets('GEN_SUM_bySBR').ChartObjects(chart_namelist[0])    
        chart_obj.Chart.ChartArea.Copy()
        #need to activate the active sheet in src_xlbook


        trc_sht = xlbook.Worksheets('ADC_SUM_0Arv')
        trc_sht.Activate()
        tar_chart_namelist = [xlbook.ActiveSheet.ChartObjects(n).Name for n in range(1,no_of_charts+1,1)]
        tar_chart_obj = xlbook.ActiveSheet.ChartObjects(tar_chart_namelist[0])
        tar_chart_obj.Activate()
        tar_chart_obj.Chart.ChartArea.Select()
        xlbook.ActiveChart.Paste(Type = -4122)

    if test ==10:
        xlFile =  UseExcel(r'c:\data\temp\cda_format_trial.xls')
        xlFile.show()
        src_filename = r'c:\data\temp\xls_fmt_template.xls'
        src_sheet = 'GEN_SUM_bySBR'
        src_chartname = 'Chart 1'
        target_sheet= 'ADC_SUM_0Arv'
        target_chartname = 'Chart 1'
        xlFile.chart_formatter_fr_ext_xls(src_filename, src_sheet, src_chartname, target_sheet, target_chartname)


if (__name__ == "__main__"):
    testing()
    