Module VBModule
    Sub Main()
        'Workbook Object'
        'Open a Workbook'
        Workbooks.Open "C:\MyFolder\MyWorkbook.xlsx"
        
        'Save the activate workbook'
        ActiveWorkbook.Save 
        
        'Close a Workbook without saving changes'
        ActiveWorkbook.Close SaveChanges:=False 
        
        
        'Worksheet Object'
        'Activate a specific worksheet'
        Worksheets("Sheet1").Activate
        
        'Rename a worksheet'
        Worksheets("Sheet1").Name = "DataSheet"
        
        'Hide a worksheet'
        Worksheets("DataSheet").Visible = xlSheetHidden
        
        
        'Range Object'
        'Set the value of a single cell'
        Range("A1").Value = "Hello, VBA"
        
        'Format a range of cells'
        Range("A1:A5").Font.Bold = True
        Range("A1:A5").Interior.Color = RGB(255, 255, 0) 'Yellow background
        
        'Copy data from one range to another'
        Range("B1:B5").Value = Range("A1:A5").Value
    End Sub
End Module
