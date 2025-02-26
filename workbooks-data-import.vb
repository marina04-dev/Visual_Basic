Module VBModule
    Sub Main()
        '' Opening a Workbook
        Sub OpenWorkbook()
            Workbooks.Open "C:\MyFiles\DataWorkbook.xlsx"
        End Sub
        
        ''If the workbook is password-protected, you can include the password in the code:
        Sub OpenProtectedWorkbook()
            Workbooks.Open "C:\MyFiles\SecureWorkbook.xlsx", Password:="mypassword"
        End Sub
        
        ''Saving a Workbook
        Sub SaveActiveWorkbook()
            ActiveWorkbook.Save
        End Sub
        
        ''Save a workbook with a new name:
        Sub SaveAsNewFile()
            ActiveWorkbook.SaveAs "C:\MyFiles\NewWorkbook.xlsx"
        End Sub
        
        ''Closing a Workbook without saving 
        Sub CloseWithoutSaving()
            ActiveWorkbook.Close SaveChanges:=False
        End Sub
        
        ''Closing a Workbook and save changes 
        Sub CloseAndSave()
            ActiveWorkbook.Close SaveChanges:=True 
        End Sub
        
        ''Importing a CSV File
        Sub ImportCSV()
            Workbooks.OpenText Filename:="C:\MyFiles\Data.csv", DataType:=xlDelimited, Comma:=True
        End Sub
        
        ''Exporting a Range to CSV
        Sub ExportToCSV()
            ActiveWorkbook.SaveAs Filename:="C:\MyFiles\ExportedData.csv", FileFormat:=xlCSV
        End Sub
        
        ''Copying a Range Between Workbooks
        Sub CopyDataBetweenWorkbooks()
            Dim SourceWB As Workbook
            Dim TargetWB As Workbook

            Set SourceWB = Workbooks.Open("C:\MyFiles\Source.xlsx")
            Set TargetWB = Workbooks.Open("C:\MyFiles\Target.xlsx")

            SourceWB.Sheets(1).Range("A1:B10").Copy Destination:=TargetWB.Sheets(1).Range("A1")

            SourceWB.Close SaveChanges:=False
            TargetWB.Save
        End Sub
    End Sub
End Module
