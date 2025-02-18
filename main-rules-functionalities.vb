Public Module Program
	Public Sub Main(args() As string)
		Range("A1").Font.Bold = True 'Makes Font Bold'
		Rnge("A1").Interior.Color = RGB(255, 255, 0) 'Applies Background Color Yellow'
		Range("A1").Font.Color = RGB(255, 0, 0) 'Applies Font Color Red'
		Range("A1").Value = "Hello VBA" 'Inserts a string to cell A1'
		Cells(2,1).Value = 12345
		Range("B1").Formula - "=SUM(A1:A10)" 'Applies a Formula'
		Range("B2").Value = Range("A1").Value 'Copying a Value from Another Cell to Another'
		
		
		'Apply Conditional Formatting to Highlight Cells According to their Values'
		dataRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="90")
		dataRange.FormatConditions(1).Interior.Color = RGB(0, 255, 0)
		dataRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="80")
		dataRange.FormatConditions(2).Interior.Color = RGGB(0, 255, 0)
		
		'Removing Conditional Formatting'
		dataRange.FormatConditions.Delete
		
		'Adding a New Worksheet'
		Sheets.Add 'Adds a new worksheet before the active sheet'
		Sheets.Add After:=Worksheets(Worksheets.Count) 'Adds a worksheet at the end'
		
		'Renaming a Worksheet'
		ActiveSheet.Name = "Total Sales" 'Renames the ActiveSheet to Total Sales'
		
		'Delete Worksheet'
		Application.DisplayAlerts = False  'Disables The Confirmation Prompt'
		Worksheets("Sheet1").Delete 'Deletes The Specified Worksheet'
		Application.DisplayAlerts = True 'Re-enables Alerts'
		
		
		'Activating a Worksheet'
		Worksheets("SalesData").Activate 'Activates the Worksheet named SalesData'
		
		'Looping Trough All Worksheets'
		Dim ws As Worksheet
		For Each ws In Worksheets
		  ws.Activate
		  'Perform Actions on Worksheet'
		Next ws 
		
		
		'Variables Declarations Example'
		Dim Total As Integer 
		Total = 10 + 20 'Stores The Result Of the Calculation'
		
		'Declaring Variables'
		Dim VariableName As DataType 
		Dim Age As Integer
		Dim Name As String 
		Dim Salary As Double 
		Dim IsAdmin As Boolean 
		Dim HireDate As Date 
		
		'Initializing Variables'
		Age = 25 
		Name = "John"
		IsAdmin = True 
		HireDate = #1/1/2025#
		
		
		'Using Variables in A Macro'
		Dim Price As Double 
		Dim Quantity As Integer
		Dim Total As Double 
		
		Price = 15.99
		Quantity = 10 
		Total = Price * Quantity
		
		MsgBox "The Total is: " & Total
	End Sub
End Module
