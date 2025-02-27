
Module VBModule
    Sub Main()
        ''Writing VBA Code for Data Validation
        ''Adding Data Validation Rules
        Sub AddDataValidation()
            With Range("A1:A10").Validation
            .Delete 'Remove existing validation rules
            .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="1", Formula2:="100"
            .InputTitle = "Valid Numbers"
            .ErrorTitle = "Invalid Entry"
            .InputMessage = "Please enter a number between 1 and 100."
            .ErrorMessage = "Only numbers between 1 and 100 are allowed."
            End With
        End Sub
        
        ''Removing Data Validation Rules
        Sub RemoveDataValidation()
            Range("A1:A10").Validation.Delete
        End Sub
        
        ''Automating Data Input and Formatting
        ''Populating Cells with Data
        Sub PopulateData()
            Range("A1").Value = "Product Name"
            Range("A2").Value = "Widget"
            Range("B1").Value = "Price"
            Range("B2").Value = 19.99
        End Sub
        
        ''Applying Conditional Formatting
        ''This code applies a formatting rule that highlights cells in column B with values greater than 50 in red.
        Sub ApplyConditionalFormatting()
            Dim rng As Range
            Set rng = Range("B2:B10")

            rng.FormatConditions.Delete 'Clear existing rules
            rng.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="50"
            rng.FormatConditions(1).Font.Color = RGB(255, 255, 255) 'White font
            rng.FormatConditions(1).Interior.Color = RGB(255, 0, 0) 'Red background
        End Sub
    End Sub
End Module
