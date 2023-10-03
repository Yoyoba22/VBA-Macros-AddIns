Attribute VB_Name = "ConvertKey1"
Sub ConvertKey11()
Attribute ConvertKey11.VB_ProcData.VB_Invoke_Func = "J\n14"
'
' ConvertKey11 Makro
'
' Tastenkombination: Strg+Umschalt+J
'
Selection.TextToColumns Destination:=ActiveCell, DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="(", FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
    ActiveCell.Offset(0, 1).Range("A1:A10").Select
    Selection.ClearContents
    ActiveCell.Select
    ActiveCell.FormulaR1C1 = "=LEFT(RC[-1],LEN(RC[-1])-1)"
    ActiveCell.Select
    Selection.AutoFill Destination:=ActiveCell.Range("A1:A20")
    ActiveCell.Range("A1:A20").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(LEFT(RC[-1],LEN(RC[-1])-1)),"""",(LEFT(RC[-1],LEN(RC[-1])-1)))"
    ActiveCell.Select
    Selection.AutoFill Destination:=ActiveCell.Range("A1:A20")
    ActiveCell.Range("A1:A20").Select
    
Dim rng As Range, ar As Range

Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual

If Selection.Cells.Count = 1 Then
   Set rng = Selection
Else
   Set rng = Selection.SpecialCells(xlCellTypeVisible)
End If

For Each ar In rng.Areas
    ar.Value = ar.Value
Next ar
  
Application.ScreenUpdating = True
Application.EnableEvents = True
Application.Calculation = xlCalculationAutomatic
Application.CalculateFull
    
    
End Sub
