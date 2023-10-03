Attribute VB_Name = "ProviderProdukt1"

' Tastenkombination: Strg+M

Sub ProviderProdukt()
Attribute ProviderProdukt.VB_ProcData.VB_Invoke_Func = "m\n14"


Dim rng As Range
Dim sTemp As String
Dim BaseURL As String

If Selection.Cells.Count = 1 Then
    Set rng = Selection
Else
    Set rng = Selection.SpecialCells(xlCellTypeVisible)
End If

    BaseURL = "https://erp.digitecgalaxus.ch/de/"

On Error Resume Next
For Each Cell In rng
    Set Cell = ActiveCell.Activate
    sTemp = Cell.Value
    Application.StatusBar = "Opening " & sTemp & ". Please wait..."
    sTemp = BaseURL & "ProviderProduct/Search?ProviderIds.rendered=&ProviderProductKeys=&Gtins=" & Cell.Value & "&ProviderBrands=&ManufacturerKeys=&MatchingState=Any&search=Suche"
    ThisWorkbook.FollowHyperlink _
    Address:=sTemp
Next Cell
Application.StatusBar = "Done."

End Sub
