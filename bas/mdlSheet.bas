Attribute VB_Name = "mdlSheet"
'Callback for comboBox getText
Sub GetSheet(control As IRibbonControl, ByRef returnedVal)
    returnedVal = GetSheetNome
End Sub
'Callback for comboBox onChange
Sub SetSheet(control As IRibbonControl, text As String)
    SetSheetNome text
End Sub

Public Function SetSheetNome(pSheet As String)
    ThisWorkbook.Names("userSheet").value = pSheet
    ThisWorkbook.Save
End Function

Public Function GetSheetNome() As String
    GetSheetNome = Replace(Replace(ThisWorkbook.Names("userSheet").value, "=", ""), Chr(34), "")
End Function

Public Function GetSheetID(pSheet As String) As Integer
    Dim dict As Dictionary
    Set dict = New Dictionary
    With dict
        .add "CLIENTE", 10
        .add "INDICE", 11
        .add "LINHA", 12
        .add "RELACIONAMENTO", 13
        .add "VENDEDOR", 14
    End With
    
    GetSheetID = dict.item(pSheet)
End Function

