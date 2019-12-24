Attribute VB_Name = "mdlOperacao"
'Callback for comboBox getText
Sub GetOperacao(control As IRibbonControl, ByRef returnedVal)
    returnedVal = GetOperacaoNome
End Sub
'Callback for comboBox onChange
Sub SetOperacao(control As IRibbonControl, text As String)
    SetOperacaoNome text
End Sub

Public Function SetOperacaoNome(pOperacao As String)
    ThisWorkbook.Names("userOperacao").value = pOperacao
    ThisWorkbook.Save
End Function

Public Function GetOperacaoNome() As String
    GetOperacaoNome = Replace(Replace(ThisWorkbook.Names("userOperacao").value, "=", ""), Chr(34), "")
End Function

Public Function GetOperacaoID(pOperacao As String) As Integer
    Dim dict As Dictionary
    Set dict = New Dictionary
    With dict
        .add "ADM", 1:   .add "UPD", 2
    End With
    
    GetOperacaoID = dict.item(pOperacao)
End Function
