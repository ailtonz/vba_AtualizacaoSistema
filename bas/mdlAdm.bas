Attribute VB_Name = "mdlAdm"
Option Explicit
Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub cmdBanco_Click(control As IRibbonControl)
Dim filename As clsPathAndFiles: Set filename = New clsPathAndFiles
Dim sBancoCaminho As String:      sBancoCaminho = Replace(Replace(ThisWorkbook.Names("userBancoCaminho").value, "=", ""), Chr(34), "")


Dim f As Variant

For Each f In filename.PickFiles("Selecione o Banco", Array("Arquivos MsAccess", "*.mdb*"), False)
    If CStr(f) <> "" Then
        ThisWorkbook.Names("userBancoCaminho").value = CStr(f)
    End If
    
Next

Set filename = Nothing

End Sub

Sub cmdBancoUsuario_Click(control As IRibbonControl)
'##############################################
' NOME DO BANCO.
'##############################################

Dim sTitle As String:       sTitle = "Usuário do Banco"
Dim sMessage As String:     sMessage = "Entre com o nome do usuário administrativo do banco:"
Dim sDefault As String:     sDefault = Replace(Replace(ThisWorkbook.Names("userBancoUsuario").value, "=", ""), Chr(34), "")
Dim response As Variant:    response = InputBox(sMessage, sTitle, (sDefault), 100, 100)
    
Select Case StrPtr(response)

    Case 0
         MsgBox "Atualização cancelada.", 64, sTitle
        Exit Sub
        
    Case Else
         ThisWorkbook.Names("userBancoUsuario").value = (response)
         MsgBox "Usuario atualizado: " & (response) & ".", 64, sTitle

End Select
    
End Sub


Sub cmdBancoSenha_Click(control As IRibbonControl)
'##############################################
' SENHA DO BANCO.
'##############################################

Dim sTitle As String:       sTitle = "Senha do Banco"
Dim sMessage As String:     sMessage = "Entre com a senha do banco:"
Dim sDefault As String:     sDefault = Replace(Replace(ThisWorkbook.Names("userBancoSenha").value, "=", ""), Chr(34), "")
Dim response As Variant:    response = InputBox(sMessage, sTitle, (sDefault), 100, 100)
    
Select Case StrPtr(response)

    Case 0
         MsgBox "Atualização cancelada.", 64, sTitle
        Exit Sub
        
    Case Else
         ThisWorkbook.Names("userBancoSenha").value = (response)
         MsgBox "Senha atualizada: " & (response) & ".", 64, sTitle

End Select
    
End Sub

Sub cmdUpdate_Click(control As IRibbonControl)
Dim objBanco As Banco: Set objBanco = New Banco
Dim ws As Worksheet: Set ws = Worksheets(GetSheetNome)
Dim lrow As Long, x As Long

ws.Visible = xlSheetVeryHidden
ws.Visible = xlSheetVisible
ws.Activate

ws.Range("A:A").Activate
ActiveCell.EntireColumn.Hidden = True
ws.Range("B2").Activate
        
lrow = ws.Cells(Rows.count, 2).End(xlUp).Offset(1, 0).Row

With objBanco
    
    For x = 2 To lrow - 1
        .Exe CStr(ws.Range("A" & x).value)
        ws.Range("B" & x).value = "OK"
    Next

End With

End Sub

Public Sub Principal()
Dim filename As clsPathAndFiles: Set filename = New clsPathAndFiles
Dim sBancoCaminho As String:     sBancoCaminho = Replace(Replace(ThisWorkbook.Names("userBancoCaminho").value, "=", ""), Chr(34), "")

If filename.FileExist(sBancoCaminho) Then
    Call UpdateGeral
Else
    MsgBox "Base não encontrada!", vbOKOnly + vbCritical, "Atualização."
End If

    
Set filename = Nothing

End Sub

Public Sub UpdateGeral()
Dim objBanco As Banco: Set objBanco = New Banco
Dim i As Integer
Dim ws As Worksheet
Dim lrow As Long, x As Long

For i = 1 To Worksheets.count
    
    If Worksheets(i).Name <> "Update" Then
        Set ws = Worksheets(Worksheets(i).Name)
        ws.Activate
        
        ws.Range("A:A").Activate
        ActiveCell.EntireColumn.Hidden = True
        ws.Range("B2").Activate
                
        lrow = ws.Cells(Rows.count, 2).End(xlUp).Offset(1, 0).Row
        
        With objBanco
            
            For x = 2 To lrow - 1
                .Exe CStr(ws.Range("A" & x).value)
                ws.Range("B" & x).value = "OK"
            Next
        
        End With
    
    End If
    
Next i

End Sub

Public Sub Update()
Dim objBanco As Banco: Set objBanco = New Banco
Dim ws As Worksheet: Set ws = Worksheets("Update")
Dim lrow, x As Long: lrow = ws.Cells(Rows.count, 2).End(xlUp).Offset(1, 0).Row

With objBanco
    
    For x = 2 To lrow - 1
        .Exe CStr(ws.Range("B" & x).value)
        ws.Range("C" & x).value = "OK"
    Next

End With

Set objBanco = Nothing

End Sub

Public Function AtivarPlanilha(Bloqueio As Boolean)
''    ActiveWindow.DisplayWorkbookTabs = Bloqueio
    
    Application.DisplayFormulaBar = Bloqueio
    
    shUpdate.Activate
    ActiveWindow.DisplayHeadings = Bloqueio
    ActiveWindow.DisplayGridlines = Bloqueio
    
    shClienteNovo.Activate
    ActiveWindow.DisplayHeadings = Bloqueio
    ActiveWindow.DisplayGridlines = Bloqueio
    
    shVendedorNovo.Activate
    ActiveWindow.DisplayHeadings = Bloqueio
    ActiveWindow.DisplayGridlines = Bloqueio
    
    shVendedorRelacionar.Activate
    ActiveWindow.DisplayHeadings = Bloqueio
    ActiveWindow.DisplayGridlines = Bloqueio
    
End Function

Sub testeWorksheets()
Dim i As Integer

    For i = 1 To Worksheets.count
        MsgBox Worksheets(i).Name
    Next i

End Sub
