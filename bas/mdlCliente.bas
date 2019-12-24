Attribute VB_Name = "mdlCliente"
Option Explicit

Sub carregarScriptCliente()

Dim lrow, x As Long: lrow = shClienteNovo.Cells(Rows.count, 2).End(xlUp).Offset(1, 0).Row


For x = 2 To lrow - 1
    With shClienteNovo
        If .Range("C" & x).value <> "" Then .Range("A" & x).value = ClienteNovo(.Range("C" & x).value)
    End With
Next

End Sub

Public Function ClienteNovo(DESCRICAO As String) As String

Dim sSQL As String


sSQL = " INSERT INTO admCategorias ( codRelacao, Categoria ) " & _
"SELECT " & _
"   TOP 1 (SELECT admCategorias.codCategoria FROM admCategorias Where Categoria='CLIENTES' and codRelacao = 0) AS idRelacao" & _
"   , '" & DESCRICAO & "' AS strDescricao FROM admCategorias"

ClienteNovo = sSQL


End Function






'Private Sub teste()
'Dim obj As CADASTRO: Set obj = New CADASTRO
'
'With obj
'    .TIPO = "Cadastro"
'    .VALOR = "cliente"
'    .SCRIPT = "insert into CADASTRO (ID,NOME,TELEFONE) VALUES (1,'AILTON','555-8888')"
'End With
'
''For Each sCads In cCol
''    With sCads
''        .TIPO = "Cadastro"
''        .VALOR = "cliente"
''        .SCRIPT = "insert into CADASTRO (ID,NOME,TELEFONE) VALUES (1,'AILTON','555-8888')"
''    End With
''Next
'
''    For Each cCol In sCads
''        With sCads
''            .TIPO = "Cadastro"
''            .VALOR = "cliente"
''            .SCRIPT = "insert into CADASTRO (ID,NOME,TELEFONE) VALUES (1,'AILTON','555-8888')"
''        End With
''    Next
'
'End Sub


'Function gerenciarConteudo(sMessage As String, sTitle As String, Optional sDefault As String) As String
'
'Dim response As Variant
'
''If sDefault = "" Then sDefault = Format(Date, "dd/mm/yy") & " - " & Format(Now(), "hh:mm")
'
'response = InputBox(sMessage, sTitle, UCase(sDefault), 100, 100)
'
'Select Case StrPtr(response)
'
'    Case 0
'         MsgBox "Atualização cancelada.", 64, sTitle
'        Exit Function
'
'    Case Else
'         gerenciarConteudo = UCase(response)
'         MsgBox "Dia cadastrado: " & UCase(response) & ".", 64, sTitle
'
'End Select
'
'End Function



