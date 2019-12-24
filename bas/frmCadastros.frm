VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCadastros 
   Caption         =   "CADASTROS"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5595
   OleObjectBlob   =   "frmCadastros.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCadastros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dict As Dictionary

'Public cCads        As New Collection
'Private sCads       As New CADASTRO
'
'
'Private Sub btnAdicionar_Click()
'
'    For Each sCads In cCads
'        With sCads
'            .TIPO = "Cadastro"
'            .VALOR = "cliente"
'            .SCRIPT = "insert into CADASTRO (ID,NOME,TELEFONE) VALUES (1,'AILTON','555-8888')"
'        End With
'    Next
'
'End Sub
'
'
'Private Sub UserForm_Initialize()
''Set colBaseFixa = New Collection
'
'
'End Sub

Private Sub btnAdicionar_Click()

Dim sTitle As String:       sTitle = "CADASTRO"
Dim sMessage As String:     sMessage = "Entre com o dado:"
Dim sDefault As String:     sDefault = "" 'Replace(Replace(ThisWorkbook.Names("userBancoSenha").value, "=", ""), Chr(34), "")
Dim response As Variant:    response = InputBox(sMessage, sTitle, (sDefault), 100, 100)

Dim cad As New CADASTRO
    
Select Case StrPtr(response)

    Case 0
         MsgBox "Atualização cancelada.", 64, sTitle
        Exit Sub
        
    Case Else
    
        cad.PreencheListbox Me.lstCadastros, CStr(response)
        MsgBox "Item adicionado: " & (response) & ".", 64, sTitle

End Select


End Sub

Private Sub btnRemover_Click()

Dim cad As New CADASTRO

cad.RemoverListbox Me.lstCadastros

End Sub

Private Sub UserForm_Initialize()
Dim cad As New CADASTRO

cad.PreencheListbox Me.lstCadastros, CStr("aa")
cad.PreencheListbox Me.lstCadastros, CStr("bb")
cad.PreencheListbox Me.lstCadastros, CStr("cc")
cad.PreencheListbox Me.lstCadastros, CStr("dd")
cad.PreencheListbox Me.lstCadastros, CStr("ee")
cad.PreencheListbox Me.lstCadastros, CStr("ff")

Set cad = Nothing



End Sub
