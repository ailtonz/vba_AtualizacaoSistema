Attribute VB_Name = "Testes"

Sub teste2()

'PermissoesUsuarios "C:\Users\user\Desktop\Orcamentos\db\dbOrcamentos-v03_130729-1108.mdb", "abc", "PERMISSÃO DE USUÁRIOS"

'CadastrarIntervalosEdicoes "C:\Users\user\Desktop\Orcamentos\db\dbOrcamentos-v03_130729-1108.mdb", "abc", "CADASTRO - INTERVALOS DE EDIÇÃO"

'RelacionarEtapas "C:\Users\user\Desktop\Orcamentos\db\dbOrcamentos-v03_130729-1108.mdb", "abc", "RELACIONAR - ETAPAS"

End Sub


Sub teste()
Dim ws As Worksheet
Dim wsExc As Worksheet

Dim strGuiaTarefa As String

Dim lRow As Long
Dim cCol As Long

''SELEÇÃO DA GUIA DE APOIO
Set ws = Worksheets(GuiaApoio)

''CAMINHO DO BANCO DE DADOS
Dim strBanco As String: strBanco = ws.Range(BancoLocal).Value

'ENCONTRAR PRIMEIRA LINHA VAZIA NA GUIA
lRow = ws.Cells(Rows.count, 5).End(xlUp).Offset(1, 0).Row

For x = 2 To lRow - 1
    With ws

        ''SELECIONAR GUIA - " TAREFA "
        strGuiaTarefa = .Cells(x, 5).Value
        Set wsExc = Worksheets(strGuiaTarefa)
        wsExc.Activate
        
'        ''ENCONTRAR ULTIMA COLUNA
'        cCol = wsExc.Cells(1, Columns.Count).End(xlToLeft).Column
'
'        MsgBox cCol

        MsgBox wsExc.Range("A1")


    End With
Next x

ws.Activate


End Sub


'
'Sub teste2()
'
'    MsgBox Worksheets("APOIO").Range(BancoLocal).Value
'
'End Sub
'
'Sub teste()
'
'Dim lRow As Long
'Dim lPart As Long
'Dim ws As Worksheet
'Dim x As Long
'
''Set ws = Worksheets(ActiveSheet.Name)
'Set ws = Worksheets("APOIO")
'
''find  first empty row in database
'lRow = ws.Cells(Rows.Count, 2).End(xlUp).Offset(1, 0).Row
'
'
'For x = 1 To lRow - 1
'    With ws
'        MsgBox .Cells(x, 1).Value
'    End With
'
'Next x
'
''With ws
''    If Me.txtData.Value <> "" Then .Cells(lRow, 1).Value = Format(CDate(Me.txtData.Value), "mm/dd/yy")
''    .Cells(lRow, 2).Value = Me.cboTelefones.Value
''    .Cells(lRow, 3).Value = CCur(Me.txtValor.Value)
''End With
'
'ThisWorkbook.Save
'
'End Sub
'
'
'Sub teste4()
'Dim ws As Worksheet
'Dim cCol As Long
'
'Set ws = Worksheets("Plan1")
'ws.Activate
'
'MsgBox ws.Cells(1, Columns.Count).End(xlToLeft).Column
'
'
'End Sub
