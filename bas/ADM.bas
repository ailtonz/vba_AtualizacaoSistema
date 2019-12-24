Attribute VB_Name = "ADM"
''CONSTANTES
Public Const BancoLocal As String = "A2"
Public Const SenhaBanco As String = "abc"
Public Const GuiaApoio As String = "Apoio"

Sub Atualizar(ByVal Control As IRibbonControl)
    
    AtualizarBanco GuiaApoio
    MsgBox "Banco atualizado!"
End Sub

Sub ExecutarTarefas(ByVal Control As IRibbonControl)
Dim ws As Worksheet
Dim strBanco As String
Dim strSheet As String: strSheet = ActiveSheet.Name

Set ws = Worksheets(GuiaApoio)
strBanco = ws.Range(BancoLocal).Value

Select Case strSheet

    Case "ADM - EDI��O X ETAPA"
        RelacionarEtapas strBanco, SenhaBanco, strSheet

    Case "ADM - PERMISS�ES DE USU�RIOS"
        PermissoesUsuarios strBanco, SenhaBanco, strSheet
        
    Case "ADM - EXCLUIR PERMISS�ES"
        PermissoesUsuarios strBanco, SenhaBanco, strSheet
                
    Case "ADM - CADASTROS DE FORMULARIOS"
        CadastrarFormulario strBanco, SenhaBanco, strSheet
        
    Case "ADM - INTERVALOS DE EDI��ES"
        CadastrarIntervaloEdicao strBanco, SenhaBanco, strSheet
        
End Select

MsgBox "Tarefas concluidas!"

End Sub


Function AtualizarBanco(ByVal strGuia As String)

Dim db As DAO.Database
Dim fd As Office.FileDialog
Dim ws As Worksheet
Dim lRow As Long

Dim strBanco As String
Dim strQry As String
Dim strSQL As String
Dim strDbDestino As String


Inicio:

Set ws = Worksheets(strGuia)
strBanco = ws.Range(BancoLocal).Value

'SELECIONAR O BANCO
If Not getFileStatus(strBanco) Then

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Filters.Clear
    fd.Filters.Add "BDs do Access", "*.MDB"
    fd.Title = "Por favor selecione a base de dados para uso da planilha."
    fd.AllowMultiSelect = False
    
    'ATUALIZAR CAMINHO DO BANCO
    If fd.Show = -1 Then
        ws.Range(BancoLocal).Value = fd.SelectedItems(1)
        ThisWorkbook.Save
        GoTo Inicio
    End If
    
'ATUALIZAR BANCO
Else
    ''
    bkp
    
    'CARREGAR BANCO
    Set db = DBEngine.OpenDatabase(strBanco, False, False, "MS Access;PWD=" & SenhaBanco)
        
        
    'ENCONTRAR PRIMEIRA LINHA VAZIA NA GUIA
    lRow = ws.Cells(Rows.Count, 2).End(xlUp).Offset(1, 0).Row

    'CARREGAR PARAMETROS DAS NOVAS CONSULTAS
    For x = 2 To lRow - 1
        With ws
            'NOME DA CONSULTA
            strQry = .Cells(x, 2).Value

            'COMANDOS DA CONSULTA
            strSQL = .Cells(x, 3).Value

            'VERIFICAR A EXISTENCIA DA CONSULTA NO BANCO
            If Not qryExists(strQry, strBanco, SenhaBanco) Then
                'CRIAR CONSULTA NO BANCO DE DADOS
                db.CreateQueryDef strQry, strSQL
                'SE MARCADO "CONSULTA DDL" EXECUTAR CONSULTA
                If .Cells(x, 4).Value = "x" Then db.QueryDefs(strQry).Execute
            Else
                'EXCLUS�O DE CONSULTA
                db.QueryDefs.Delete strQry
                'CRIAR CONSULTA NO BANCO DE DADOS
                db.CreateQueryDef strQry, strSQL
            End If

            'MARCAR CASO N�O DDL
            If .Cells(x, 4).Value <> "x" Then .Cells(x, 4).Value = "OK"

        End With

    Next x
    
    db.Close
    
    Set db = Nothing

End If

End Function



Public Function RelacionarEtapas(ByVal strCaminhoBanco As String, ByVal strSenhaBanco As String, ByVal strGuiaTarefa As String)
On Error GoTo admUsuariosPermissoes_err

Dim ws As Worksheet
Dim db As DAO.Database
Dim qdf As DAO.QueryDef

''BANCO DE DADOS
Set db = DBEngine.OpenDatabase(strCaminhoBanco, False, False, "MS Access;PWD=" & strSenhaBanco)

''NOME DA GUIA
Set ws = Worksheets(strGuiaTarefa)

''NOME DA CONSULTA
Dim strNomeConsulta As String: strNomeConsulta = ws.Range("A1")
Set qdf = db.QueryDefs(strNomeConsulta)

'ENCONTRAR PRIMEIRA LINHA VAZIA NA GUIA
lRow = ws.Cells(Rows.Count, 2).End(xlUp).Offset(1, 0).Row

''PARAMETROS DA CONSULTA
For x = 2 To lRow - 1
    If (ws.Cells(x, 4).Value) = "" Then
        With qdf
            .Parameters("NM_ETAPA") = ws.Cells(x, 2).Value
            .Parameters("INTERVALO_EDICAO") = ws.Cells(x, 3).Value
            .Execute
            ws.Cells(x, 4).Value = "ok"
        End With
    End If
Next x

qdf.Close
db.Close

admUsuariosPermissoes_Fim:

    Set db = Nothing
    Set qdf = Nothing
    Set ws = Nothing
    
    Exit Function
admUsuariosPermissoes_err:
    MsgBox Err.Description
    Resume admUsuariosPermissoes_Fim
End Function


Public Function CadastrarFormulario(ByVal strCaminhoBanco As String, ByVal strSenhaBanco As String, ByVal strGuiaTarefa As String)
On Error GoTo admUsuariosPermissoes_err

Dim ws As Worksheet
Dim db As DAO.Database
Dim qdf As DAO.QueryDef

''BANCO DE DADOS
Set db = DBEngine.OpenDatabase(strCaminhoBanco, False, False, "MS Access;PWD=" & strSenhaBanco)

''NOME DA GUIA
Set ws = Worksheets(strGuiaTarefa)

''NOME DA CONSULTA
Dim strNomeConsulta As String: strNomeConsulta = ws.Range("A1")
Set qdf = db.QueryDefs(strNomeConsulta)

'ENCONTRAR PRIMEIRA LINHA VAZIA NA GUIA
lRow = ws.Cells(Rows.Count, 2).End(xlUp).Offset(1, 0).Row

''PARAMETROS DA CONSULTA
For x = 2 To lRow - 1
    If (ws.Cells(x, 6).Value) = "" Then
        With qdf
            .Parameters("NM_FORMULARIO") = ws.Cells(x, 2).Value
            .Parameters("NM_INICIO") = ws.Cells(x, 3).Value
            .Parameters("NM_TERMINIO") = ws.Cells(x, 4).Value
            .Parameters("SN_OCULTO") = ws.Cells(x, 5).Value
            .Execute
            ws.Cells(x, 6).Value = "ok"
        End With
    End If
Next x

qdf.Close
db.Close

admUsuariosPermissoes_Fim:

    Set db = Nothing
    Set qdf = Nothing
    Set ws = Nothing
    
    Exit Function
admUsuariosPermissoes_err:
    MsgBox Err.Description
    Resume admUsuariosPermissoes_Fim
End Function

Public Function CadastrarIntervaloEdicao(ByVal strCaminhoBanco As String, ByVal strSenhaBanco As String, ByVal strGuiaTarefa As String)
On Error GoTo admUsuariosPermissoes_err

Dim ws As Worksheet
Dim db As DAO.Database
Dim qdf As DAO.QueryDef
Dim strValor As String

''BANCO DE DADOS
Set db = DBEngine.OpenDatabase(strCaminhoBanco, False, False, "MS Access;PWD=" & strSenhaBanco)

''NOME DA GUIA
Set ws = Worksheets(strGuiaTarefa)

''NOME DA CONSULTA
Dim strNomeConsulta As String: strNomeConsulta = ws.Range("A1")
Set qdf = db.QueryDefs(strNomeConsulta)

'ENCONTRAR PRIMEIRA LINHA VAZIA NA GUIA
lRow = ws.Cells(Rows.Count, 2).End(xlUp).Offset(1, 0).Row

''PARAMETROS DA CONSULTA
For x = 2 To lRow - 1
    If (ws.Cells(x, 7).Value) = "" Then
        With qdf
            .Parameters("NM_EDICAO") = ws.Cells(x, 2).Value
            .Parameters("NM_SELECAO") = Replace(ws.Cells(x, 3).Value, ";", ",")
            .Parameters("VL_PADRAO") = ws.Cells(x, 4).Value
            .Parameters("SN_EDITAR") = ws.Cells(x, 5).Value
            .Parameters("SN_ATUALIZAR") = ws.Cells(x, 6).Value
            .Execute
            ws.Cells(x, 7).Value = "ok"
        End With
    End If
Next x

qdf.Close
db.Close

admUsuariosPermissoes_Fim:

    Set db = Nothing
    Set qdf = Nothing
    Set ws = Nothing
    
    Exit Function
admUsuariosPermissoes_err:
    MsgBox Err.Description
    Resume admUsuariosPermissoes_Fim
End Function

Public Function PermissoesUsuarios(ByVal strCaminhoBanco As String, ByVal strSenhaBanco As String, ByVal strGuiaTarefa As String)
On Error GoTo admUsuariosPermissoes_err

Dim ws As Worksheet
Dim db As DAO.Database
Dim qdf As DAO.QueryDef

''BANCO DE DADOS
Set db = DBEngine.OpenDatabase(strCaminhoBanco, False, False, "MS Access;PWD=" & strSenhaBanco)

''NOME DA GUIA
Set ws = Worksheets(strGuiaTarefa)

''NOME DA CONSULTA
Dim strNomeConsulta As String: strNomeConsulta = ws.Range("A1")
Set qdf = db.QueryDefs(strNomeConsulta)

'ENCONTRAR PRIMEIRA LINHA VAZIA NA GUIA
lRow = ws.Cells(Rows.Count, 2).End(xlUp).Offset(1, 0).Row

''PARAMETROS DA CONSULTA
For x = 2 To lRow - 1
    If (ws.Cells(x, 5).Value) = "" Then
        With qdf
            .Parameters("NM_USUARIO") = ws.Cells(x, 2).Value
            .Parameters("NM_CATEGORIA") = ws.Cells(x, 3).Value
            .Parameters("NM_PERMISSAO") = ws.Cells(x, 4).Value
            .Execute
            ws.Cells(x, 5).Value = "ok"
        End With
    End If
Next x

qdf.Close
db.Close

admUsuariosPermissoes_Fim:

    Set db = Nothing
    Set qdf = Nothing
    Set ws = Nothing
    
    Exit Function
admUsuariosPermissoes_err:
    MsgBox Err.Description
    Resume admUsuariosPermissoes_Fim
End Function

Sub bkp()
Dim ws As Worksheet

Set ws = Worksheets(GuiaApoio)

Dim strControle As String: strControle = Controle
Dim strBanco As String: strBanco = ws.Range(BancoLocal).Value
Dim strDbDestino As String: strDbDestino = getPath(strBanco) & strControle & ".zip"
    
Compact strBanco, strDbDestino

End Sub
