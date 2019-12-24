VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Banco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Const MAX_TENTATIVAS        As Integer = 3
Private Const MAX_SLEEP             As Long = 500
Private Const ERR_SELECT            As Long = 1
Private Const ERR_EXECUTE           As Long = 2
Private Const ERR_CONECT            As Long = 3

Private iSql                        As String
Private iCn                         As New ADODB.Connection
Private iRs                         As New ADODB.Recordset

Property Get sql() As String
    sql = iSql
End Property
Property Let sql(value As String)
    iSql = value
End Property
Property Get Rs() As ADODB.Recordset
    Set Rs = iRs
End Property
'================================================================================================================
'=   FIM DAS PROPRIEDADES DA CLASSE   ===========================================================================
'================================================================================================================
Public Function sel(Optional sql As String, Optional pManterConexao As Boolean)
    Dim intTentativas As Integer
    If sql <> vbNullString Then iSql = sql
    If iSql = vbNullString Then Exit Function
    On Error GoTo ErrHandler
        
    conecta
    Set iRs = Nothing
    iRs.CursorLocation = adUseClient

    Do Until intTentativas = MAX_TENTATIVAS
        Err.Clear
        iRs.Open iSql, iCn, ADODB.adOpenForwardOnly, adLockReadOnly
        If Err.Number = 0 Then
            Set iRs.ActiveConnection = Nothing
            If Not pManterConexao Then fecha
            Exit Do
        Else
            intTentativas = intTentativas + 1
            Sleep MAX_SLEEP
        End If
    Loop
Exit Function
ErrHandler:
    If intTentativas = MAX_TENTATIVAS Then
        Debug.Print " -- Erro " & Err.Description & " " & Now() & " -- " & vbNewLine & iSql
        Err.Raise vbObjectError + 512 + ERR_SELECT
    Else
        intTentativas = intTentativas + 1
        Sleep MAX_SLEEP
        Resume
    End If
End Function
Public Function Exe(Optional sql As String, Optional pManterConexao As Boolean)
    Dim intTentativas As Integer
    If sql <> vbNullString Then iSql = sql
    If iSql = vbNullString Then Exit Function
    On Error GoTo ErrHandler
    
    conecta
    
    Do Until intTentativas = MAX_TENTATIVAS
        Err.Clear
        iCn.Execute iSql, adExecuteNoRecords
        If Err.Number = 0 Then
            If Not pManterConexao Then fecha
            Exit Do
        Else
            intTentativas = intTentativas + 1
            Sleep MAX_SLEEP
        End If
    Loop
Exit Function
ErrHandler:
    If intTentativas = MAX_TENTATIVAS Then
        Debug.Print " -- Erro " & Err.Description & " " & Now() & " -- " & vbNewLine & iSql
        Err.Raise vbObjectError + 512 + ERR_EXECUTE
    Else
        intTentativas = intTentativas + 1
        Sleep MAX_SLEEP
        Resume
    End If
End Function
Public Function conecta()
    Dim intTentativas As Integer
    On Error GoTo ErrHandler
    If iCn.State = 0 Then
        iCn.ConnectionTimeout = 3600
        iCn.CommandTimeout = 3600

        Do Until intTentativas = MAX_TENTATIVAS
            iCn.Open
            Exit Do
        Loop
    End If
Exit Function
ErrHandler:
    If intTentativas = MAX_TENTATIVAS Then
        Debug.Print " -- Erro " & Err.Description & " " & Now() & " -- "
        Err.Raise vbObjectError + 512 + ERR_CONECT
    Else
        intTentativas = intTentativas + 1
        Sleep MAX_SLEEP
        Resume
    End If
End Function
Public Function fecha()
    On Error Resume Next
        iCn.Close
    On Error GoTo 0
End Function
Private Sub Class_Initialize()
   
    Dim sBancoCaminho As String:     sBancoCaminho = Replace(Replace(ThisWorkbook.Names("userBancoCaminho").value, "=", ""), Chr(34), "")
    Dim sBancoUsuario As String:     sBancoUsuario = Replace(Replace(ThisWorkbook.Names("userBancoUsuario").value, "=", ""), Chr(34), "")
    Dim sBancoSenha As String:     sBancoSenha = Replace(Replace(ThisWorkbook.Names("userBancoSenha").value, "=", ""), Chr(34), "")
    
    
    ' ## Local
'    iCn.connectionString = "Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=" & sBancoCaminho & ";"
    iCn.connectionString = "Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=" & sBancoCaminho & ";" & ";Uid=" & sBancoUsuario & ";PWD=" & sBancoSenha & ""
    
       
    
    
'    ' ## Servido Azure
'    iCn.connectionString = "Provider = SQLOLEDB" & _
                              ";Data Source=191.237.250.225" & _
                              ";Initial Catalog=SISCOB" & _
                              ";User ID=SISCOB_HML" & _
                              ";Password=@Cc3nture.[0m!"
    
'    ' ## Local
'    iCn.ConnectionString = "Provider = SQLOLEDB" & _
                          ";Data Source=MW7ADLGIKZ1B5B\SQLEXPRESS" & _
                          ";Initial Catalog=SISCOB" & _
                          ";User ID=sa" & _
                          ";Password=123456"
End Sub
