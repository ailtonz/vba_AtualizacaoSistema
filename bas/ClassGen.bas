Attribute VB_Name = "ClassGen"
Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Option Explicit


'Sub TESTE()
'
'    NewClass "CADASTRO", Array("TIPO", "STRING", "VALOR", "STRING", "SCRIPT", "STRING")
'
'
'End Sub


Function NewClass(pName As String, pProperties As Variant)
    Dim strCaminho As String
    Dim strGetLet As String
    Dim intLinha1 As Integer
    Dim intLinha2 As Integer
    Dim intLinha3 As Integer
    Dim i As Double

    On Error Resume Next
    Kill ThisWorkbook.Path & "\*.cls"
    On Error GoTo 0
        
    strCaminho = ThisWorkbook.Path & "\" & pName & ".cls"
    Open strCaminho For Output As #1
    Print #1, "VERSION 1.0 CLASS"
    Print #1, "BEGIN"
    Print #1, "  MultiUse = -1  'True"
    Print #1, "END"
    Print #1, "Attribute VB_Name = " & Chr(34) & pName & Chr(34)
    Print #1, "Attribute VB_GlobalNameSpace = False"
    Print #1, "Attribute VB_Creatable = False"
    Print #1, "Attribute VB_PredeclaredId = False"
    Print #1, "Attribute VB_Exposed = False"
    Print #1, "Option Explicit"
    Print #1, vbNullString
    Print #1, FormataAtributo(True, "clsBanco", "Banco", 42)
    Print #1, FormataAtributo(True, "Col", "New Collection", 42)
    'Print #1, "Private iCol As New Collection"
    
    For i = LBound(pProperties) To UBound(pProperties) Step 2
        Print #1, FormataAtributo(True, CStr(pProperties(i)), CStr(pProperties(i + 1)), 42)
    Next i
    'Print #1, vbNullString
    
    '##- Gera os Gets e Lets
    
    intLinha1 = 80
    intLinha2 = 150
    intLinha3 = 180
    
    For i = LBound(pProperties) To UBound(pProperties) Step 2
        strGetLet = ""
        Select Case UCase(Trim(pProperties(i + 1)))
            Case "STRING", "INTEGER", "DOUBLE", "DATE", "BOOLEAN", "LONG"
            
                strGetLet = "Property Get " & pProperties(i) & "() As " & pProperties(i + 1) & ":"
                strGetLet = strGetLet & WorksheetFunction.Rept(" ", intLinha1 - Len(strGetLet))
                strGetLet = strGetLet & pProperties(i) & " = i" & pProperties(i) & ":"
                strGetLet = strGetLet & WorksheetFunction.Rept(" ", intLinha2 - Len(strGetLet))
                strGetLet = strGetLet & "End Property"
                Print #1, strGetLet
                
                strGetLet = "Property Let " & pProperties(i) & "(pValue As " & pProperties(i + 1) & ")" & ":"
                strGetLet = strGetLet & WorksheetFunction.Rept(" ", intLinha1 - Len(strGetLet))
                strGetLet = strGetLet & pProperties(i) & " = i" & pProperties(i) & ":"
                strGetLet = strGetLet & WorksheetFunction.Rept(" ", intLinha2 - Len(strGetLet))
                strGetLet = strGetLet & "End Property"
                Print #1, strGetLet

            Case Else
            
                strGetLet = "Property Get " & pProperties(i) & "() as " & pProperties(i + 1) & ":"
                strGetLet = strGetLet & WorksheetFunction.Rept(" ", intLinha1 - Len(strGetLet))
                strGetLet = strGetLet & "If i" & pProperties(i) & " Is Nothing Then Set i" & pProperties(i) & " = " & pProperties(i + 1) & ":"
                strGetLet = strGetLet & WorksheetFunction.Rept(" ", intLinha2 - Len(strGetLet))
                strGetLet = strGetLet & "Set " & pProperties(i) & " = i" & pProperties(i) & ":"
                strGetLet = strGetLet & WorksheetFunction.Rept(" ", intLinha3 - Len(strGetLet))
                strGetLet = strGetLet & "End Property"
                Print #1, strGetLet
                
                strGetLet = "Property Let " & pProperties(i) & "(pValue as " & pProperties(i + 1) & ")" & ":"
                strGetLet = strGetLet & WorksheetFunction.Rept(" ", intLinha1 - Len(strGetLet))
                strGetLet = strGetLet & "Set i" & pProperties(i) & " = pValue" & ":"
                strGetLet = strGetLet & WorksheetFunction.Rept(" ", intLinha2 - Len(strGetLet))
                strGetLet = strGetLet & "End Property"
                Print #1, strGetLet
                
            
'                Print #1, "Property Get " & pProperties(i) & "() as Collection"
'                Print #1, "    If i" & pProperties(i) & " Is Nothing Then Set i" & pProperties(i) & " = " & pProperties(i + 1)
'                Print #1, "    Set " & pProperties(i) & " = i" & pProperties(i)
'                Print #1, "End Property"
'                Print #1, "Property Let " & pProperties(i) & "(pValue as " & pProperties(i + 1) & ")"
'                Print #1, "    Set i" & pProperties(i) & " = pValue"
'                Print #1, "End Property"
        End Select
    Next i
    
    Print #1, vbNullString
    Print #1, "''---------------"
    Print #1, "'' administra��o"
    Print #1, "''---------------"
    Print #1, vbNullString
    Print #1, "Public Function NewEnum() As IUnknown"
    Print #1, "Attribute NewEnum.VB_UserMemId = -4"
    Print #1, "    Set NewEnum = iCol.[_NewEnum]"
    Print #1, "End Function"
    Print #1, vbNullString
    Print #1, "Private Sub Class_Initialize()"
    Print #1, "    Set iCol = New Collection"
    Print #1, "End Sub"
    Print #1, vbNullString
    Print #1, "Private Sub Class_Terminate()"
    Print #1, "    Set iCol = Nothing"
    Print #1, "End Sub"
    Print #1, vbNullString
    Print #1, "Public Sub add(ByVal rec As " & pName & ", Optional ByVal key As Variant, Optional ByVal before As Variant, Optional ByVal after As Variant)"
    Print #1, "    iCol.add rec, key, before, after"
    Print #1, "End Sub"
    Print #1, vbNullString
    Print #1, "Public Sub all(ByVal rec As " & pName & ")"
    Print #1, "    iCol.add rec, rec.id"
    Print #1, "End Sub"
    Print #1, vbNullString
    Print #1, "Public Function count() As Long"
    Print #1, "    count = iCol.count"
    Print #1, "End Function"
    Print #1, vbNullString
    Print #1, "Public Sub remove(ByVal i As Integer)"
    Print #1, "    iCol.remove i"
    Print #1, "End Sub"
    Print #1, vbNullString
    Print #1, "Public Function " & pName & "(ByVal i As Variant) As " & pName & ""
    Print #1, "    Set " & pName & " = iCol.item(i)"
    Print #1, "End Function"
    Print #1, vbNullString
    Print #1, "Public Property Get Itens() As Collection"
    Print #1, "    Set Itens = iCol"
    Print #1, "End Property"
    Print #1, vbNullString
    Print #1, "Public Property Get item(i As Variant) As " & pName & ""
    Print #1, "    Set item = iCol(i)"
    Print #1, "End Property"
    Print #1, vbNullString
    Print #1, vbNullString
    Print #1, "''---------------"
    Print #1, "'' FUN��ES"
    Print #1, "''---------------"

    
    Close #1
    
    DoEvents: Sleep 100
    DoEvents: Sleep 100
    
    ImportaClasse ActiveWorkbook, strCaminho, pName
    Kill strCaminho

End Function
Function FormataAtributo(pPrivate As Boolean, pNome As String, pTipo As String, Optional pTamanho As Integer) As String
    Dim iPrefixo As String
    iPrefixo = IIf(pPrivate, "Private ", "Public ") & "i" & PrimMaiusc(pNome)
    If pTamanho - Len(iPrefixo) >= 0 Then
        FormataAtributo = iPrefixo & WorksheetFunction.Rept(" ", pTamanho - Len(iPrefixo)) & "As " & PrimMaiusc(pTipo)
    Else
        FormataAtributo = iPrefixo & " As " & PrimMaiusc(pTipo)
    End If
End Function
Function ImportaClasse(ByVal wb As Workbook, ByVal CompFileName As String, pNomeClasse As String)
    If Dir(CompFileName) <> vbNullString Then
        On Error Resume Next
        wb.VBProject.VBComponents.remove wb.VBProject.VBComponents(pNomeClasse)
        DoEvents
        wb.VBProject.VBComponents.Import CompFileName
        On Error GoTo 0
    End If
    Set wb = Nothing
End Function
Function QuebraLinha(pTexto As String) As String
    QuebraLinha = "' " & WorksheetFunction.Rept("-", 90 - Len(pTexto)) & " " & pTexto
End Function
Function PrimMaiusc(pTexto As String) As String
    PrimMaiusc = UCase(Left(pTexto, 1)) & Right(pTexto, Len(pTexto) - 1)
End Function
