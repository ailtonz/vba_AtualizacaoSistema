VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CADASTRO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private iClsBanco                         As Banco
Private iCol                              As New Collection
Private iTIPO                             As String
Private iVALOR                            As String
Private iSCRIPT                           As String
Private iID                               As String
Property Get TIPO() As String:                                                  TIPO = iTIPO:                                                         End Property
Property Let TIPO(pValue As String):                                            TIPO = iTIPO:                                                         End Property
Property Get VALOR() As String:                                                 VALOR = iVALOR:                                                       End Property
Property Let VALOR(pValue As String):                                           VALOR = iVALOR:                                                       End Property
Property Get SCRIPT() As String:                                                SCRIPT = iSCRIPT:                                                     End Property
Property Let SCRIPT(pValue As String):                                          SCRIPT = iSCRIPT:                                                     End Property
Property Get ID() As String:                                                    ID = iID:                                                             End Property
Property Let ID(pValue As String):                                              ID = iID:                                                             End Property

''---------------
'' administração
''---------------

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = iCol.[_NewEnum]
End Function

Private Sub Class_Initialize()
    Set iCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set iCol = Nothing
End Sub

Public Sub add(ByVal rec As CADASTRO, Optional ByVal key As Variant, Optional ByVal before As Variant, Optional ByVal after As Variant)
    iCol.add rec, key, before, after
End Sub

Public Sub all(ByVal rec As CADASTRO)
    iCol.add rec, rec.ID
End Sub

Public Function count() As Long
    count = iCol.count
End Function

Public Sub remove(ByVal i As Integer)
    iCol.remove i
End Sub

Public Function CADASTRO(ByVal i As Variant) As CADASTRO
    Set CADASTRO = iCol.item(i)
End Function

Public Property Get Itens() As Collection
    Set Itens = iCol
End Property

Public Property Get item(i As Variant) As CADASTRO
    Set item = iCol(i)
End Property


''---------------
'' FUNÇÕES
''---------------

Public Function PreencheListbox(ByRef pList As MSForms.ListBox, pValor As String)
    
    With pList
        .AddItem pValor
    End With
    
End Function

Public Function RemoverListbox(ByRef pList As MSForms.ListBox)
Dim intCurrentRow As Integer

With pList
    For intCurrentRow = .ListCount - 1 To 0 Step -1
        DoEvents
        If .Selected(intCurrentRow) Then
            .RemoveItem .ListIndex
        End If
    Next intCurrentRow
End With
    
End Function
