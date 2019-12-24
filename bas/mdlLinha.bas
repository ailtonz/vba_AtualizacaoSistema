Attribute VB_Name = "mdlLinha"
Option Explicit

Sub carregarScriptLinha()

Dim lrow, x As Long: lrow = shClienteNovo.Cells(Rows.count, 2).End(xlUp).Offset(1, 0).Row


For x = 2 To lrow - 1
    With shLinha
        If .Range("C" & x).value <> "" Then
            .Range("A" & x).value = LinhaNovo(.Range("C" & x).value, .Range("C" & x).value, _
                                    .Range("D" & x).value, _
                                    .Range("E" & x).value)
        End If
    End With
Next

End Sub

Public Function LinhaNovo(sAntigo As String, sNovo As String, sVal01 As String, sVal02 As String) As String

Dim sSQL As String


sSQL = "UPDATE admCategorias " & _
    " SET admCategorias.Categoria = UCase('" & sNovo & "') " & _
    " , admCategorias.Descricao01 = '" & sVal01 & "' " & _
    " , admCategorias.Descricao02 = '" & sVal02 & "' " & _
    " WHERE  " & _
    " (((admCategorias.Categoria)='" & sAntigo & "')  " & _
    " AND  " & _
    " ((admCategorias.codRelacao)=(SELECT admCategorias.codCategoria FROM admCategorias Where Categoria = 'LINHA' and codRelacao = 0)));"

LinhaNovo = sSQL


End Function



