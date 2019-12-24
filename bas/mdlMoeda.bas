Attribute VB_Name = "mdlMoeda"
Sub carregarScriptMoeda()

Dim lrow, x As Long: lrow = shMoeda.Cells(Rows.count, 2).End(xlUp).Offset(1, 0).Row

For x = 2 To lrow - 1
    With shMoeda
        If .Range("C" & x).value <> "" Then
            .Range("A" & x).value = MoedaNovo(.Range("C" & x).value, _
                                                .Range("D" & x).value)
        End If
    End With
Next

End Sub

Public Function MoedaNovo(sMoeda As String, sValor As String) As String

Dim sSQL As String

sSQL = "UPDATE admcategorias SET admcategorias.Descricao01 = '" & sValor & "' WHERE (((admcategorias.categoria)='" & sMoeda & "') AND ((admcategorias.codRelacao)=(SELECT admCategorias.codCategoria FROM admCategorias Where Categoria='MOEDA' and codRelacao = 0)))"

MoedaNovo = sSQL

End Function

