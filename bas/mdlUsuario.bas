Attribute VB_Name = "mdlUsuario"
Sub carregarScriptUsuarioNovo()

Dim lrow, x As Long: lrow = shVendedorNovo.Cells(Rows.count, 2).End(xlUp).Offset(1, 0).Row

For x = 2 To lrow - 1
    With shVendedorNovo
        If .Range("C" & x).value <> "" Then
            .Range("A" & x).value = UsuarioNovo(.Range("C" & x).value, _
                                                .Range("D" & x).value, _
                                                .Range("E" & x).value, _
                                                .Range("F" & x).value, _
                                                .Range("G" & x).value, _
                                                .Range("H" & x).value, _
                                                .Range("I" & x).value, _
                                                .Range("J" & x).value)
        End If
    End With
Next

End Sub

Public Function UsuarioNovo(CODUSUARIO As String, NOME_USUARIO As String, EMAIL_USUARIO As String, G_CONTAS As String, TELEFONE As String, CEL_01 As String, CEL_02 As String, ID_NEXTEL As String) As String

Dim sSQL As String

sSQL = "UPDATE qryUsuarios " & _
"   SET qryUsuarios.Usuario = '" & UCase(NOME_USUARIO) & "' " & _
"   , qryUsuarios.Codigo = UCase('" & UCase(CODUSUARIO) & "') " & _
"   , qryUsuarios.eMail = LCase('" & LCase(EMAIL_USUARIO) & "')  " & _
"   , qryUsuarios.TELEFONE = '" & TELEFONE & "'  " & _
"   , qryUsuarios.CEL_01 = '" & CEL_01 & "'  " & _
"   , qryUsuarios.CEL_02 = '" & CEL_02 & "'  " & _
"WHERE (((qryUsuarios.DPTO)='VENDAS'));  "

UsuarioNovo = sSQL

End Function

Public Function UsuarioRelacionar(NOME_USUARIO As String) As String

Dim sSQL As String

sSQL = "UPDATE qryUsuariosUsuarios SET qryUsuariosUsuarios.Usuarios = '" & NOME_USUARIO & "' " & _
"WHERE (((qryUsuariosUsuarios.Usuario)='" & NOME_USUARIO & "'));"

UsuarioRelacionar = sSQL

End Function


Sub carregarScriptUsuarioAdministrador()

Dim lrow, x As Long: lrow = shVendedorRelacionar.Cells(Rows.count, 2).End(xlUp).Offset(1, 0).Row

For x = 2 To lrow - 1
    With shVendedorRelacionar
        If .Range("C" & x).value <> "" Then
            .Range("A" & x).value = UsuarioAdministrador(.Range("C" & x).value, _
                                                .Range("D" & x).value)
        End If
    End With
Next

End Sub

Public Function UsuarioAdministrador(NOME_USUARIO As String, NOME_ADMINISTRADOR As String) As String

Dim sSQL As String

sSQL = "INSERT INTO admCategorias ( Categoria, Descricao01, codRelacao, codCategoria ) " & _
" SELECT 'Usuarios' AS NM_CATEGORIA, '" & LCase(NOME_ADMINISTRADOR) & "' AS Usuario, qryUsuarios.codCategoria, Format(IIf(IsNull([NOVO_CONTROLE]),1,[NOVO_CONTROLE]),'000') AS controle " & _
" FROM qryUsuarios, admSubNumeroDeCategoriaNovo " & _
" WHERE (((qryUsuarios.Usuario)='" & LCase(NOME_USUARIO) & "')) ORDER BY 1; "


'sSQL = "INSERT INTO admCategorias ( Categoria, Descricao01, codRelacao, codCategoria ) " & _
'" SELECT 'Usuarios' AS NM_CATEGORIA, '" & LCase(NOME_ADMINISTRADOR) & "' AS Usuario, qryUsuarios.codCategoria, Format(IIf(IsNull([NOVO_CONTROLE]),1,[NOVO_CONTROLE]),'000') AS controle " & _
'" FROM qryUsuarios, admSubNumeroDeCategoriaNovo " & _
'" WHERE (((qryUsuarios.Usuario)='" & LCase(NOME_USUARIO) & "')) ORDER BY [NM_ADMINISTRADOR]; "


UsuarioAdministrador = sSQL

End Function

