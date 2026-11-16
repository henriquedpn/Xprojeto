Attribute VB_Name = "modUsuarios"
'==============================================================================
' Módulo: modUsuarios - cadastro, hierarquia (chefe) e permissões
'
' Permissões:
'   - Cada colaborador gerencia os próprios projetos.
'   - O chefe (coluna Gestor) gerencia os projetos dos subordinados diretos e
'     indiretos (hierarquia completa).
'   - Admin gerencia todos os projetos e o cadastro de usuários.
'==============================================================================
Option Explicit

Private mUsr As Object   ' login -> Array(nome, disciplina, gestor, perfil)

Public Sub ResetUsuarios()
    Set mUsr = Nothing
End Sub

Private Function Dic() As Object
    Dim lo As ListObject, v As Variant, i As Long, login As String
    If mUsr Is Nothing Then
        Set mUsr = CreateObject("Scripting.Dictionary")
        mUsr.CompareMode = vbTextCompare
        Set lo = ThisWorkbook.Worksheets(SH_USR).ListObjects(TBL_USR)
        If Not lo.DataBodyRange Is Nothing Then
            v = lo.DataBodyRange.Value
            For i = 1 To UBound(v, 1)
                login = LCase$(Texto(v(i, U_LOGIN)))
                If Len(login) > 0 And Not mUsr.Exists(login) Then
                    mUsr.Add login, Array(Texto(v(i, U_NOME)), Texto(v(i, U_DISC)), _
                                          LCase$(Texto(v(i, U_GESTOR))), Texto(v(i, U_PERFIL)))
                End If
            Next i
        End If
    End If
    Set Dic = mUsr
End Function

Private Function Campo(ByVal login As String, ByVal i As Long) As String
    Dim a As Variant
    If Dic().Exists(login) Then
        a = Dic().Item(login)
        Campo = a(i)
    End If
End Function

Public Function UsuarioExiste(ByVal login As String) As Boolean
    UsuarioExiste = Dic().Exists(login)
End Function

Public Function NomeUsuario(ByVal login As String) As String
    NomeUsuario = Campo(login, 0)
    If Len(NomeUsuario) = 0 Then NomeUsuario = login
End Function

Public Function DisciplinaUsuario(ByVal login As String) As String
    DisciplinaUsuario = Campo(login, 1)
End Function

Public Function GestorDe(ByVal login As String) As String
    GestorDe = Campo(login, 2)
End Function

Public Function PerfilUsuario(ByVal login As String) As String
    PerfilUsuario = Campo(login, 3)
    If Len(PerfilUsuario) = 0 Then PerfilUsuario = PERFIL_COLAB
End Function

Public Function EhAdmin(Optional ByVal login As String = "") As Boolean
    If Len(login) = 0 Then login = UsuarioAtual()
    EhAdmin = (StrComp(Campo(login, 3), PERFIL_ADMIN, vbTextCompare) = 0)
End Function

' Verdadeiro se "gestor" está acima de "colab" na hierarquia
Public Function EhSubordinado(ByVal colab As String, ByVal gestor As String) As Boolean
    Dim atual As String, passos As Long
    atual = GestorDe(colab)
    Do While Len(atual) > 0 And passos < 50          ' 50 = proteção contra ciclos
        If StrComp(atual, gestor, vbTextCompare) = 0 Then EhSubordinado = True: Exit Function
        atual = GestorDe(atual)
        passos = passos + 1
    Loop
End Function

Public Function PodeGerenciar(ByVal colab As String) As Boolean
    Dim eu As String
    eu = UsuarioAtual()
    If Not UsuarioExiste(eu) Then Exit Function
    PodeGerenciar = EhAdmin(eu) Or StrComp(colab, eu, vbTextCompare) = 0 Or EhSubordinado(colab, eu)
End Function

Public Function PodeVer(ByVal colab As String) As Boolean
    PodeVer = VerTodos() Or PodeGerenciar(colab)
End Function

Public Sub VerificarCadastroUsuario()
    If Not UsuarioExiste(UsuarioAtual()) Then
        MsgBox "Seu usuário de rede '" & UsuarioAtual() & "' ainda não está cadastrado." & vbLf & _
               "Peça ao administrador para incluí-lo na aba Usuarios." & vbLf & vbLf & _
               "Até lá você não poderá cadastrar projetos.", vbExclamation, APP_NOME
    End If
End Sub

'------------------------------------------------------------------------------
' Copia o cadastro vindo do servidor para a tabela local tblUsuarios
'------------------------------------------------------------------------------
Public Sub GravarUsuariosLocal(ByVal vUsr As Variant)
    Dim ws As Worksheet, lo As ListObject, n As Long, i As Long, j As Long, out() As Variant

    If EditandoUsuarios() Then Exit Sub   ' Admin editando: não sobrescrever a edição
    Set ws = ThisWorkbook.Worksheets(SH_USR)
    Set lo = ws.ListObjects(TBL_USR)

    If IsArray(vUsr) Then
        ReDim out(1 To UBound(vUsr, 1), 1 To U_NCOLS)
        For i = 1 To UBound(vUsr, 1)
            If Len(Texto(vUsr(i, U_LOGIN))) > 0 Then
                n = n + 1
                For j = 1 To U_NCOLS
                    out(n, j) = Texto(vUsr(i, j))
                Next j
                out(n, U_LOGIN) = LCase$(out(n, U_LOGIN))
                out(n, U_GESTOR) = LCase$(out(n, U_GESTOR))
            End If
        Next i
    End If

    Desproteger ws
    EscreverTabela lo, out, n, U_NCOLS
    If n > 1 Then
        With lo.Sort
            .SortFields.Clear
            .SortFields.Add Key:=lo.ListColumns(U_NOME).DataBodyRange, Order:=xlAscending
            .Header = xlYes
            .Apply
        End With
    End If
    ValidacoesUsuarios lo
    ProtegerUsuarios
    ResetUsuarios
End Sub

Public Sub ValidacoesUsuarios(ByVal lo As ListObject)
    If lo.DataBodyRange Is Nothing Then Exit Sub
    ValidarLista lo.ListColumns(U_PERFIL).DataBodyRange, "lstPerfil", True
    ValidarLista lo.ListColumns(U_GESTOR).DataBodyRange, "lstColaboradores", False
    ValidarLista lo.ListColumns(U_DISC).DataBodyRange, "lstDisciplinas", False
End Sub

' Escreve uma matriz em uma tabela (ListObject), redimensionando-a
Public Sub EscreverTabela(ByVal lo As ListObject, ByRef dados As Variant, ByVal n As Long, ByVal nCols As Long)
    Dim ws As Worksheet, linhas As Long
    Set ws = lo.Parent
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.ClearContents
    linhas = n
    If linhas < 1 Then linhas = 1
    lo.Resize ws.Range(lo.HeaderRowRange.Cells(1, 1), lo.HeaderRowRange.Cells(1, nCols).Offset(linhas, 0))
    ' limpa sobras abaixo da tabela (quando ela diminui)
    ws.Range(lo.HeaderRowRange.Cells(1, 1).Offset(linhas + 1, 0), _
             lo.HeaderRowRange.Cells(1, nCols).Offset(linhas + 500, 0)).Clear
    If n > 0 Then lo.DataBodyRange.Resize(n, nCols).Value = dados
    ' reaplica o filtro que o usuário tiver aplicado na tabela
    On Error Resume Next
    If lo.ShowAutoFilter Then lo.AutoFilter.ApplyFilter
End Sub

Public Sub ValidarLista(ByVal rg As Range, ByVal nomeLista As String, ByVal restrito As Boolean)
    On Error Resume Next
    With rg.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=IIf(restrito, xlValidAlertStop, xlValidAlertInformation), _
             Operator:=xlBetween, Formula1:="=" & nomeLista
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
End Sub

'------------------------------------------------------------------------------
' Botões da aba Usuarios (somente Admin)
'------------------------------------------------------------------------------
Public Sub EditarUsuarios()
    If Not EhAdmin() Then
        MsgBox "Somente administradores podem alterar o cadastro de usuários.", vbExclamation, APP_NOME
        Exit Sub
    End If
    NR("cfgEditandoUsuarios").Value = "SIM"
    ProtegerUsuarios
    ThisWorkbook.Worksheets(SH_USR).Activate
    MsgBox "Modo de edição liberado." & vbLf & vbLf & _
           "- Inclua linhas logo abaixo da tabela (ela se expande sozinha)." & vbLf & _
           "- Login = usuário de rede do Windows (ex.: joao.silva)." & vbLf & _
           "- Gestor = login do chefe imediato." & vbLf & vbLf & _
           "Ao terminar, clique em 'Publicar usuários'.", vbInformation, APP_NOME
End Sub

Public Sub CancelarEdicaoUsuarios()
    NR("cfgEditandoUsuarios").Value = "NAO"
    CarregarDoServidor False
End Sub

Public Sub PublicarUsuarios()
    Dim lo As ListObject, v As Variant, out() As Variant, n As Long, i As Long, j As Long
    Dim vistos As Object, login As String, msg As String, temAdmin As Boolean, gest As String

    If Not EhAdmin() Then
        MsgBox "Somente administradores podem publicar o cadastro de usuários.", vbExclamation, APP_NOME
        Exit Sub
    End If
    Set lo = ThisWorkbook.Worksheets(SH_USR).ListObjects(TBL_USR)
    If lo.DataBodyRange Is Nothing Then Exit Sub
    v = lo.DataBodyRange.Value
    Set vistos = CreateObject("Scripting.Dictionary")
    vistos.CompareMode = vbTextCompare

    ReDim out(1 To UBound(v, 1), 1 To U_NCOLS)
    For i = 1 To UBound(v, 1)
        login = LCase$(Texto(v(i, U_LOGIN)))
        If Len(login) > 0 Then
            If vistos.Exists(login) Then msg = msg & "- Login repetido: " & login & vbLf
            vistos(login) = True
            n = n + 1
            For j = 1 To U_NCOLS
                out(n, j) = Texto(v(i, j))
            Next j
            out(n, U_LOGIN) = login
            out(n, U_GESTOR) = LCase$(out(n, U_GESTOR))
            If Len(out(n, U_PERFIL)) = 0 Then out(n, U_PERFIL) = PERFIL_COLAB
            If out(n, U_PERFIL) = PERFIL_ADMIN Then temAdmin = True
        End If
    Next i
    For i = 1 To n
        gest = out(i, U_GESTOR)
        If Len(gest) > 0 And Not vistos.Exists(gest) Then
            msg = msg & "- Gestor '" & gest & "' de '" & out(i, U_LOGIN) & "' não está cadastrado." & vbLf
        End If
        If gest = out(i, U_LOGIN) Then msg = msg & "- '" & gest & "' não pode ser gestor de si mesmo." & vbLf
    Next i
    If Not temAdmin Then msg = msg & "- É necessário pelo menos um usuário com perfil Admin." & vbLf
    If Len(msg) > 0 Then
        MsgBox "Corrija antes de publicar:" & vbLf & vbLf & msg, vbExclamation, APP_NOME
        Exit Sub
    End If

    If PublicarUsuariosServidor(out, n, msg) Then
        NR("cfgEditandoUsuarios").Value = "NAO"
        CarregarDoServidor True
        MsgBox n & " usuário(s) publicados no servidor.", vbInformation, APP_NOME
    Else
        MsgBox "Não foi possível publicar:" & vbLf & msg, vbExclamation, APP_NOME
    End If
End Sub
