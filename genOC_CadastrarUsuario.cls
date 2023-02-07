VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsgenOC_CadastrarUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Declara as classes utilizadas
Private objMensagem                 As clsgenMensagem
Private objCadastrarUsuario         As clsgenUsuario
Private objListaUsuario             As clsgenListaUsuarios
Private objAlterarSenha             As clsgenUsuario

' Declara os formulários usados nesta classe
Private objFormCadastrarUsuario     As frmgenCadastroUsuarios
Private objFormAlterarSenha         As frmgenAlterarSenha

' Declara variáveis comuns
Private lID_USUARIO                 As Long
Private sTipoOperacao               As String
Private lQuantidadeRegistros        As Long

Public Property Get TipoOperacao() As String

    TipoOperacao = sTipoOperacao
    
End Property

Public Property Let TipoOperacao(ByVal sNewValue As String)

    sTipoOperacao = sNewValue
    
    ' Chama  método que controla a habilitação e desabilitação dos objetos,
    ' conforme o tipo de operação
    Call HabilitaDesabilitaObjetos
    
End Property

Public Sub TelaCadastrarUsuarios()

On Error GoTo ErrorHandler:

    ' Instância o formulário de pré ativação e chama em seguida
    Set objFormCadastrarUsuario = New frmgenCadastroUsuarios
    
    ' Intância a classe de mensagem
    Set objMensagem = New clsgenMensagem
    
    ' Informar o objeto controlador
    objFormCadastrarUsuario.InformaClasseOC Me
    
    ' Instância a classe de dados para cadastrar usuários
    Set objCadastrarUsuario = New clsgenUsuario
    
    ' Informar o nome do banco desejado
    objCadastrarUsuario.InformaBanco frmgenMDI_SGAPlus.objBanco
            
    ' Informar o nome do objeto de mensagem
    objCadastrarUsuario.InformaMensagem objMensagem
    
    ' Chama o método de exibição do GRID
    Call ExibirLista
    
Finalizar:

    Screen.MousePointer = vbDefault
    
    If objMensagem.TipoMsg <> "" Then
    
        objMensagem.ExibeMensagem
        
    End If
    
    Set objMensagem = Nothing
    
    On Error GoTo 0
    
    Exit Sub
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: TelaCadastrarUsuarios " & _
                            "da classe clsgenOC_CadastrarUsuario"
    GoTo Finalizar

End Sub

Public Sub TelaAlterarSenha()

    Dim bRetorno As Boolean

On Error GoTo ErrorHandler:

    ' Instância o formulário de pré ativação e chama em seguida
    Set objFormAlterarSenha = New frmgenAlterarSenha
    
    ' Intância a classe de mensagem
    Set objMensagem = New clsgenMensagem
    
    ' Informar o objeto controlador
    objFormAlterarSenha.InformaClasseOC Me
    
    ' Instância a classe de dados para cadastrar usuários
    Set objAlterarSenha = New clsgenUsuario
    
    ' Informar o nome do banco desejado
    objAlterarSenha.InformaBanco frmgenMDI_SGAPlus.objBanco
            
    ' Informar o nome do objeto de mensagem
    objAlterarSenha.InformaMensagem objMensagem
    
    ' Excutar o método para recuperar os dados de senha do usuário
    bRetorno = objAlterarSenha.RecuperarDados(GbID_Usuario)

    ' Se ocorreu algum erro na DLL
    If Not bRetorno Then

        GoTo Finalizar

    End If

    Call AtribuirCamposNaTelaSenha
    
Finalizar:

    Screen.MousePointer = vbDefault
    
    If objMensagem.TipoMsg <> "" Then
    
        objMensagem.ExibeMensagem
        
    End If
    
    Set objMensagem = Nothing
    
    On Error GoTo 0
    
    Exit Sub
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: TelaAlterarSenha " & _
                            "da classe clsgenOC_CadastrarUsuario"
    GoTo Finalizar

End Sub

Public Sub ExibirLista()

On Error GoTo ErrorHandler:

    Dim bRetorno As Boolean
    
    Screen.MousePointer = vbArrowHourglass
    
    ' Intância a classe de mensagem
    Set objMensagem = New clsgenMensagem
    
    ' Informar o nome do objeto de mensagem
    objCadastrarUsuario.InformaMensagem objMensagem
    
    ' Intância a classe da lista de usuários
    Set objListaUsuario = Nothing
    Set objListaUsuario = New clsgenListaUsuarios
    
    ' Limpa a lista
    objFormCadastrarUsuario.lstUsuarios.Clear

    ' Executar o método de montagem da lista na classe de analisar order entry
    bRetorno = objCadastrarUsuario.MontarLista(objFormCadastrarUsuario.optLogin.Value, _
                                               objFormCadastrarUsuario.txtBusca.Text)
    
    ' Se ocorreu algum erro na DLL exibi-lo
    If Not bRetorno Then
        
        GoTo Finalizar
    
    End If
    
    ' Preenche a lista de usuários
    bRetorno = PreencheLista()
    
    If Not bRetorno Then
        
        GoTo Finalizar
    
    End If
    
    ' Posiciona o na primeira lista caso encontre algum registro
    objFormCadastrarUsuario.lstUsuarios.ListIndex = 0
    
Finalizar:

    Screen.MousePointer = vbDefault
    
    If objMensagem.TipoMsg <> "" Then
    
        objMensagem.ExibeMensagem
        
    End If
    
    objMensagem.TipoMsg = ""
    
    On Error GoTo 0
    
    Exit Sub
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: ExibirLista " & _
                            "da classe clswidOC_CadastrarUsuario"

    GoTo Finalizar

End Sub

Public Function GravarDados(ByVal pyTela As Byte) As Boolean

On Error GoTo ErrorHandler:

    Dim bRetorno As Boolean

    GravarDados = False

    ' Intância a classe de mensagem
    Set objMensagem = New clsgenMensagem

    ' Inicia a transação
    bRetorno = frmgenMDI_SGAPlus.objBanco.Iniciar_Transacao()

        ' Se ocorreu algum erro na DLL
        If Not bRetorno Then

            GoTo Finalizar

        End If

        If pyTela = "1" Then  ' Cadastro de usuário
        
            ' Criticar os campos da tela
            bRetorno = CriticarCamposTelaUsuario
            
            If Not bRetorno Then
            
                GoTo Finalizar
                
            End If
        
            ' Informar o nome do objeto de mensagem
            objCadastrarUsuario.InformaMensagem objMensagem
        
            ' Atribui os campos da tela nas propriedades
            Call AtribuirCamposDaTelaUsuarioNasPropriedades
            
            ' Excutar o método para gravar os dados do usuário
            bRetorno = objCadastrarUsuario.GravarDados(sTipoOperacao)
            
        ElseIf pyTela = "2" Then  ' Alteração de senha
    
            ' Criticar os campos da tela
            bRetorno = CriticarCamposTelaSenha
            
            If Not bRetorno Then
            
                GoTo Finalizar
                
            End If
        
            ' Informar o nome do objeto de mensagem
            objAlterarSenha.InformaMensagem objMensagem
        
            ' Atribui os campos da tela nas propriedades
            Call AtribuirCamposDaTelaSenhaNasPropriedades
            
            ' Executar o método que verifica se a senha atual é válida
            bRetorno = objAlterarSenha.SenhaValida(GbID_Usuario, _
                                                   objFormAlterarSenha.txtSenhaAtual.Text)
            
            If Not bRetorno Then
            
                objFormAlterarSenha.txtSenhaAtual.SetFocus
                
                GoTo Finalizar
                
            End If
            
            ' Excutar o método para gravar a senha
            bRetorno = objAlterarSenha.GravarSenha(GbID_Usuario)
    
        End If

        ' Se ocorreu algum erro na DLL
        If Not bRetorno Then

            GoTo Finalizar

        End If

    ' Finaliza a transação
    frmgenMDI_SGAPlus.objBanco.Finalizar_Transacao
            
    ' Atualiza a mensagem de retorno como gravação realizada com sucesso
    objMensagem.TipoMsg = "A"
    objMensagem.Descricao = "Gravação realizada com sucesso"

    GravarDados = True

Finalizar:

    Screen.MousePointer = vbDefault

    If objMensagem.TipoMsg <> "" Then

        objMensagem.ExibeMensagem

    End If
    
    If Not GravarDados Then
    
        frmgenMDI_SGAPlus.objBanco.Cancelar_Transacao

    End If

    Set objMensagem = Nothing
    
    On Error GoTo 0

    Exit Function

ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: GravarDados " & _
                            "da classe clsgenOC_CadastrarUsuario"

    GoTo Finalizar

End Function

Public Sub ExcluirDados()

On Error GoTo ErrorHandler:

    Dim bRetorno As Boolean

    ' Intância a classe de mensagem
    Set objMensagem = New clsgenMensagem
    
    objMensagem.TipoMsg = "Q"
    objMensagem.Descricao = "Confirma a exlusão do usuário selecionado?"
    
    If objMensagem.ExibeMensagem = "S" Then

        ' Informar o nome do objeto de mensagem
        objCadastrarUsuario.InformaMensagem objMensagem
    
        ' Inicia a transação
        frmgenMDI_SGAPlus.objBanco.Iniciar_Transacao
    
           ' Excutar o método de exclusão dos dados do usuário
            bRetorno = objCadastrarUsuario.ExcluirDados(objCadastrarUsuario.ID_USUARIO)
    
            ' Se ocorreu algum erro na DLL
            If Not bRetorno Then
    
                frmgenMDI_SGAPlus.objBanco.Cancelar_Transacao
    
                GoTo Finalizar
    
            End If
    
        ' Finaliza a transação
        frmgenMDI_SGAPlus.objBanco.Finalizar_Transacao
        
        ' Se existia somente um registro limpar o objeto de entrada para buscar
        If lQuantidadeRegistros = 1 Then
        
            ' Limpa o campo de busca
            objFormCadastrarUsuario.txtBusca.Text = ""
            
        End If
        
        ' Chama o método de exibição do GRID
        Call ExibirLista
        
    Else
    
        objMensagem.TipoMsg = ""
        
    End If
    
Finalizar:

    Screen.MousePointer = vbDefault

    If objMensagem.TipoMsg <> "" Then

        objMensagem.ExibeMensagem

    End If

    Set objMensagem = Nothing
    
    On Error GoTo 0

    Exit Sub

ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: ExcluirDados " & _
                            "da classe clsgenOC_CadastrarUsuario"

    Exit Sub

End Sub

Public Sub ResetarSenha()

On Error GoTo ErrorHandler:

    Dim bRetorno As Boolean

    ' Intância a classe de mensagem
    Set objMensagem = New clsgenMensagem
    
    objMensagem.TipoMsg = "Q"
    objMensagem.Descricao = "Confirma o reset de senha do usuário selecionado?"
    
    If objMensagem.ExibeMensagem = "S" Then

        ' Informar o nome do objeto de mensagem
        objCadastrarUsuario.InformaMensagem objMensagem
    
        ' Inicia a transação
        frmgenMDI_SGAPlus.objBanco.Iniciar_Transacao
    
       ' Excutar o método de exclusão dos dados do usuário
        bRetorno = objCadastrarUsuario.ResetSenha(objCadastrarUsuario.ID_USUARIO)

        ' Se ocorreu algum erro na DLL
        If Not bRetorno Then

            frmgenMDI_SGAPlus.objBanco.Cancelar_Transacao

            GoTo Finalizar

        End If

        ' Finaliza a transação
        frmgenMDI_SGAPlus.objBanco.Finalizar_Transacao
        
        ' Se existia somente um registro limpar o objeto de entrada para buscar
        If lQuantidadeRegistros = 1 Then
        
            ' Limpa o campo de busca
            objFormCadastrarUsuario.txtBusca.Text = ""
            
        End If
        
    Else
    
        objMensagem.TipoMsg = ""
        
    End If
    
Finalizar:

    Screen.MousePointer = vbDefault

    If objMensagem.TipoMsg <> "" Then

        objMensagem.ExibeMensagem

    End If

    Set objMensagem = Nothing
    
    On Error GoTo 0

    Exit Sub

ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: ExcluirDados " & _
                            "da classe clsgenOC_CadastrarUsuario"

    Exit Sub

End Sub

Private Sub AtribuirCamposDaTelaUsuarioNasPropriedades()

    objCadastrarUsuario.Login = TrataPlick(objFormCadastrarUsuario.txtLogin)
    objCadastrarUsuario.Nome = TrataPlick(objFormCadastrarUsuario.TxtNome)
    objCadastrarUsuario.EMail = TrataPlick(objFormCadastrarUsuario.txtMail)
    objCadastrarUsuario.Ramal = TrataPlick(objFormCadastrarUsuario.txtRamal)
    objCadastrarUsuario.Matricula = TrataPlick(objFormCadastrarUsuario.txtMatricula)
    
End Sub

Private Sub Class_Terminate()

    Set objCadastrarUsuario = Nothing
    Set objMensagem = Nothing
    Set objListaUsuario = Nothing
    
End Sub

Private Function CriticarCamposTelaUsuario() As Boolean

On Error GoTo ErrorHandler:

    CriticarCamposTelaUsuario = False
    
    If Len(objFormCadastrarUsuario.txtLogin.Text) = 0 Then
    
        objMensagem.TipoMsg = "A"
        objMensagem.Descricao = "Login do usuário obrigatório."
        objFormCadastrarUsuario.txtLogin.SetFocus
        
        GoTo Finalizar
    
    ElseIf Len(objFormCadastrarUsuario.TxtNome.Text) = 0 Then
    
        objMensagem.TipoMsg = "A"
        objMensagem.Descricao = "Nome do usuário obrigatório."
        objFormCadastrarUsuario.TxtNome.SetFocus
    
        GoTo Finalizar
    
    ElseIf Len(objFormCadastrarUsuario.txtMatricula.Text) = 0 Then
    
        objMensagem.TipoMsg = "A"
        objMensagem.Descricao = "Matrícula do usuário obrigatório."
        objFormCadastrarUsuario.txtMatricula.SetFocus
    
        GoTo Finalizar
    
    End If
    
    CriticarCamposTelaUsuario = True

Finalizar:

    If objMensagem.TipoMsg <> "" Then
    
        Screen.MousePointer = vbDefault
        
        objMensagem.ExibeMensagem
        
    End If
    
    On Error GoTo 0
    
    Exit Function
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: CriticarCamposTelaUsuario " & _
                            "da classe clsgenOC_CadastrarUsuario"
    
    GoTo Finalizar

End Function

Public Sub Usuario_CF()
'
'Prepara objetos para chamada da tela Cadastro_Usuario_CF
'Passa parametro de origem da chamada
'Tela Usuarios parametro 1
'Tela CF parametro 2

Dim objCadastrarUsuario_CF  As clsgenOC_CadastrarUsuario_CF

On Error GoTo ErrorHandler

    Set objCadastrarUsuario_CF = New clsgenOC_CadastrarUsuario_CF
    
    Set objMensagem = New clsgenMensagem
    
    objCadastrarUsuario_CF.TelaCadastrarUsuarios_CF 1, objCadastrarUsuario.Login, objCadastrarUsuario.ID_USUARIO
    
    Set objCadastrarUsuario_CF = Nothing

Finalizar:

    Screen.MousePointer = vbDefault
    
    If objMensagem.TipoMsg <> "" Then
    
        objMensagem.ExibeMensagem
        
    End If
    
    Set objMensagem = Nothing
    
    On Error GoTo 0
    
    Exit Sub
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método:  Usuario_CF" & _
                            "da classe clsgenOC_CadastrarUsuario"
    GoTo Finalizar

End Sub

Public Function SaiTela(ByVal pyTela As Byte) As Boolean

On Error GoTo ErrorHandler:

    Dim bAlterou As Boolean

    SaiTela = False

    If pyTela = "1" Then  ' Cadastro de usuário
    
        bAlterou = AlterouDadosTelaUsuario
    
    ElseIf pyTela = "2" Then  ' Alteração de senha

        bAlterou = AlterouDadosTelaSenha

    End If

    ' Chama o método para verificar se foram alterados alguns dados
    If bAlterou Then
       
       objMensagem.TipoMsg = "Q"
       objMensagem.Descricao = "Foram feitas alterações nos dados da tela que " & _
                               "ainda não foram Salvas! Deseja sair assim mesmo ?"
       
       
    End If
    
    SaiTela = True
    
Finalizar:

    If objMensagem.TipoMsg <> "" Then
    
        Screen.MousePointer = vbDefault
    
        SaiTela = objMensagem.ExibeMensagem = "S"
    
    End If
    
    On Error GoTo 0
    
    Exit Function
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: SaiTela " & _
                            "da classe clsgenOC_CadastrarUsuarios"
    
    GoTo Finalizar

End Function

Public Function AlterouDadosTelaUsuario() As Boolean

On Error GoTo ErrorHandler:

    AlterouDadosTelaUsuario = True
    
    Set objMensagem = New clsgenMensagem
    
    If objCadastrarUsuario.Login <> objFormCadastrarUsuario.txtLogin Then
    
        GoTo Finalizar
        
    End If

    If objCadastrarUsuario.Nome <> objFormCadastrarUsuario.TxtNome Then
    
        GoTo Finalizar
        
    End If
    
    If objCadastrarUsuario.EMail <> objFormCadastrarUsuario.txtMail Then
    
        GoTo Finalizar
        
    End If
    
    If objCadastrarUsuario.Ramal <> objFormCadastrarUsuario.txtRamal Then
    
        GoTo Finalizar
        
    End If
    
    If objCadastrarUsuario.Matricula <> objFormCadastrarUsuario.txtMatricula Then
    
        GoTo Finalizar
        
    End If
    
    AlterouDadosTelaUsuario = False
    
Finalizar:

    On Error GoTo 0

    Exit Function
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: AlterouDadosTelaUsuario " & _
                            "da classe clsgenOC_CadastrarUsuarios"
    
    GoTo Finalizar

End Function

Private Function PreencheLista() As Boolean

    lQuantidadeRegistros = 0
    
    PreencheLista = False
    
    ' Ler o objeto de coleção da lista de usuários
    For Each objListaUsuario In objCadastrarUsuario.ListaUsuario

        With objFormCadastrarUsuario

            If .optLogin.Value = True Then

                ' Login do usuário
                .lstUsuarios.AddItem (objListaUsuario.Login)

            ElseIf objFormCadastrarUsuario.optNome.Value = True Then

                ' Nome do usuário
                .lstUsuarios.AddItem (objListaUsuario.Nome)

            End If
            
        End With

        lQuantidadeRegistros = lQuantidadeRegistros + 1
    
    Next
    
    PreencheLista = True
    
End Function

Public Sub RefreshDados()
    
    Dim objCadastrarUsuario_CF As clsgenOC_CadastrarUsuario_CF
    
    Set objCadastrarUsuario_CF = New clsgenOC_CadastrarUsuario_CF
    
    ' Atualiza os dados de usuário na tela
    With objFormCadastrarUsuario
        .txtLogin = objCadastrarUsuario.ListaUsuario(objFormCadastrarUsuario.lstUsuarios.ListIndex + 1).Login
        .TxtNome = objCadastrarUsuario.ListaUsuario(objFormCadastrarUsuario.lstUsuarios.ListIndex + 1).Nome
        .txtMail = objCadastrarUsuario.ListaUsuario(objFormCadastrarUsuario.lstUsuarios.ListIndex + 1).Mail
        .txtRamal = objCadastrarUsuario.ListaUsuario(objFormCadastrarUsuario.lstUsuarios.ListIndex + 1).Ramal
        .txtMatricula = objCadastrarUsuario.ListaUsuario(objFormCadastrarUsuario.lstUsuarios.ListIndex + 1).Matricula
    
        objCadastrarUsuario.ID_USUARIO = objCadastrarUsuario.ListaUsuario(objFormCadastrarUsuario.lstUsuarios.ListIndex + 1).ID_USUARIO
        objCadastrarUsuario.Login = objCadastrarUsuario.ListaUsuario(objFormCadastrarUsuario.lstUsuarios.ListIndex + 1).Login
        objCadastrarUsuario.Nome = objCadastrarUsuario.ListaUsuario(objFormCadastrarUsuario.lstUsuarios.ListIndex + 1).Nome
        objCadastrarUsuario.EMail = objCadastrarUsuario.ListaUsuario(objFormCadastrarUsuario.lstUsuarios.ListIndex + 1).Mail
        objCadastrarUsuario.Ramal = objCadastrarUsuario.ListaUsuario(objFormCadastrarUsuario.lstUsuarios.ListIndex + 1).Ramal
        objCadastrarUsuario.Matricula = objCadastrarUsuario.ListaUsuario(objFormCadastrarUsuario.lstUsuarios.ListIndex + 1).Matricula
    
        'Guarda login usuario selecionado na lista para passar para classe de cadastro Usu X CF
        objCadastrarUsuario_CF.Login_CF = objCadastrarUsuario.Login
        objCadastrarUsuario_CF.ID_Login_CF = objCadastrarUsuario.ID_USUARIO
    
   
    End With

    Set objCadastrarUsuario_CF = Nothing
    
End Sub

Private Sub HabilitaDesabilitaObjetos()

    ' Coloca cor de fundo como default para habilitação
    If (sTipoOperacao = "I") Or (sTipoOperacao = "A") Then ' Inclusão ou alteração
        
        ' Muda cor de fundo dos objetos de entrada de dados para branco
        objFormCadastrarUsuario.txtLogin.BackColor = &H80000005
        objFormCadastrarUsuario.TxtNome.BackColor = &H80000005
        objFormCadastrarUsuario.txtMail.BackColor = &H80000005
        objFormCadastrarUsuario.txtRamal.BackColor = &H80000005
        objFormCadastrarUsuario.txtMatricula.BackColor = &H80000005
        
        ' Muda cor de escrita dos objetos de entrada de dados para preto
        objFormCadastrarUsuario.txtLogin.ForeColor = &H80000008
        objFormCadastrarUsuario.TxtNome.ForeColor = &H80000008
        objFormCadastrarUsuario.txtMail.ForeColor = &H80000008
        objFormCadastrarUsuario.txtRamal.ForeColor = &H80000008
        objFormCadastrarUsuario.txtMatricula.ForeColor = &H80000008
        
        ' Habilita o frame dos objetos de dados
        objFormCadastrarUsuario.frameDados.Enabled = True
        
        ' Desabilita botões de inclusão, alteração, exlusão e frames de busca e lista
        objFormCadastrarUsuario.cmdIncluir.Enabled = False
        objFormCadastrarUsuario.cmdAlterar.Enabled = False
        objFormCadastrarUsuario.cmdExcluir.Enabled = False
        objFormCadastrarUsuario.frameBusca.Enabled = False
        objFormCadastrarUsuario.frameLista.Enabled = False
        objFormCadastrarUsuario.cmdreset.Enabled = False
        
        ' Habilita os botões gravar, limpar e cancelar
        objFormCadastrarUsuario.cmdSalvar.Enabled = True
        objFormCadastrarUsuario.cmdLimpar.Enabled = True
        objFormCadastrarUsuario.cmdCancelar.Enabled = True
        
    Else
    
        ' Muda cor de fundo dos objetos de entrada de dados para cinza
        objFormCadastrarUsuario.txtLogin.BackColor = &H80000004
        objFormCadastrarUsuario.TxtNome.BackColor = &H80000004
        objFormCadastrarUsuario.txtMail.BackColor = &H80000004
        objFormCadastrarUsuario.txtRamal.BackColor = &H80000004
        objFormCadastrarUsuario.txtMatricula.BackColor = &H80000004
    
        ' Muda cor de escrita dos objetos de entrada de dados azul
        objFormCadastrarUsuario.txtLogin.ForeColor = &HC00000
        objFormCadastrarUsuario.TxtNome.ForeColor = &HC00000
        objFormCadastrarUsuario.txtMail.ForeColor = &HC00000
        objFormCadastrarUsuario.txtRamal.ForeColor = &HC00000
        objFormCadastrarUsuario.txtMatricula.ForeColor = &HC00000
        
        ' Desabilita o frame dos objetos de dados
        objFormCadastrarUsuario.frameDados.Enabled = False
    
        ' Habilita botões de inclusão, alteração, exlusão e frames de busca e lista
        objFormCadastrarUsuario.cmdIncluir.Enabled = True
        objFormCadastrarUsuario.cmdAlterar.Enabled = True
        objFormCadastrarUsuario.cmdExcluir.Enabled = True
        objFormCadastrarUsuario.frameBusca.Enabled = True
        objFormCadastrarUsuario.frameLista.Enabled = True
        objFormCadastrarUsuario.cmdreset.Enabled = True
        
        ' Desabilita os botões gravar, limpar e cancelar
        objFormCadastrarUsuario.cmdSalvar.Enabled = False
        objFormCadastrarUsuario.cmdLimpar.Enabled = False
        objFormCadastrarUsuario.cmdCancelar.Enabled = False
    
    End If

End Sub

Public Sub LimparDadosTelaUsuario()

    objFormCadastrarUsuario.txtLogin.Text = ""
    objFormCadastrarUsuario.TxtNome.Text = ""
    objFormCadastrarUsuario.txtMail.Text = ""
    objFormCadastrarUsuario.txtRamal.Text = ""
    objFormCadastrarUsuario.txtMatricula.Text = ""
    
    Exit Sub
        
End Sub

Public Sub Saida(ByVal pyTela As Byte)
    
    If pyTela = "1" Then  ' Cadastro de usuário
    
        Unload objFormCadastrarUsuario
        
    ElseIf pyTela = "2" Then  ' Alteração de senha

        Unload objFormAlterarSenha

    End If

End Sub

Public Sub AtribuirCamposNaTelaSenha()

    ' Atribui as propriedades nos objetos da tela
    objFormAlterarSenha.lblLogin.Caption = objAlterarSenha.Login
    objFormAlterarSenha.txtSenhaAtual = ""
    objFormAlterarSenha.txtNovaSenha.Text = objAlterarSenha.NovaSenha
    objFormAlterarSenha.txtConfirmaSenha.Text = objAlterarSenha.ConfirmaSenha

End Sub

Public Sub LimparDadosTelaSenha()

    objFormAlterarSenha.txtSenhaAtual.Text = ""
    objFormAlterarSenha.txtNovaSenha.Text = ""
    objFormAlterarSenha.txtConfirmaSenha.Text = ""
    
    Exit Sub
        
End Sub

Public Function AlterouDadosTelaSenha() As Boolean

On Error GoTo ErrorHandler:

    AlterouDadosTelaSenha = True
    
    Set objMensagem = New clsgenMensagem
    
    If objAlterarSenha.NovaSenha <> objFormAlterarSenha.txtNovaSenha Then
    
        GoTo Finalizar
        
    End If
    
    If objAlterarSenha.ConfirmaSenha <> objFormAlterarSenha.txtConfirmaSenha Then
    
        GoTo Finalizar
        
    End If
    
    AlterouDadosTelaSenha = False
    
Finalizar:

    On Error GoTo 0

    Exit Function
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: AlterouDadosTelaSenha " & _
                            "da classe clsgenOC_CadastrarUsuarios"
    
    GoTo Finalizar

End Function

Private Function CriticarCamposTelaSenha() As Boolean

On Error GoTo ErrorHandler:

    CriticarCamposTelaSenha = False
    
    If Len(objFormAlterarSenha.txtSenhaAtual.Text) = 0 Then
    
        objMensagem.TipoMsg = "A"
        objMensagem.Descricao = "Senha atual não informada."
        objFormAlterarSenha.txtSenhaAtual.SetFocus
        
        GoTo Finalizar
    
    ElseIf objFormAlterarSenha.txtNovaSenha.Text <> objFormAlterarSenha.txtConfirmaSenha.Text Then
    
        objMensagem.TipoMsg = "A"
        objMensagem.Descricao = "A confirmação da senha diferente da nova senha."
        objFormAlterarSenha.txtConfirmaSenha.SetFocus
        
        GoTo Finalizar
    
    End If
    
    CriticarCamposTelaSenha = True

Finalizar:

    If objMensagem.TipoMsg <> "" Then
    
        Screen.MousePointer = vbDefault
        
        objMensagem.ExibeMensagem
        
    End If
    
    On Error GoTo 0
    
    Exit Function
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: CriticarCamposTelaSenha " & _
                            "da classe clsgenOC_CadastrarUsuario"
    
    GoTo Finalizar

End Function

Private Sub AtribuirCamposDaTelaSenhaNasPropriedades()

    objAlterarSenha.Senha = TrataPlick(objFormAlterarSenha.txtSenhaAtual)
    objAlterarSenha.NovaSenha = TrataPlick(objFormAlterarSenha.txtNovaSenha)
    objAlterarSenha.ConfirmaSenha = TrataPlick(objFormAlterarSenha.txtConfirmaSenha)
    
End Sub


