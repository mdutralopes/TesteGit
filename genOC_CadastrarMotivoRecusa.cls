VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsgenOC_CadastrarMotivoRecusa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Declara as classes utilizadas
Private objMensagem                 As clsgenMensagem
Private objCadastrarMotivoRecusa    As ClsGenMotivoRecusa
Private objListaMotivoRecusa        As clsgenListaMotivoRecusa
Private objAlterarSenha             As ClsGenMotivoRecusa

' Declara os formulários usados nesta classe
Private objFormCadastrarMotivoRecusa    As frmgenCadastroMotivoRecusa

' Declara variáveis comuns
Private lID_MotivoRecusa            As Long
Private sTipoOperacao               As String
Private sDescricao                  As String
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

Public Sub TelaCadastrarMotivoRecusa()

On Error GoTo ErrorHandler:

    ' Instância o formulário de pré ativação e chama em seguida
    Set objFormCadastrarMotivoRecusa = New frmgenCadastroMotivoRecusa
    
    ' Intância a classe de mensagem
    Set objMensagem = New clsgenMensagem
    
    ' Informar o objeto controlador
    objFormCadastrarMotivoRecusa.InformaClasseOC Me
    
    ' Instância a classe de dados para cadastrar usuários
    Set objCadastrarMotivoRecusa = New ClsGenMotivoRecusa
    
    ' Informar o nome do banco desejado
    objCadastrarMotivoRecusa.InformaBanco frmgenMDI_SGAPlus.objBanco
            
    ' Informar o nome do objeto de mensagem
    objCadastrarMotivoRecusa.InformaMensagem objMensagem
    
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
                            "Houve um erro de execução no método: TelaCadastrarMotivoRecusas " & _
                            "da classe clsgenOC_CadastrarMotivoRecusa"
    GoTo Finalizar

End Sub

Public Sub ExibirLista()

On Error GoTo ErrorHandler:

    Dim bRetorno As Boolean
    
    lQuantidadeRegistros = 0
    
    Screen.MousePointer = vbArrowHourglass
    
    ' Intância a classe de mensagem
    Set objMensagem = New clsgenMensagem
    
    ' Informar o nome do objeto de mensagem
    objCadastrarMotivoRecusa.InformaMensagem objMensagem
    
    ' Intância a classe da lista de usuários
    Set objListaMotivoRecusa = Nothing
    Set objListaMotivoRecusa = New clsgenListaMotivoRecusa
    
    ' Limpa a lista
    objFormCadastrarMotivoRecusa.lstMotivoRecusa.Clear

    ' Executar o método de montagem da lista na classe de analisar order entry
    bRetorno = objCadastrarMotivoRecusa.MontarLista(objFormCadastrarMotivoRecusa.txtBusca.Text)
    
    ' Se ocorreu algum erro na DLL exibi-lo
    If Not bRetorno Then
        
        'Lista vazia, pesquisa nao trouxe registros
        objFormCadastrarMotivoRecusa.cmdAlterar.Enabled = False
        objFormCadastrarMotivoRecusa.cmdExcluir.Enabled = False
        
        LimparDadosTela
        
        GoTo Finalizar
        
    Else
        
        HabilitaDesabilitaObjetos
    
    End If
    
    ' Preenche a lista de usuários
    bRetorno = PreencheLista()
    
    If Not bRetorno Then
        
        GoTo Finalizar
    
    End If
    
    ' Posiciona o na primeira lista caso encontre algum registro
    objFormCadastrarMotivoRecusa.lstMotivoRecusa.ListIndex = 0
    
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
                            "da classe clswidOC_CadastrarMotivoRecusa"

    GoTo Finalizar

End Sub

Public Function GravarDados() As Boolean

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

        ' Criticar os campos da tela
        bRetorno = CriticarCamposTela
        
        If Not bRetorno Then
        
            GoTo Finalizar
            
        End If
    
        ' Informar o nome do objeto de mensagem
        objCadastrarMotivoRecusa.InformaMensagem objMensagem
    
        ' Atribui os campos da tela nas propriedades
        Call AtribuirCamposDaTelaNasPropriedades
        
        ' Excutar o método para gravar os dados da recusa
        bRetorno = objCadastrarMotivoRecusa.GravarDados(sTipoOperacao)
            
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
                            "da classe clsgenOC_CadastrarMotivoRecusa"

    GoTo Finalizar

End Function

Public Sub ExcluirDados()

On Error GoTo ErrorHandler:

    Dim bRetorno As Boolean

    ' Intância a classe de mensagem
    Set objMensagem = New clsgenMensagem
    
    objMensagem.TipoMsg = "Q"
    objMensagem.Descricao = "Confirma a exlusão do Motivo selecionado?"
    If objMensagem.ExibeMensagem = "S" Then

        ' Informar o nome do objeto de mensagem
        objCadastrarMotivoRecusa.InformaMensagem objMensagem
    
        ' Inicia a transação
        frmgenMDI_SGAPlus.objBanco.Iniciar_Transacao
    
           ' Excutar o método de exclusão dos dados do usuário
            bRetorno = objCadastrarMotivoRecusa.ExcluirDados(objCadastrarMotivoRecusa.ID_MOTIVO_RECUSA)
    
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
            objFormCadastrarMotivoRecusa.txtBusca.Text = ""
            
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
                            "da classe clsgenOC_CadastrarMotivoRecusa"

    Exit Sub

End Sub

Private Sub AtribuirCamposDaTelaNasPropriedades()

    objCadastrarMotivoRecusa.Descricao = TrataPlick(objFormCadastrarMotivoRecusa.txtDescricao)
    
End Sub

Private Sub Class_Terminate()

    Set objCadastrarMotivoRecusa = Nothing
    Set objMensagem = Nothing
    Set objListaMotivoRecusa = Nothing
    
End Sub

Private Function CriticarCamposTela() As Boolean

On Error GoTo ErrorHandler:

    CriticarCamposTela = False
    
    If Len(objFormCadastrarMotivoRecusa.txtDescricao.Text) = 0 Then
    
        objMensagem.TipoMsg = "A"
        objMensagem.Descricao = "Descrição é obrigatória."
        objFormCadastrarMotivoRecusa.txtDescricao.SetFocus
        
        GoTo Finalizar
    
    End If
    
    CriticarCamposTela = True

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
                            "Houve um erro de execução no método: CriticarCamposTela " & _
                            "da classe clsgenOC_CadastrarMotivoRecusa"
    
    GoTo Finalizar

End Function

Public Function SaiTela() As Boolean

On Error GoTo ErrorHandler:

    Dim bAlterou As Boolean

    SaiTela = False

    ' Chama o método para verificar se foram alterados alguns dados
    If AlterouDadosTela Then
       
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
                            "da classe clsgenOC_CadastrarMotivoRecusa"
    
    GoTo Finalizar

End Function

Public Function AlterouDadosTela() As Boolean

On Error GoTo ErrorHandler:

    AlterouDadosTela = True
    
    Set objMensagem = New clsgenMensagem
    
    If objCadastrarMotivoRecusa.Descricao <> objFormCadastrarMotivoRecusa.txtDescricao Then
    
        GoTo Finalizar
        
    End If

    AlterouDadosTela = False
    
Finalizar:

    On Error GoTo 0

    Exit Function
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: AlterouDadosTela " & _
                            "da classe clsgenOC_CadastrarMotivoRecusas"
    
    GoTo Finalizar

End Function

Private Function PreencheLista() As Boolean

    lQuantidadeRegistros = 0
    
    PreencheLista = False
    
    ' Ler o objeto de coleção da lista de motivos recusa
    For Each objListaMotivoRecusa In objCadastrarMotivoRecusa.ListaMotivoRecusa

        With objFormCadastrarMotivoRecusa

            ' Descricao Motivo Recusa
            .lstMotivoRecusa.AddItem (objListaMotivoRecusa.Descricao)
        
        End With

        lQuantidadeRegistros = lQuantidadeRegistros + 1
    
    Next
    
    PreencheLista = True
    
End Function

Public Sub RefreshDados()

    
    If lQuantidadeRegistros > 0 Then
    ' Atualiza os dados de usuário na tela
    With objFormCadastrarMotivoRecusa
    
        .txtDescricao = objCadastrarMotivoRecusa.ListaMotivoRecusa(objFormCadastrarMotivoRecusa.lstMotivoRecusa.ListIndex + 1).Descricao
        
        objCadastrarMotivoRecusa.ID_MOTIVO_RECUSA = objCadastrarMotivoRecusa.ListaMotivoRecusa(objFormCadastrarMotivoRecusa.lstMotivoRecusa.ListIndex + 1).ID_MOTIVO_RECUSA
        objCadastrarMotivoRecusa.Descricao = objCadastrarMotivoRecusa.ListaMotivoRecusa(objFormCadastrarMotivoRecusa.lstMotivoRecusa.ListIndex + 1).Descricao
        
    End With

    End If

End Sub

Private Sub HabilitaDesabilitaObjetos()

    ' Coloca cor de fundo como default para habilitação
    If (sTipoOperacao = "I") Or (sTipoOperacao = "A") Then ' Inclusão ou alteração
        
        ' Muda cor de fundo dos objetos de entrada de dados para branco
        objFormCadastrarMotivoRecusa.txtDescricao.BackColor = &H80000005
        
        ' Muda cor de escrita dos objetos de entrada de dados para preto
        objFormCadastrarMotivoRecusa.txtDescricao.ForeColor = &H80000008
        
        ' Habilita o frame dos objetos de dados
        objFormCadastrarMotivoRecusa.frameDados.Enabled = True
        
        ' Desabilita botões de inclusão, alteração, exlusão e frames de busca e lista
        objFormCadastrarMotivoRecusa.cmdIncluir.Enabled = False
        objFormCadastrarMotivoRecusa.cmdAlterar.Enabled = False
        objFormCadastrarMotivoRecusa.cmdExcluir.Enabled = False
        objFormCadastrarMotivoRecusa.frameBusca.Enabled = False
        objFormCadastrarMotivoRecusa.frameLista.Enabled = False
        
        ' Habilita os botões gravar, limpar e cancelar
        objFormCadastrarMotivoRecusa.cmdSalvar.Enabled = True
        objFormCadastrarMotivoRecusa.cmdLimpar.Enabled = True
        objFormCadastrarMotivoRecusa.cmdCancelar.Enabled = True
        
    Else
    
        ' Muda cor de fundo dos objetos de entrada de dados para cinza
        objFormCadastrarMotivoRecusa.txtDescricao.BackColor = &H80000004
        
        ' Muda cor de escrita dos objetos de entrada de dados azul
        objFormCadastrarMotivoRecusa.txtDescricao.ForeColor = &HC00000
        
        ' Desabilita o frame dos objetos de dados
        objFormCadastrarMotivoRecusa.frameDados.Enabled = False
    
        ' Habilita botões de inclusão, alteração, exlusão e frames de busca e lista
        objFormCadastrarMotivoRecusa.cmdIncluir.Enabled = True
        objFormCadastrarMotivoRecusa.cmdAlterar.Enabled = True
        objFormCadastrarMotivoRecusa.cmdExcluir.Enabled = True
        objFormCadastrarMotivoRecusa.frameBusca.Enabled = True
        objFormCadastrarMotivoRecusa.frameLista.Enabled = True
        
        ' Desabilita os botões gravar, limpar e cancelar
        objFormCadastrarMotivoRecusa.cmdSalvar.Enabled = False
        objFormCadastrarMotivoRecusa.cmdLimpar.Enabled = False
        objFormCadastrarMotivoRecusa.cmdCancelar.Enabled = False
    
    End If

End Sub

Public Sub LimparDadosTela()

    objFormCadastrarMotivoRecusa.txtDescricao.Text = ""
    objCadastrarMotivoRecusa.Descricao = ""
    
    Exit Sub
        
End Sub

Public Sub Saida()
    
    Unload objFormCadastrarMotivoRecusa
        
End Sub

