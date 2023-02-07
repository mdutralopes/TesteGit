VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsgenOC_CadastrarTipoInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Declara as classes utilizadas
Private objMensagem                     As clsgenMensagem
Private objCadastrarTipoInterface       As clsgenTipoInterface
Private objListaTipoInterface           As clsgenListaTipoInterface

' Declara os formulários usados nesta classe
Private objFormCadastrarTipoInterface As frmgenCadastroTipoInterface

'Private clsgenOC_CadastrarTipoInterface   As frmgenCadastroTipoInterface

' Declara variáveis comuns
Private lID              As Long
Private sTipoOperacao                   As String
Private lQuantidadeRegistros            As Long

Public Property Get TipoOperacao() As String

    TipoOperacao = sTipoOperacao
    
End Property

Public Property Let TipoOperacao(ByVal sNewValue As String)

    sTipoOperacao = sNewValue
    
    ' Chama  método que controla a habilitação e desabilitação dos objetos,
    ' conforme o tipo de operação
    Call HabilitaDesabilitaObjetos
    
End Property

Public Sub TelaCadastrarTipoInterfaces()

On Error GoTo ErrorHandler:

    ' Instância o formulário de pré ativação e chama em seguida
    Set objFormCadastrarTipoInterface = New frmgenCadastroTipoInterface
    
    ' Intância a classe de mensagem
    Set objMensagem = New clsgenMensagem
    
    ' Informar o objeto controlador
    objFormCadastrarTipoInterface.InformaClasseOC Me
    
    ' Instância a classe de dados para cadastrar usuários
    Set objCadastrarTipoInterface = New clsgenTipoInterface
    
    ' Informar o nome do banco desejado
    objCadastrarTipoInterface.InformaBanco frmgenMDI_SGAPlus.objBanco
            
    ' Informar o nome do objeto de mensagem
    objCadastrarTipoInterface.InformaMensagem objMensagem
    
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
                            "Houve um erro de execução no método: TelaCadastrarTipoInterfaces " & _
                            "da classe clsgenOC_CadastrarTipoInterface"
    GoTo Finalizar

End Sub

Public Sub ExibirLista()

On Error GoTo ErrorHandler:

    Dim bRetorno As Boolean
    
    Screen.MousePointer = vbArrowHourglass
    
    lQuantidadeRegistros = 0
    
    ' Intância a classe de mensagem
    Set objMensagem = New clsgenMensagem
    
    ' Informar o nome do objeto de mensagem
    objCadastrarTipoInterface.InformaMensagem objMensagem
    
    ' Intância a classe da lista de usuários
    Set objListaTipoInterface = Nothing
    Set objListaTipoInterface = New clsgenListaTipoInterface
    
    ' Limpa a lista
    objFormCadastrarTipoInterface.lstTipoInterface.Clear
    
    ' Executar o método de montagem da lista na classe de analisar order entry
    bRetorno = objCadastrarTipoInterface.MontarLista(objFormCadastrarTipoInterface.optConsultaEmbratel.Value, _
                                                     objFormCadastrarTipoInterface.optConsultaCliente.Value, _
                                                     objFormCadastrarTipoInterface.txtBusca.Text)
    
    
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
    objFormCadastrarTipoInterface.lstTipoInterface.ListIndex = 0
    
Finalizar:

    ' Se não encontrou registro desabilitar os botôes de alteração e exclusão
    If lQuantidadeRegistros = 0 Then
    
        objFormCadastrarTipoInterface.cmdAlterar.Enabled = False
        objFormCadastrarTipoInterface.cmdExcluir.Enabled = False
        
        Call LimparDadosTelaTipoInterface
        
    Else
    
        objFormCadastrarTipoInterface.cmdAlterar.Enabled = True
        objFormCadastrarTipoInterface.cmdExcluir.Enabled = True
        
    End If
    
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
                            "da classe clswidOC_CadastrarTipoInterface"

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

        ' Criticar os campos da tela
        bRetorno = CriticarCamposTelaTipoInterface
            
        If Not bRetorno Then
        
            GoTo Finalizar
            
        End If
    
        ' Informar o nome do objeto de mensagem
        objCadastrarTipoInterface.InformaMensagem objMensagem
        
        ' Atribui os campos da tela nas propriedades
        Call AtribuirCamposNasPropriedades
            
        ' Excutar o método para gravar os dados do usuário
        bRetorno = objCadastrarTipoInterface.GravarDados(sTipoOperacao)
            
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
                            "da classe clsgenOC_CadastrarTipoInterface"

    GoTo Finalizar

End Function

Public Sub ExcluirDados()

On Error GoTo ErrorHandler:

    Dim bRetorno As Boolean

    ' Intância a classe de mensagem
    Set objMensagem = New clsgenMensagem
    
    objMensagem.TipoMsg = "Q"
    objMensagem.Descricao = "Confirma a exlusão do Tipo de Interface selecionado?"
    If objMensagem.ExibeMensagem = "S" Then

        ' Informar o nome do objeto de mensagem
        objCadastrarTipoInterface.InformaMensagem objMensagem
    
        ' Inicia a transação
        frmgenMDI_SGAPlus.objBanco.Iniciar_Transacao
    
           ' Excutar o método de exclusão dos dados do usuário
            bRetorno = objCadastrarTipoInterface.ExcluirDados(objCadastrarTipoInterface.ID)
    
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
            objFormCadastrarTipoInterface.txtBusca.Text = ""
            
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
                            "da classe clsgenOC_CadastrarTipoInterface"

    Exit Sub

End Sub

Private Sub AtribuirCamposNasPropriedades()

    objCadastrarTipoInterface.Descricao = TrataPlick(objFormCadastrarTipoInterface.TxtDescricao)
'    objCadastrarTipoInterface.Tipo_Interface =
    If objFormCadastrarTipoInterface.OptEmbratel.Value = True Then
    
        objCadastrarTipoInterface.Tipo_Interface = "E"
    
    ElseIf objFormCadastrarTipoInterface.optCliente.Value = True Then
    
        objCadastrarTipoInterface.Tipo_Interface = "C"
    
    End If

End Sub

Private Sub Class_Terminate()

    Set objCadastrarTipoInterface = Nothing
    Set objMensagem = Nothing
    Set objListaTipoInterface = Nothing
    
End Sub

Private Function CriticarCamposTelaTipoInterface() As Boolean

On Error GoTo ErrorHandler:

    CriticarCamposTelaTipoInterface = False
    
'    If Len(clsgenOC_CadastrarTipoInterface.txtCodigo_Controle.Text) = 0 Then
''        objMensagem.TipoMsg = "A"
''        objMensagem.Descricao = "Código de Controle é obrigatória."
''        clsgenOC_CadastrarTipoInterface.txtCodigo_Controle.SetFocus
''
'        GoTo Finalizar
'    Else
    
    If Len(objFormCadastrarTipoInterface.TxtDescricao.Text) = 0 Then
        objMensagem.TipoMsg = "A"
        objMensagem.Descricao = "Descrição é obrigatória."
        objFormCadastrarTipoInterface.TxtDescricao.SetFocus
        
        GoTo Finalizar
    End If
    
    If ((objFormCadastrarTipoInterface.OptEmbratel.Value = False) And (objFormCadastrarTipoInterface.optCliente.Value = False)) Then
        objMensagem.TipoMsg = "A"
        objMensagem.Descricao = "Tipo de Interface é obrigatória."
        objFormCadastrarTipoInterface.TxtDescricao.SetFocus
        
        GoTo Finalizar
    End If
    
    CriticarCamposTelaTipoInterface = True

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
                            "Houve um erro de execução no método: CriticarCamposTelaTipoInterface " & _
                            "da classe clsgenOC_CadastrarTipoInterface"
    
    GoTo Finalizar

End Function

Public Function SaiTela(ByVal pyTela As Byte) As Boolean

On Error GoTo ErrorHandler:

    Dim bAlterou As Boolean

    SaiTela = False

    bAlterou = AlterouDadosTelaTipoInterface
    
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
                            "da classe clsgenOC_CadastrarTipoInterfaces"
    
    GoTo Finalizar

End Function

Public Function AlterouDadosTelaTipoInterface() As Boolean

On Error GoTo ErrorHandler:

    AlterouDadosTelaTipoInterface = True
    
    Set objMensagem = New clsgenMensagem
    
    If objCadastrarTipoInterface.Descricao <> objFormCadastrarTipoInterface.TxtDescricao Then
    
        GoTo Finalizar
        
    End If

    AlterouDadosTelaTipoInterface = False
    
Finalizar:

    On Error GoTo 0

    Exit Function
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: AlterouDadosTelaTipoInterface " & _
                            "da classe clsgenOC_CadastrarTipoInterfaces"
    
    GoTo Finalizar

End Function

Private Function PreencheLista() As Boolean

    lQuantidadeRegistros = 0
    
    PreencheLista = False
    
    ' Ler o objeto de coleção da lista de tipos de pendências
    For Each objListaTipoInterface In objCadastrarTipoInterface.ListaTipoInterface

        With objFormCadastrarTipoInterface
            
            ' Descricao Motivo Recusa
             .lstTipoInterface.AddItem (objListaTipoInterface.Descricao)
        
        End With

        lQuantidadeRegistros = lQuantidadeRegistros + 1
    
    Next
    
    PreencheLista = True
    
End Function

Public Sub RefreshDados()

    If lQuantidadeRegistros > 0 Then
    ' Atualiza os dados de usuário na tela
    With objFormCadastrarTipoInterface
        
        .TxtDescricao = objCadastrarTipoInterface.ListaTipoInterface(objFormCadastrarTipoInterface.lstTipoInterface.ListIndex + 1).Descricao
        
        If (objCadastrarTipoInterface.ListaTipoInterface(objFormCadastrarTipoInterface.lstTipoInterface.ListIndex + 1).Tipo_Interface) = "E" Then
        
        .OptEmbratel.Value = True
        
        ElseIf (objCadastrarTipoInterface.ListaTipoInterface(objFormCadastrarTipoInterface.lstTipoInterface.ListIndex + 1).Tipo_Interface) = "C" Then
        
            .optCliente.Value = True
        
        End If
        
    End With
    
        objCadastrarTipoInterface.Descricao = objCadastrarTipoInterface.ListaTipoInterface(objFormCadastrarTipoInterface.lstTipoInterface.ListIndex + 1).Descricao
        objCadastrarTipoInterface.Tipo_Interface = objCadastrarTipoInterface.ListaTipoInterface(objFormCadastrarTipoInterface.lstTipoInterface.ListIndex + 1).Tipo_Interface
        objCadastrarTipoInterface.ID = objCadastrarTipoInterface.ListaTipoInterface(objFormCadastrarTipoInterface.lstTipoInterface.ListIndex + 1).ID
    End If


End Sub

Private Sub HabilitaDesabilitaObjetos()

    ' Coloca cor de fundo como default para habilitação
    If (sTipoOperacao = "I") Or (sTipoOperacao = "A") Then ' Inclusão ou alteração
        
        ' Muda cor de fundo dos objetos de entrada de dados para branco
        objFormCadastrarTipoInterface.TxtDescricao.BackColor = &H80000005

        
        ' Muda cor de escrita dos objetos de entrada de dados para preto
        objFormCadastrarTipoInterface.TxtDescricao.ForeColor = &H80000008

        
        ' Habilita o frame dos objetos de dados
        objFormCadastrarTipoInterface.frameDados.Enabled = True
        
        ' Desabilita botões de inclusão, alteração, exlusão e frames de busca e lista
        objFormCadastrarTipoInterface.cmdIncluir.Enabled = False
        objFormCadastrarTipoInterface.cmdAlterar.Enabled = False
        objFormCadastrarTipoInterface.cmdExcluir.Enabled = False
        objFormCadastrarTipoInterface.frameBusca.Enabled = False
        objFormCadastrarTipoInterface.frameLista.Enabled = False
        
        ' Habilita os botões gravar, limpar e cancelar
        objFormCadastrarTipoInterface.cmdSalvar.Enabled = True
        objFormCadastrarTipoInterface.cmdLimpar.Enabled = True
        objFormCadastrarTipoInterface.cmdCancelar.Enabled = True
        
    Else
    
        ' Muda cor de fundo dos objetos de entrada de dados para cinza
        objFormCadastrarTipoInterface.TxtDescricao.BackColor = &H80000004

        
        ' Muda cor de escrita dos objetos de entrada de dados azul
        objFormCadastrarTipoInterface.TxtDescricao.ForeColor = &HC00000
        
        ' Desabilita o frame dos objetos de dados
        objFormCadastrarTipoInterface.frameDados.Enabled = False
    
        ' Habilita botões de inclusão, alteração, exlusão e frames de busca e lista
        objFormCadastrarTipoInterface.cmdIncluir.Enabled = True
        objFormCadastrarTipoInterface.cmdAlterar.Enabled = True
        objFormCadastrarTipoInterface.cmdExcluir.Enabled = True
        objFormCadastrarTipoInterface.frameBusca.Enabled = True
        objFormCadastrarTipoInterface.frameLista.Enabled = True
        
        ' Desabilita os botões gravar, limpar e cancelar
        objFormCadastrarTipoInterface.cmdSalvar.Enabled = False
        objFormCadastrarTipoInterface.cmdLimpar.Enabled = False
        objFormCadastrarTipoInterface.cmdCancelar.Enabled = False
    
    End If

End Sub

Public Sub LimparDadosTelaTipoInterface()

    objFormCadastrarTipoInterface.TxtDescricao.Text = ""
    objFormCadastrarTipoInterface.OptEmbratel.Value = False
    objFormCadastrarTipoInterface.optCliente.Value = False
    
    objCadastrarTipoInterface.Descricao = ""
    objCadastrarTipoInterface.Tipo_Interface = ""

    
    Exit Sub
        
End Sub

Public Sub Saida(ByVal pyTela As Byte)
    
    Unload objFormCadastrarTipoInterface
    
End Sub


