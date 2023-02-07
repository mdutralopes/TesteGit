VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsgenOC_CadastrarTipoEquip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Declara as classes utilizadas
Private objMensagem                     As clsgenMensagem
Private objCadastrarTipoEquipamento       As clsgenTipoEquipamento
Private objListaTipoEquipamento           As clsgenListaTipoEquipamento

' Declara os formulários usados nesta classe
Private objFormCadastrarTipoEquipamento As frmgenCadastroTipoEquipamento

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

Public Sub TelaCadastrarTipoEquipamento()

On Error GoTo ErrorHandler:

    ' Instância o formulário de pré ativação e chama em seguida
    Set objFormCadastrarTipoEquipamento = New frmgenCadastroTipoEquipamento
    
    ' Intância a classe de mensagem
    Set objMensagem = New clsgenMensagem
    
    ' Informar o objeto controlador
    objFormCadastrarTipoEquipamento.InformaClasseOC Me
    
    ' Instância a classe de dados para cadastrar usuários
    Set objCadastrarTipoEquipamento = New clsgenTipoEquipamento
    
    ' Informar o nome do banco desejado
    objCadastrarTipoEquipamento.InformaBanco frmgenMDI_SGAPlus.objBanco
            
    ' Informar o nome do objeto de mensagem
    objCadastrarTipoEquipamento.InformaMensagem objMensagem
    
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
                            "Houve um erro de execução no método: TelaCadastrarTipoEquipamento " & _
                            "da classe clsgenOC_CadastrarTipoEquip"
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
    objCadastrarTipoEquipamento.InformaMensagem objMensagem
    
    ' Intância a classe da lista de usuários
    Set objListaTipoEquipamento = Nothing
    Set objListaTipoEquipamento = New clsgenListaTipoEquipamento
    
    ' Limpa a lista
    objFormCadastrarTipoEquipamento.lstTipoEquipamento.Clear
    
    ' Executar o método de montagem da lista na classe de analisar order entry
    bRetorno = objCadastrarTipoEquipamento.MontarLista(objFormCadastrarTipoEquipamento.txtBusca.Text)
    
    
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
    objFormCadastrarTipoEquipamento.lstTipoEquipamento.ListIndex = 0
    
Finalizar:

    ' Se não encontrou registro desabilitar os botôes de alteração e exclusão
    If lQuantidadeRegistros = 0 Then
    
        objFormCadastrarTipoEquipamento.CmdAlterar.Enabled = False
        objFormCadastrarTipoEquipamento.Cmdexcluir.Enabled = False
        
        Call LimparDadosTelaTipoEquipamento
        
    Else
    
        objFormCadastrarTipoEquipamento.CmdAlterar.Enabled = True
        objFormCadastrarTipoEquipamento.Cmdexcluir.Enabled = True
        
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
                            "da classe clswidOC_CadastrarTipoEquipamento"

    GoTo Finalizar

End Sub

Public Function GravarDados(ByVal pyTela As Byte) As Boolean

On Error GoTo ErrorHandler:

    Dim bRetorno As Boolean
    
    Dim Tipo As String

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
        bRetorno = CriticarCamposTelaTipoEquipamento
            
        If Not bRetorno Then
        
            GoTo Finalizar
            
        End If
    
        ' Informar o nome do objeto de mensagem
        objCadastrarTipoEquipamento.InformaMensagem objMensagem
        
        ' Atribui os campos da tela nas propriedades
        Call AtribuirCamposNasPropriedades
            
        ' Excutar o método para gravar os dados do usuário
        
        '==================================================
        '==================================================
        'Autor: Carlos Soares
        'Data: 25/02/2006
        'Motivo: Alteração VipLineIP
        If objFormCadastrarTipoEquipamento.OptModem.Value = -1 Then
            Tipo = "M"
        Else
            Tipo = "O"
        End If
        '==================================================
        '==================================================
        
        bRetorno = objCadastrarTipoEquipamento.GravarDados(sTipoOperacao, Tipo)
            
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
                            "da classe clsgenOC_CadastrarTipoEquip"

    GoTo Finalizar

End Function

Public Sub ExcluirDados()

On Error GoTo ErrorHandler:

    Dim bRetorno As Boolean

    ' Intância a classe de mensagem
    Set objMensagem = New clsgenMensagem
    
    objMensagem.TipoMsg = "Q"
    objMensagem.Descricao = "Confirma a exlusão do Tipo de Equipamento selecionado?"
    If objMensagem.ExibeMensagem = "S" Then

        ' Informar o nome do objeto de mensagem
        objCadastrarTipoEquipamento.InformaMensagem objMensagem
    
        ' Inicia a transação
        frmgenMDI_SGAPlus.objBanco.Iniciar_Transacao
    
           ' Excutar o método de exclusão dos dados do usuário
            bRetorno = objCadastrarTipoEquipamento.ExcluirDados(objCadastrarTipoEquipamento.ID)
    
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
            objFormCadastrarTipoEquipamento.txtBusca.Text = ""
            
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

    objCadastrarTipoEquipamento.Descricao = TrataPlick(objFormCadastrarTipoEquipamento.TxtDescricao)

End Sub

Private Sub Class_Terminate()

    Set objCadastrarTipoEquipamento = Nothing
    Set objMensagem = Nothing
    Set objListaTipoEquipamento = Nothing
    
End Sub

Private Function CriticarCamposTelaTipoEquipamento() As Boolean

On Error GoTo ErrorHandler:

    CriticarCamposTelaTipoEquipamento = False
    
    If Len(objFormCadastrarTipoEquipamento.TxtDescricao.Text) = 0 Then
        objMensagem.TipoMsg = "A"
        objMensagem.Descricao = "Descrição é obrigatória."
        objFormCadastrarTipoEquipamento.TxtDescricao.SetFocus
        
        GoTo Finalizar
    End If
    
    
    CriticarCamposTelaTipoEquipamento = True

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
                            "Houve um erro de execução no método: CriticarCamposTelaTipoEquipamento " & _
                            "da classe clsgenOC_CadastrarTipoEquip"
    
    GoTo Finalizar

End Function

Public Function SaiTela(ByVal pyTela As Byte) As Boolean

On Error GoTo ErrorHandler:

    Dim bAlterou As Boolean

    SaiTela = False

    bAlterou = AlterouDadosTelaTipoEquipamento
    
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
                            "da classe clsgenOC_CadastrarTipoEquip"
    
    GoTo Finalizar

End Function

Public Function AlterouDadosTelaTipoEquipamento() As Boolean

On Error GoTo ErrorHandler:

    AlterouDadosTelaTipoEquipamento = True
    
    Set objMensagem = New clsgenMensagem
    
    If objCadastrarTipoEquipamento.Descricao <> objFormCadastrarTipoEquipamento.TxtDescricao Then
    
        GoTo Finalizar
        
    End If

    AlterouDadosTelaTipoEquipamento = False
    
Finalizar:

    On Error GoTo 0

    Exit Function
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: AlterouDadosTelaTipoEquipamento " & _
                            "da classe clsgenOC_CadastrarTipoEquip"
    
    GoTo Finalizar

End Function

Private Function PreencheLista() As Boolean

    lQuantidadeRegistros = 0
    
    PreencheLista = False
    
    ' Ler o objeto de coleção da lista de tipos de pendências
    For Each objListaTipoEquipamento In objCadastrarTipoEquipamento.ListaTipoEquipamento

        With objFormCadastrarTipoEquipamento
            
            ' Descricao Motivo Recusa
             .lstTipoEquipamento.AddItem (objListaTipoEquipamento.Descricao)
        
        End With

        lQuantidadeRegistros = lQuantidadeRegistros + 1
    
    Next
    
    PreencheLista = True
    
End Function

Public Sub RefreshDados()

    If lQuantidadeRegistros > 0 Then
    ' Atualiza os dados de usuário na tela
    With objFormCadastrarTipoEquipamento
        
        .TxtDescricao = objCadastrarTipoEquipamento.ListaTipoEquipamento(objFormCadastrarTipoEquipamento.lstTipoEquipamento.ListIndex + 1).Descricao
        
        '==================================================
        '==================================================
        'Autor: Carlos Soares
        'Data: 25/02/2006
        'Motivo: Alteração VipLineIP
            If objCadastrarTipoEquipamento.ListaTipoEquipamento(objFormCadastrarTipoEquipamento.lstTipoEquipamento.ListIndex + 1).Tipo = "" Then
                objFormCadastrarTipoEquipamento.OptModem.Value = 0
                objFormCadastrarTipoEquipamento.OptOutrosEqto.Value = -1
            Else
                objFormCadastrarTipoEquipamento.OptModem.Value = -1
                objFormCadastrarTipoEquipamento.OptOutrosEqto.Value = 0
            End If
        '==================================================
        '==================================================
                
    End With
    
        objCadastrarTipoEquipamento.Descricao = objCadastrarTipoEquipamento.ListaTipoEquipamento(objFormCadastrarTipoEquipamento.lstTipoEquipamento.ListIndex + 1).Descricao
        objCadastrarTipoEquipamento.ID = objCadastrarTipoEquipamento.ListaTipoEquipamento(objFormCadastrarTipoEquipamento.lstTipoEquipamento.ListIndex + 1).ID
        
    End If


End Sub

Private Sub HabilitaDesabilitaObjetos()

    ' Coloca cor de fundo como default para habilitação
    If (sTipoOperacao = "I") Or (sTipoOperacao = "A") Then ' Inclusão ou alteração
        
        ' Muda cor de fundo dos objetos de entrada de dados para branco
        objFormCadastrarTipoEquipamento.TxtDescricao.BackColor = &H80000005

        
        ' Muda cor de escrita dos objetos de entrada de dados para preto
        objFormCadastrarTipoEquipamento.TxtDescricao.ForeColor = &H80000008
        
        objFormCadastrarTipoEquipamento.OptOutrosEqto.Value = -1
        objFormCadastrarTipoEquipamento.OptModem.Value = 0

        
        ' Habilita o frame dos objetos de dados
        objFormCadastrarTipoEquipamento.frameDados.Enabled = True
        
        ' Desabilita botões de inclusão, alteração, exlusão e frames de busca e lista
        objFormCadastrarTipoEquipamento.cmdIncluir.Enabled = False
        objFormCadastrarTipoEquipamento.CmdAlterar.Enabled = False
        objFormCadastrarTipoEquipamento.Cmdexcluir.Enabled = False
        objFormCadastrarTipoEquipamento.frameBusca.Enabled = False
        objFormCadastrarTipoEquipamento.frameLista.Enabled = False
        
        ' Habilita os botões gravar, limpar e cancelar
        objFormCadastrarTipoEquipamento.CmdSalvar.Enabled = True
        objFormCadastrarTipoEquipamento.CmdLimpar.Enabled = True
        objFormCadastrarTipoEquipamento.cmdCancelar.Enabled = True
        
    Else
    
        ' Muda cor de fundo dos objetos de entrada de dados para cinza
        objFormCadastrarTipoEquipamento.TxtDescricao.BackColor = &H80000004

        objFormCadastrarTipoEquipamento.OptOutrosEqto.Value = 0
        objFormCadastrarTipoEquipamento.OptModem.Value = -1
        
        ' Muda cor de escrita dos objetos de entrada de dados azul
        objFormCadastrarTipoEquipamento.TxtDescricao.ForeColor = &HC00000
        
        ' Desabilita o frame dos objetos de dados
        objFormCadastrarTipoEquipamento.frameDados.Enabled = False
    
        ' Habilita botões de inclusão, alteração, exlusão e frames de busca e lista
        objFormCadastrarTipoEquipamento.cmdIncluir.Enabled = True
        objFormCadastrarTipoEquipamento.CmdAlterar.Enabled = True
        objFormCadastrarTipoEquipamento.Cmdexcluir.Enabled = True
        objFormCadastrarTipoEquipamento.frameBusca.Enabled = True
        objFormCadastrarTipoEquipamento.frameLista.Enabled = True
        
        ' Desabilita os botões gravar, limpar e cancelar
        objFormCadastrarTipoEquipamento.CmdSalvar.Enabled = False
        objFormCadastrarTipoEquipamento.CmdLimpar.Enabled = False
        objFormCadastrarTipoEquipamento.cmdCancelar.Enabled = False
    
    End If

End Sub

Public Sub LimparDadosTelaTipoEquipamento()

    objFormCadastrarTipoEquipamento.TxtDescricao.Text = ""
    
    objCadastrarTipoEquipamento.Descricao = ""
    
    Exit Sub
        
End Sub

Public Sub Saida(ByVal pyTela As Byte)
    
    Unload objFormCadastrarTipoEquipamento
    
End Sub


