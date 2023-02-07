VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsgenOC_SelecionarOTSCorrecao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Declara as classes utilizadas
Public lRow                  As Long
Private sFamilia_Produto     As String
Private objMensagem          As clsgenMensagem
Private objServico           As New clsGenServico
Private objListaOTS          As New clsgenListaOTS
Private objFase              As New clsgenFase
Private lCodigoFase          As Long
Private lProduto_ID          As Long

Private lGrupo_ID            As Long

Private lFase_Id             As Long
Private lServico_ID          As Long
Private lServico_ID_Anterior As Long
Private lAcao_ID             As Long
Private sSigla_Acao          As String
Private sDataFimReal         As String
Private sDataInicioPrevisto  As String
Private iRegiao              As Integer
Private lSequencia_Acao      As Long
Private lNumeroOE            As Long
Private lAnoOE               As Long
Private lNumeroITemOE        As Long
Private lContratoID          As Long
Private lNrLocalInsta        As Long
Private lCircuitoID          As Long
Private sIndicaCCTOTemp      As String

Private objCentroFuncional       As clsgenCentroFuncional

' Declara os formulários usados nesta classe
Private objFormSelecionarOTS As frmgenSelecionarOTSCorrecao
Private sAnotacao            As String

' Declara variáveis
Private aListaCliente()      As Variant
Private iSequenciaFase       As Integer

Public Property Get Anotacao() As String

    Anotacao = sAnotacao

End Property

Public Property Let Anotacao(ByVal sNewValue As String)

    sAnotacao = sNewValue

End Property
Public Property Get Grupo_ID() As Long

    Grupo_ID = lGrupo_ID

End Property

Public Property Let Grupo_ID(ByVal vNewValue As Long)

    lGrupo_ID = vNewValue

End Property

Public Property Get Produto_ID() As Long

    Produto_ID = lProduto_ID

End Property

Public Property Let Produto_ID(ByVal vNewValue As Long)

    lProduto_ID = vNewValue

End Property

Public Property Get CodigoFase() As Long

    CodigoFase = lCodigoFase

End Property

Public Property Let CodigoFase(ByVal vNewValue As Long)

    lCodigoFase = vNewValue

End Property

Public Function TelaSelecionarOTS(ByVal psTitulo As String) As Boolean

On Error GoTo ErrorHandler:

    Dim bRetorno As Boolean
     
    Dim vTodos As String
    
    vTodos = "[Todos]"
     
    Screen.MousePointer = vbArrowHourglass
     
    Set objMensagem = New clsgenMensagem
     
    If Verifica_Janela_Aberta("frmgenSelecionarOTSCorrecao", "Correção de OTS") Then
        GoTo Finalizar
    End If
     
    Set objFormSelecionarOTS = New frmgenSelecionarOTSCorrecao
    
    objFormSelecionarOTS.informa_OC Me
    
    objFormSelecionarOTS.Produto_ID = Produto_ID
    
    objFormSelecionarOTS.Caption = psTitulo
    
    'CALBAT - 05/10/2007 ================================================
    'Inserido o combo de pesquisa por CF , a pedido da AR Aline Nascimento.        =
    ' com  a aprovação do Coordenador Jaime Canha                                              =
    '=====================================================================
    Call CarregaTituloProduto(objFormSelecionarOTS, objFormSelecionarOTS.Caption)
    '*******************************************************************
    '
    ' Responsável : Accenture – Leonardo Ferreira
    ' Data        : 08/05/2008
    ' BO          : CH789557
    ' Observação  : Define período
    '
    '*******************************************************************
        '
    objFormSelecionarOTS.dtpDataIni.Value = Format(Now - 10, "dd/mm/yyyy")
    objFormSelecionarOTS.dtpDataFim.Value = Format(Now, "dd/mm/yyyy")
    '*******************************************************************
    
    ' Executar o método de exibição dos centros funcionais
    bRetorno = ExibirListaCentroFuncional(objFormSelecionarOTS.cboCentroFunc)
    
    'RN_SGAD_21148YBJ
    If (GbFlagRN_SGAD_21148YBJ) Then
        
        objFormSelecionarOTS.cboCentroFunc.AddItem vTodos
        objFormSelecionarOTS.cboCentroFunc.ItemData(objFormSelecionarOTS.cboCentroFunc.NewIndex) = 0
        
    End If
    '=====================================================================
    '=====================================================================
    
    bRetorno = ExibirListaProduto(objFormSelecionarOTS.CboServico, Familia_Produto)
    
    objFormSelecionarOTS.CboServico.AddItem "", 0
    
    If Not bRetorno Then
        
        Unload objFormSelecionarOTS
        GoTo Finalizar
        
    End If
        
    objFormSelecionarOTS.Show
     
    GoTo Finalizar
       
Finalizar:

    Screen.MousePointer = vbDefault
    
    On Error GoTo 0
    
    If objMensagem.TipoMsg <> "" Then
    
        objMensagem.ExibeMensagem
        
    End If
    
    Set objMensagem = Nothing
    
    Exit Function
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: TelaSelecionarOTS " & _
                            "da classe clsgenOC_SelecionarOTSCorrecao"
    GoTo Finalizar

End Function

'Public Sub TelaAnotacoes(ByVal pbChamouDaSelecionarOTS As Boolean)
'
'On Error GoTo ErrorHandler:
'
'    Dim bRetorno                 As Boolean
'    Dim objFormAnotacoes     As frmgenAnotacoesAcumuladas
'
'    Screen.MousePointer = vbArrowHourglass
'
'    ' Intância a classe de mensagem
'    Set objMensagem = New clsgenMensagem
'
'    ' Instância o formulário de que exibe os números de designação e OTS
'    Set objFormAnotacoes = New frmgenAnotacoes
'
'    ' Atribui o valor da propriedade no objeto da tela correspondente ao texto
'    objFormAnotacoes.txtTexto = objFase.Anotacao
'
'    ' Atribui o título da tela
'    If lProduto_ID = 1 Then
'
'        ' Dedicated
'        objFormAnotacoes.Caption = "Worldcom Internet Dedicated"
'
'    ElseIf lProduto_ID = 3 Then
'
'        ' Diginet
'        objFormAnotacoes.Caption = "Diginet"
'
'    ElseIf lProduto_ID = 2 Then
'
'        ' BLF
'        objFormAnotacoes.Caption = "Business Link Flex"
'
'    End If
'
'    ' Se a chamada for da tela se seleção de OTS, colocar o botão de gravação visível
'    If pbChamouDaSelecionarOTS Then
'
'        objFormAnotacoes.cmdGravar.Enabled = True
'        objFormAnotacoes.cmdGravar.Visible = True
'
'        objFormSelecionarOTS.spdCircuitos.Row = objFormSelecionarOTS.spdCircuitos.ActiveRow
'        objFormAnotacoes.lblCabecalho.Caption = "Anotação OTS: " & Mid(objFormSelecionarOTS.spdCircuitos.Text, 2)
'
'    Else
'
'        objFormAnotacoes.cmdGravar.Enabled = False
'        objFormAnotacoes.cmdGravar.Visible = False
'
'        objFormAnotacoes.lblCabecalho.Caption = "Anotação OTS: " & objFase.OTS
'
'    End If
'
'    ' Informar o objeto controlador
'    objFormAnotacoes.informa_OC Me
'
'    ' Informa o banco para a classe de fase
'    objFase.InformaBanco frmgenMDI_SGAPlus.objBanco
'
'    ' Informa a mensagem para a classe de fase
'    objFase.InformaMensagem objMensagem
'
'    ' Executa o método para recuperar os de anotação da fase
'    objFormSelecionarOTS.spdCircuitos.Row = objFormSelecionarOTS.spdCircuitos.ActiveRow
'    objFormSelecionarOTS.spdCircuitos.Col = 7
'
'    bRetorno = objFase.RecuperarAnotacao(CLng(objFormSelecionarOTS.spdCircuitos.Text))
'
'    ' Se ocorreu algum erro na DLL
'    If Not bRetorno Then
'
'        GoTo Finalizar
'
'    End If
'
'    objFormAnotacoes.txtTexto.Text = objFase.Anotacao
'
'    objFormAnotacoes.Show vbModal
'
'    Screen.MousePointer = vbDefault
'
'    objFase.Anotacao = objFormAnotacoes.sTexto
'
'Finalizar:
'
'    Set objFormAnotacoes = Nothing
'
'    Exit Sub
'
'ErrorHandler:
'
'    objMensagem.TipoMsg = "E"
'    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
'                            "Houve um erro de execução no método: TelaAnotacoes " & _
'                            "da classe clsgenOC_SelecionarOTSCorrecao"
'
'    GoTo Finalizar
'
'End Sub

'Public Sub TelaAnotacoesGen(ByVal pbChamouDaSelecionarOTS As Boolean)
'
'On Error GoTo ErrorHandler:
'
'    Dim bRetorno                As Boolean
'    Dim objAnotacao             As New clsGenAnotacao
'    Dim objDetalheAnotacao      As New clsGenDetalheAnotacao
'    Dim lRow                    As Long
'    Dim lFase                   As Long
'    Dim objFormAnotacoes        As frmgenAnotacoesAcumuladas
'    Dim lSequenciaAcao          As Long
'
'    Screen.MousePointer = vbArrowHourglass
'
'    Set objMensagem = New clsgenMensagem
'
'    Set objFormAnotacoes = New frmgenAnotacoesAcumuladas
'
'    objFormAnotacoes.Top = 1290
'
'    If lProduto_ID = 1 Then
'
'        ' Dedicated
'        objFormAnotacoes.Caption = "Worldcom Internet Dedicated"
'
'    ElseIf lProduto_ID = 3 Then
'
'        ' Diginet
'        objFormAnotacoes.Caption = "Diginet"
'
'    ElseIf lProduto_ID = 2 Then
'
'        ' BLF
'        objFormAnotacoes.Caption = "Business Link Flex"
'
'    End If
'
'    objDetalheAnotacao.InformaBanco frmgenMDI_SGAPlus.objBanco
'    objDetalheAnotacao.InformaMensagem objMensagem
'
'    'Coluna 7 - Armazena o ID da FASE
'    objFormSelecionarOTS.spdCircuitos.Row = objFormSelecionarOTS.spdCircuitos.ActiveRow
'    objFormSelecionarOTS.spdCircuitos.Col = 7
'
'    lFase = CLng(objFormSelecionarOTS.spdCircuitos.Text)
'
'    If objDetalheAnotacao.RecuperarDados(lFase, "A") Then
'
'        objFormAnotacoes.txtTexto.Text = objDetalheAnotacao.Anotacao
'
'    End If
'
'    If Not pbChamouDaSelecionarOTS Then
'
'       objFormAnotacoes.txtTexto.Text = Anotacao
'
'    End If
'
'    ' Se a chamada for da tela se seleção de OTS, colocar o botão de gravação visível
'    objFormSelecionarOTS.spdCircuitos.Row = objFormSelecionarOTS.spdCircuitos.ActiveRow
'    objFormSelecionarOTS.spdCircuitos.Col = 2
'
'    If pbChamouDaSelecionarOTS Then
'
'        objFormAnotacoes.cmdGravar.Enabled = True
'        objFormAnotacoes.cmdGravar.Visible = True
'
'        objFormAnotacoes.lblCabecalho.Caption = "Anotação OTS: " & objFormSelecionarOTS.spdCircuitos.Text
'
'    Else
'
'        objFormAnotacoes.cmdGravar.Enabled = False
'        objFormAnotacoes.cmdGravar.Visible = False
'
'        objFormAnotacoes.lblCabecalho.Caption = "Anotação OTS: " & objFormSelecionarOTS.spdCircuitos.Text
'
'    End If
'
'    ' Informar o objeto controlador
'    objFormAnotacoes.informa_OC Me
'
'    objAnotacao.InformaBanco frmgenMDI_SGAPlus.objBanco
'    objAnotacao.InformaMensagem objMensagem
'
'    ' Executa o método para recuperar os de anotação da fase
'
'    'Coluna 22- Armazena a Sequencia_Acao
'    objFormSelecionarOTS.spdCircuitos.Row = objFormSelecionarOTS.spdCircuitos.ActiveRow
'    objFormSelecionarOTS.spdCircuitos.Col = 22
'    lSequenciaAcao = CLng(objFormSelecionarOTS.spdCircuitos.Text)
'
'    'Coluna 8 - Armazena o ID do Servico
'    objFormSelecionarOTS.spdCircuitos.Row = objFormSelecionarOTS.spdCircuitos.ActiveRow
'    objFormSelecionarOTS.spdCircuitos.Col = 8
'
'    bRetorno = objAnotacao.RecuperarAnotacaoServico(CLng(objFormSelecionarOTS.spdCircuitos.Text), CStr(lSequenciaAcao))
'
'    ' Se ocorreu algum erro na DLL
'
'    objFormAnotacoes.SpdListaAnotacao.MaxRows = 0
'
'    If bRetorno Then
'
'        lRow = 1
'
'        For Each objDetalheAnotacao In objAnotacao.ObjColecaoAnotacao
'
'            objFormAnotacoes.SpdListaAnotacao.MaxRows = lRow
'            objFormAnotacoes.SpdListaAnotacao.Row = lRow
'
'            objFormAnotacoes.SpdListaAnotacao.Col = 1
'            objFormAnotacoes.SpdListaAnotacao.Text = objDetalheAnotacao.NomeFase
'
'            objFormAnotacoes.SpdListaAnotacao.Col = 2
'            objFormAnotacoes.SpdListaAnotacao.Text = objDetalheAnotacao.Anotacao
'
'            objFormAnotacoes.ListaAnotacao.AddItem objDetalheAnotacao.NomeFase
'
'            lRow = lRow + 1
'
'        Next
'
'        objFormAnotacoes.ListaAnotacao.ListIndex = 0
'
'    End If
'
'    objFormAnotacoes.Show vbModal
'
'    Screen.MousePointer = vbDefault
'
'Finalizar:
'
'    Set objAnotacao = Nothing
'
'    Exit Sub
'
'ErrorHandler:
'
'    objMensagem.TipoMsg = "E"
'    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
'                            "Houve um erro de execução no método: TelaAnotacoes " & _
'                            "da classe clsgenOC_SelecionarOTSCorrecao"
'
'    GoTo Finalizar
'
'End Sub

Public Function ExibirListaCentroFuncional(ByRef cbo As ComboBox) As Boolean

On Error GoTo ErrorHandler:

    Dim bRetorno As Boolean
    Dim aListaCentroFuncional() As Variant
    
    ExibirListaCentroFuncional = False
    
    ' Intância a classe de mensagem
    Set objMensagem = New clsgenMensagem
    
    ' Limpa a combo
    cbo.Clear
    
    ' Instância a classe de centro funcional
    Set objCentroFuncional = New clsgenCentroFuncional

    ' Informar o nome do banco desejado
    objCentroFuncional.InformaBanco frmgenMDI_SGAPlus.objBanco
                
    ' Informar o nome do objeto de mensagem
    objCentroFuncional.InformaMensagem objMensagem
    
    ' Executar o método de montagem da lista de centro funcional na classe de analisar order entry
    bRetorno = objCentroFuncional.MontarListaCentroFuncional(aListaCentroFuncional)
    
    ' Se ocorreu algum erro na DLL exibi-lo
    If Not bRetorno Then
        
        GoTo Finalizar
    
    End If
    
    PreencheCombo cbo, aListaCentroFuncional

    ExibirListaCentroFuncional = True
    
    GoTo Finalizar
       
Finalizar:

    Screen.MousePointer = vbDefault
    
    If objMensagem.TipoMsg <> "" Then
    
        objMensagem.ExibeMensagem
        
    End If
    
    Set objCentroFuncional = Nothing
    
    Exit Function
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: ExibirListaCentroFuncional " & _
                            "da classe clsgenOC_SelecionarOTSCorrecao"
    
    GoTo Finalizar

End Function

Public Sub CarregarOTSCorrecao()

On Error GoTo ErrorHandler:

    Dim bRetorno       As Boolean
    Dim objOTSCorrecao As clsgenOC_OTSCorrecao
     
    Screen.MousePointer = vbArrowHourglass
    
    ' Intância a classe de mensagem
    Set objMensagem = New clsgenMensagem

    ' Instância o objeto controle de correção de OTS
    Set objOTSCorrecao = New clsgenOC_OTSCorrecao
    
    'Obtém o ID Fase
    objFormSelecionarOTS.spdCircuitos.Col = 8 'ID Fase
    objFormSelecionarOTS.spdCircuitos.Row = objFormSelecionarOTS.spdCircuitos.ActiveRow
    
    objOTSCorrecao.Fase_ID = objFormSelecionarOTS.spdCircuitos.Text

    ' Carrega a tela de correção de OTS
    objFormSelecionarOTS.spdCircuitos.Col = 9 'ID servico
    objFormSelecionarOTS.spdCircuitos.Row = objFormSelecionarOTS.spdCircuitos.ActiveRow
    
    'Informa o ID Produto. IP VPN
    objOTSCorrecao.Produto_ID = Produto_ID
    
    objFormSelecionarOTS.spdCircuitos.Col = 49
    objFormSelecionarOTS.spdCircuitos.Row = objFormSelecionarOTS.spdCircuitos.ActiveRow
    
    'Informa o ID Grupo.
    objOTSCorrecao.Grupo_ID = objFormSelecionarOTS.spdCircuitos.Text
    
    Call objOTSCorrecao.InformaOCSelecionarOTS(Me)
    
    'objFormSelecionarOTS.spdCircuitos.Col = 50
    'objFormSelecionarOTS.spdCircuitos.Row = objFormSelecionarOTS.spdCircuitos.ActiveRow
    
    'Informa o ID Grupo.
    'objOTSCorrecao = objFormSelecionarOTS.spdCircuitos.Text
    
    
    ' Carrega a tela de correção de OTS
    objFormSelecionarOTS.spdCircuitos.Col = 9 'ID servico
    objFormSelecionarOTS.spdCircuitos.Row = objFormSelecionarOTS.spdCircuitos.ActiveRow
    
    bRetorno = objOTSCorrecao.TelaOTSCorrecao(CLng("0" & objFormSelecionarOTS.spdCircuitos.Text))
    
    ' Se ocorreu algum erro na DLL
    If Not bRetorno Then
        
        GoTo Finalizar
    
    End If
    
Finalizar:

    Screen.MousePointer = vbDefault
    
    Set objOTSCorrecao = Nothing
    Set objMensagem = Nothing
    
    Exit Sub
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: TelaAnotacoes " & _
                            "da classe clsgenOC_SelecionarOTSCorrecao"
    
    GoTo Finalizar

End Sub

Public Function ExibirListaOTS() As Boolean

On Error GoTo ErrorHandler:

    Dim bRetorno            As Boolean
    Dim sDataInicial        As String
    Dim sDataFinal          As String
    Dim iCentroFuncional    As Long
    Dim sAcao               As String
    Dim sStatusARealizar    As String
    Dim sStatusRealizado    As String
    Dim sStatusOTS          As String
    Dim sStatusFaseRealocar As String
    Dim sContaCliente       As String
    Dim sRegioes            As String
    
    
    ExibirListaOTS = False
    
    ' Intância a classe de mensagem
    Set objMensagem = New clsgenMensagem
    
    'Valida a escolha de pelo menos um status
    If (objFormSelecionarOTS.chkOTSOperEmAndamento.Value = 0 And _
            objFormSelecionarOTS.chkOTSOperEncaminhadas.Value = 0) Then

        ' Atualiza a mensagem de retorno como gravação realizada com sucesso
        objMensagem.TipoMsg = "A"
        objMensagem.Descricao = "Selecione pelo menos um status"
        GoTo Finalizar
    
    End If
    
    '*******************************************************************
    '
    ' Responsável : Accenture – Gladson M França
    ' Data        : 08/05/2008
    ' BO          : CH789557
    ' Observação  : Retirarei a obrigatoriedade de Criticar o CF
    '
    '*******************************************************************
    If Trim(objFormSelecionarOTS.cboCentroFunc.Text) = "" Then
    
        iCentroFuncional = 0
    '*******************************************************************
    Else
        
        If objFormSelecionarOTS.cboCentroFunc.ListIndex > -1 Then
            
            If objFormSelecionarOTS.cboCentroFunc.ListIndex < objFormSelecionarOTS.cboCentroFunc.ListCount() Then
               
               iCentroFuncional = objFormSelecionarOTS.cboCentroFunc.ItemData(objFormSelecionarOTS.cboCentroFunc.ListIndex)
            
            Else
                
                iCentroFuncional = 0

            End If
        Else
           
           iCentroFuncional = 0
        
        End If
        
    End If
    
    
    ' Instância a classe de serviço
    objFormSelecionarOTS.spdCircuitos.MaxRows = 0
    
    objServico.InformaBanco frmgenMDI_SGAPlus.objBanco
    objServico.InformaMensagem objMensagem
    
    'Retirado do código do vagnerb
    'bRetorno = objServico.MontarListaOTSCorrecao(Familia_Produto, Produto_ID, "C")
    
    '*******************************************************************
    '
    ' Responsável : Accenture – Gladson M França
    ' Data        : 15/04/2008
    ' BO          : CH789557
    ' Observação  : Os valores dos campos dos novos filtros são aqui salvos nas novas propriedades "CNL", "Sigla", "Numero_Designacao", "Data_Inicio" e "Data_Fim" da classe clsGenServico.
    '
    '*******************************************************************
    objServico.CNL = objFormSelecionarOTS.txtCNL
    objServico.Sigla = objFormSelecionarOTS.txtSigla
    objServico.Numero_Designacao = objFormSelecionarOTS.txtNumeroDesignacao
    objServico.Data_Inicio = objFormSelecionarOTS.dtpDataIni.Day & "/" & objFormSelecionarOTS.dtpDataIni.Month & "/" & objFormSelecionarOTS.dtpDataIni.Year
    objServico.Data_Fim = objFormSelecionarOTS.dtpDataFim.Day & "/" & objFormSelecionarOTS.dtpDataFim.Month & "/" & objFormSelecionarOTS.dtpDataFim.Year
    '*******************************************************************
    
    '=================================================================
    'CALLBAT - 21/05/2012 - BSoD Net Light
    '=================================================================
    'bRetorno = objServico.MontarListaOTSCorrecao(Familia_Produto, _
                                                 Produto_ID, _
                                                 "C", _
                           objFormSelecionarOTS.chkOTSOperEmAndamento.Value = 1, _
                           objFormSelecionarOTS.chkOTSOperEncaminhadas.Value = 1, _
                           iCentroFuncional, _
                           TrataPlick(objFormSelecionarOTS.txtCliente.Text))
                           
    bRetorno = objServico.MontarListaOTSCorrecao(Familia_Produto, _
                                                 Produto_ID, _
                                                 "C','L", _
                           objFormSelecionarOTS.chkOTSOperEmAndamento.Value = 1, _
                           objFormSelecionarOTS.chkOTSOperEncaminhadas.Value = 1, _
                           iCentroFuncional, _
                           TrataPlick(objFormSelecionarOTS.txtCliente.Text))
    '=================================================================
    '=================================================================
    
    If Not bRetorno Then
        
        GoTo Finalizar
    
    End If
    
    ' Preenche a grid
    bRetorno = PreencheGrid()
    
    If Not bRetorno Then
        
        GoTo Finalizar
    
    End If
    
    ExibirListaOTS = True
    
    GoTo Finalizar
       
Finalizar:

    Screen.MousePointer = vbDefault
    
    If objMensagem.TipoMsg <> "" Then
    
        objMensagem.ExibeMensagem
        
    End If
    
    Set objMensagem = Nothing
    
    Exit Function
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: ExibirListaOTS " & _
                            "da classe clsgenOC_SelecionarOTSCorrecaoCorrecao"
    
    GoTo Finalizar

End Function

Public Function PreencheGrid() As Boolean

    Dim iLinha          As Integer
    Dim sStatus         As String
    Dim sRealocacao     As String
    
    iLinha = 0
    
    PreencheGrid = False
    
    objFormSelecionarOTS.spdCircuitos.MaxCols = 50
    
    'A o objeto de coleção da lista de OE's
    For Each objListaOTS In objServico.ListaOTS

        iLinha = iLinha + 1
        objFormSelecionarOTS.spdCircuitos.MaxRows = iLinha

        With objFormSelecionarOTS.spdCircuitos

            ' Atribui o valor da linha e define um range para cores
            
            .Row = iLinha
            .BlockMode = True
            .Row2 = iLinha
            .Col = 1
            .Col2 = .MaxCols
            
            'Status da OTS de Correção
            If objListaOTS.StatusOTS = "AN" Then
                
                .ForeColor = 12582912  ' Azul
            
            End If
            
            .Col = 1
            .Text = objListaOTS.NomeProduto
            
            .Col = 2
            If objListaOTS.Acao = "ATV" Then
        
                .Text = "Ativação"
                
            ElseIf objListaOTS.Acao = "ALT" Then
            
                .Text = "Alteração"
                
            ElseIf objListaOTS.Acao = "DES" Then
            
                .Text = "Desativação"
                
            ElseIf objListaOTS.Acao = "CAN" Then
            
                .Text = "Cancelamento"
                
            ElseIf (objListaOTS.Acao = "ATG") Then
            
                .Text = "Correção de Ativação"
                
            ElseIf (objListaOTS.Acao = "ALG") Then
            
                .Text = "Correção de Alteração"
                
            ElseIf (objListaOTS.Acao = "ALC") Then
            
                .Text = "Alteração Operacional"
                
            ElseIf (objListaOTS.Acao = "ATO") Then
            
                .Text = "Correção de  Ativação - Operacional"
            
            ElseIf (objListaOTS.Acao = "ALO") Then
            
                .Text = "Correção de Alteração - Operacional"
                
            End If
        
            If objListaOTS.Origem = "SGA" Then
               
               .Text = .Text & " Operacional"
               
            End If
            
            ' Número da OTS
            .Col = 3
            
            'Padronização do número de OTS implementado a pedido do SPRADO em 08/09/2003
            .Text = objListaOTS.CentroFuncional & _
                    " - " & _
                    objListaOTS.Produto & _
                    " - " & _
                    CStr(objListaOTS.NumeroOTS) & _
                    "/" & _
                    objListaOTS.AnoOTS
            
            
            ' Designação do Circuito
            .Col = 4
            ' Dedicated - VPN RA - VPN Direct
            If (CLng(objListaOTS.Grupo_ID) = GbIDGrupoWCOM) Then
               
               .Text = objListaOTS.CNL_Ponta_A & _
                       " WD " & _
                       Format(objListaOTS.NumeroDesignacao, "00000")
            
            ElseIf (CLng(objListaOTS.ID_Produto) = GbIDProdutoDiginet) Or _
                (CLng(objListaOTS.ID_Produto) = GbIDProdutoEmbratelIPL) Then

                   If objListaOTS.CNL_Ponta_A <> "" Or _
                      objListaOTS.CNL_Ponta_B <> "" Then
                      
                      .Text = objListaOTS.CNL_Ponta_A & " " & objListaOTS.CNL_Ponta_B & _
                  IIf(CLng(objListaOTS.ID_Produto) = GbIDProdutoDiginet, " DNI ", "IPL") & _
                                         objListaOTS.Codigo_NP & " " & _
                           Format(objListaOTS.NumeroDesignacao, "00000")
                   
                   End If
                
            ElseIf CLng(objListaOTS.Grupo_ID) = GbIDGrupoBLF Then
               
               .Text = objListaOTS.CNL_Ponta_A & _
                       "/BF/" & _
                       Format(objListaOTS.NumeroDesignacao, "00000")
            
            ElseIf (CLng(objListaOTS.Grupo_ID) = GbIDGrupoBLD) Then
               
               .Text = objListaOTS.CNL_Ponta_A & _
                       "/IP/" & _
                       Format(objListaOTS.NumeroDesignacao, "00000")
            
            
            'Designação VPN. IP VPN
            ElseIf CLng(objListaOTS.Grupo_ID) = GbIDProdutoBIV Then
               
               .Text = objListaOTS.CNL_Ponta_A & _
                       "/IV/" & _
                       Format(objListaOTS.NumeroDesignacao, "00000")
            
            'Designação JAH.
            ElseIf CLng(objListaOTS.ID_Produto) = GbIDProdutoJAH Then
               
               .Text = objListaOTS.CNL_Ponta_A & _
                       "/JAH/" & _
                       Format(objListaOTS.NumeroDesignacao, "00000")
            
            End If
                
            ' Nome do Cliente
            .Col = 5
            .Text = objListaOTS.NomeCliente
            
            ' Data Solicitacao
            .Col = 6
            .Text = Format(objListaOTS.DataSolicitacaoCorrecao, "dd/mm/yyyy")
            
            .Col = 7
            
            If objListaOTS.Origem = "SGA" Then
               
               .Text = "Em Andamento"
               
            ElseIf objListaOTS.Origem = "SOE" Then
            
               .Text = "Encaminhada "
            
            End If

            .Col = 8
            .Text = objListaOTS.Id_Fase

            .Col = 9
            .Text = objListaOTS.ID_Servico
            
            .Col = 10
            .Text = objListaOTS.DataInicioPrevisto
            
            .Col = 11
            .Text = Format(objListaOTS.DataInicioPrevAnterior, "dd/mm/yyyy")
            
            .Col = 12
            .Text = Format(objListaOTS.DataFimPrevAnterior, "dd/mm/yyyy")
            
            .Col = 13
            If objListaOTS.DataInicioReal = "0:00:00" Or objListaOTS.DataInicioReal = "00:00:00" Then
            
               .Text = ""
            
            Else
            
               .Text = objListaOTS.DataInicioReal
            
            End If
            
            .Col = 14
            .Text = objListaOTS.Acao
            
            .Col = 15
            .Text = objListaOTS.Ots
            
            .Col = 16
            .Text = objListaOTS.CNL_Ponta_A
            
            .Col = 17
            .Text = objListaOTS.CNL_Ponta_B
            
            .Col = 18
            .Text = objListaOTS.Codigo_NP
            
            .Col = 19
            .Text = objListaOTS.NumeroDesignacao
            
            .Col = 20
            .Text = objListaOTS.NomeCliente
            
            .Col = 21
            .Text = objListaOTS.Origem
            
            .Col = 22
            .Text = objListaOTS.Correcao
            
            .Col = 23
            .Text = objListaOTS.StatusOTS
            
            .Col = 38
            .Text = objListaOTS.ID_Produto
            
            .Col = 39
            .Text = objListaOTS.ID_ACAO
            
            .Col = 40
            .Text = objListaOTS.Acao
            
            'GbFlagRN_SGAD_28616PAX
            .Col = 48
            .Text = objListaOTS.RegistroSaip
            
            
            'Tipo do Item. IP VPN
            '.Col = 48
            '.Text = objListaOTS.TipoItem    'Site, Topologia, Extranet
            
            .Col = 49
            .Text = objListaOTS.Grupo_ID
            
            .Col = 50
            .Text = objListaOTS.Nome_Grupo
            
        End With
    
    Next
    
    PreencheGrid = True

End Function

Public Function RealocarCentroFuncional(ByVal pbChamouDaSelecionarOTS As Boolean, _
                                        Optional psTipoCorrecao As String, _
                                        Optional pbIndicaChamadaConsulta As Boolean) As Boolean

On Error GoTo ErrorHandler:

    Dim bRetorno                    As Boolean
    Dim objDadosCtrlCorrecao        As New clsGenDadosControleCorrecao
    Dim objFormAnotacoes            As frmGenObservacao
    Dim objDadosProduto             As clsGenProduto
    
    Dim objDadosFase                As clsgenFase
    Dim objDadosUsuario             As clsgenUsuario
    Dim lSequenciaAcao              As Long
    Dim lIDFase                     As Long
    Dim lIDServico                  As Long
    Dim lIdProduto                  As Long
    
    If objFormSelecionarOTS.spdCircuitos.ActiveRow <= 0 Or objFormSelecionarOTS.spdCircuitos.MaxRows = 0 Then
    
        Exit Function
        
    End If
    
    Screen.MousePointer = vbArrowHourglass
    
    Set objMensagem = New clsgenMensagem
    
    Set objFormAnotacoes = New frmGenObservacao

    objFormAnotacoes.Top = 1290
    
    objFormSelecionarOTS.spdCircuitos.Row = objFormSelecionarOTS.spdCircuitos.ActiveRow
    objFormSelecionarOTS.spdCircuitos.Col = 38
    
    lIdProduto = CLng(objFormSelecionarOTS.spdCircuitos.Text)
    
    Set objDadosProduto = New clsGenProduto
    
    objDadosProduto.InformaBanco frmgenMDI_SGAPlus.objBanco
    objDadosProduto.InformaMensagem objMensagem
    
    If objDadosProduto.RecuperarDados(lIdProduto) Then
        
        objFormAnotacoes.Caption = objDadosProduto.Descricao
    
    End If
    
    objDadosCtrlCorrecao.InformaBanco frmgenMDI_SGAPlus.objBanco
    objDadosCtrlCorrecao.InformaMensagem objMensagem
    
    'Coluna 8 - Armazena o ID da Fase
    objFormSelecionarOTS.spdCircuitos.Row = lRow
    objFormSelecionarOTS.spdCircuitos.Col = 8

    lIDFase = CLng(objFormSelecionarOTS.spdCircuitos.Text)

    'Coluna 9 - Armazena o ID do Servico
    objFormSelecionarOTS.spdCircuitos.Col = 9
    lIDServico = CLng(objFormSelecionarOTS.spdCircuitos.Text)

    Set objDadosFase = New clsgenFase
    
    objDadosFase.InformaBanco frmgenMDI_SGAPlus.objBanco
    objDadosFase.InformaMensagem objMensagem
    
'   lIDFase = objDadosFase.ObtemIDFaseCorrigir(lIDServico, "C")
    
    'Obtem anotacao registrada para justificar a realocação do Centro Funcional da fase
    
    If objDadosCtrlCorrecao.RecuperarDados(lIDFase, "C") Then
        
        objFormAnotacoes.FrameObsCorrecao.Caption = "Motivo Solicitação"
        objFormAnotacoes.TxtObsCorrecao.Text = objDadosCtrlCorrecao.Justificativa
        objFormAnotacoes.LblDataSolicitacao.Caption = objDadosCtrlCorrecao.DataSolicitacao
        
        Set objDadosUsuario = New clsgenUsuario
        
        objDadosUsuario.InformaBanco frmgenMDI_SGAPlus.objBanco
        objDadosUsuario.InformaMensagem objMensagem
        
        bRetorno = objDadosUsuario.RecuperarDados(objDadosCtrlCorrecao.UsuarioIDSolic)
        
        If bRetorno Then
            
            objFormAnotacoes.LblUsuarioSolicitacao.Caption = objDadosUsuario.Login
             
        End If
        
        Set objDadosUsuario = Nothing
        
    Else
        
        objFormAnotacoes.LabelSolicitacao.Visible = False
        objFormAnotacoes.LblUsuarioSolicitacao.Visible = False
        objFormAnotacoes.LabelData.Visible = False
        objFormAnotacoes.LblDataSolicitacao.Visible = False
    
    End If

    ' Informar o objeto controlador
    objFormAnotacoes.informa_OC Me

    ' Se a chamada for da tela se seleção de OTS, colocar o botão de gravação visível
    objFormSelecionarOTS.spdCircuitos.Row = lRow
    objFormSelecionarOTS.spdCircuitos.Col = 3

    If pbChamouDaSelecionarOTS Then

        objFormAnotacoes.cmdGravar.Enabled = True
        objFormAnotacoes.cmdGravar.Visible = True

        objFormAnotacoes.lblCabecalho.Caption = "OTS: " & _
                                                objFormSelecionarOTS.spdCircuitos.Text

    Else

        objFormAnotacoes.cmdGravar.Enabled = False
        objFormAnotacoes.cmdGravar.Visible = False
        objFormAnotacoes.txtTexto.Enabled = False
        objFormAnotacoes.OptCorrecao.Enabled = False
        objFormAnotacoes.OptRealocacao.Enabled = False
        objFormAnotacoes.OptDevolucao.Enabled = False
        
        objFormAnotacoes.lblCabecalho.Caption = "OTS: " & _
                                                objFormSelecionarOTS.spdCircuitos.Text
    
    End If

    If psTipoCorrecao = "C" Then
       
       objFormAnotacoes.OptCorrecao.Value = True
    
    ElseIf psTipoCorrecao = "D" Then
    
       objFormAnotacoes.OptDevolucao.Value = True
       objFormAnotacoes.OptCorrecao.Enabled = False
       objFormAnotacoes.OptRealocacao.Enabled = False
    
    End If
    
    If pbIndicaChamadaConsulta Then
    
        objFormAnotacoes.framePrincipal.Visible = False
        objFormAnotacoes.FrameObsCorrecao.Height = 2640
        objFormAnotacoes.TxtObsCorrecao.Height = 2310
        objFormAnotacoes.cmdGravar.Visible = False
        
    End If
    
    objFormAnotacoes.Show vbModal

    Screen.MousePointer = vbDefault

Finalizar:

    Set objDadosCtrlCorrecao = Nothing

    Exit Function

ErrorHandler:

    objMensagem.TipoMsg = "E"
    
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: TelaAnotacoes " & _
                            "da classe clsgenOC_SelecionarOTS"

    objMensagem.ExibeMensagem
    
    GoTo Finalizar

End Function


Public Function GravarJustificativaDevolucao() As Boolean

On Error GoTo ErrorHandler:

    Dim bRetorno                   As Boolean
    Dim objDadosCtrlCorrecao       As New clsGenDadosControleCorrecao
    Dim objDadosFase               As clsgenFase
    Dim objdadosServico            As clsGenServico
    Dim lIDFase                    As Long
    Dim lIDServico                 As Long
    Dim sDataIniPrev               As String
    Dim sDataFimPrev               As String
    Dim FlagIniciouTransacao       As Boolean
    
    
    'GbFlagRN_SGAD_28616PAX
    Dim objDadosIntegracao         As New clsBlf_IntegracaoWebIP
    Dim sTabela                    As String
    
    GravarJustificativaDevolucao = False
    
    FlagIniciouTransacao = False
    
    Screen.MousePointer = vbArrowHourglass
    
    Set objMensagem = New clsgenMensagem

    ' Inicia a transação
    bRetorno = frmgenMDI_SGAPlus.objBanco.Iniciar_Transacao()

    ' Se ocorreu algum erro na DLL
    If Not bRetorno Then
    
        GoTo Finalizar
        
    End If

    FlagIniciouTransacao = True
    
    objDadosCtrlCorrecao.InformaBanco frmgenMDI_SGAPlus.objBanco
    objDadosCtrlCorrecao.InformaMensagem objMensagem
    
    'Coluna 11 ===> Data inicio Previsto Anterior
    objFormSelecionarOTS.spdCircuitos.Row = objFormSelecionarOTS.spdCircuitos.ActiveRow
    objFormSelecionarOTS.spdCircuitos.Col = 11
    
    sDataIniPrev = objFormSelecionarOTS.spdCircuitos.Text
    
    'Coluna 12 ===> Data Fim Previsto Anterior
    objFormSelecionarOTS.spdCircuitos.Row = objFormSelecionarOTS.spdCircuitos.ActiveRow
    objFormSelecionarOTS.spdCircuitos.Col = 12
    
    sDataFimPrev = objFormSelecionarOTS.spdCircuitos.Text
    
    'Coluna 9 ===> ID SERVICO
    objFormSelecionarOTS.spdCircuitos.Row = objFormSelecionarOTS.spdCircuitos.ActiveRow
    objFormSelecionarOTS.spdCircuitos.Col = 9
    
    lIDServico = CLng(objFormSelecionarOTS.spdCircuitos.Text)
    
    'Coluna 8 ===> ID FASE
    objFormSelecionarOTS.spdCircuitos.Row = objFormSelecionarOTS.spdCircuitos.ActiveRow
    objFormSelecionarOTS.spdCircuitos.Col = 8
    
    lIDFase = CLng(objFormSelecionarOTS.spdCircuitos.Text)
    
    objDadosCtrlCorrecao.FaseID = CLng(objFormSelecionarOTS.spdCircuitos.Text)

    objDadosCtrlCorrecao.DataExecucao = Format(Date, "dd/mm/yyyy")
    objDadosCtrlCorrecao.UsuarioIDExec = GbID_Usuario
    
    'GbFlagRN_SGAD_28616PAX
    If (GbFlagRN_SGAD_28616PAX) Then
    
        'Registro Processado no SAIP. Coluna 48
        objFormSelecionarOTS.spdCircuitos.Row = objFormSelecionarOTS.spdCircuitos.ActiveRow
        objFormSelecionarOTS.spdCircuitos.Col = 48
        
        objDadosCtrlCorrecao.RegistroSaip = ("" & objFormSelecionarOTS.spdCircuitos.Text)
        
    End If
    
    bRetorno = objDadosCtrlCorrecao.AtualizarExecucao(lIDFase, "C")
    
    If objMensagem.TipoMsg <> "" Then
    
        GoTo ErrorHandler
        
    End If

    objDadosCtrlCorrecao.Justificativa = TrataPlick(Anotacao)
    objDadosCtrlCorrecao.DataSolicitacao = ""
    
    objDadosCtrlCorrecao.DataInicioPrevAnt = "" & sDataIniPrev
    objDadosCtrlCorrecao.DataFimPrevAnt = "" & sDataFimPrev
    objDadosCtrlCorrecao.TipoCorrecao = "D"
    objDadosCtrlCorrecao.UsuarioIDSolic = GbID_Usuario
    
    If Not objDadosCtrlCorrecao.FaseExistente(lIDFase, "D") Then
    
        bRetorno = objDadosCtrlCorrecao.IncluirDados("D")
        
    Else
    
        bRetorno = objDadosCtrlCorrecao.AtualizarDados(lIDFase, "D")
    
    End If
        
    If objMensagem.TipoMsg <> "" Then
    
        GoTo ErrorHandler
        
    End If
    
    Set objDadosFase = New clsgenFase
    
    objDadosFase.InformaBanco frmgenMDI_SGAPlus.objBanco
    objDadosFase.InformaMensagem objMensagem
    
    bRetorno = objDadosFase.GravarCorrecao(lIDFase, "N")
        
    If objMensagem.TipoMsg <> "" Then
    
        GoTo ErrorHandler
        
    End If
        
    objDadosFase.Data_Inicio_Previsto = sDataIniPrev
    objDadosFase.Data_Fim_Previsto = sDataFimPrev
    
    bRetorno = objDadosFase.AtualizaPeriodoPrevisto(lIDFase)
        
    If objMensagem.TipoMsg <> "" Then
    
        GoTo ErrorHandler
        
    End If
    
    Set objdadosServico = New clsGenServico
    
    objdadosServico.InformaBanco frmgenMDI_SGAPlus.objBanco
    objdadosServico.InformaMensagem objMensagem
    
    bRetorno = objdadosServico.GravarCorrecao(lIDServico, "N")
    
    ' Ocorreu um erro ao executar o método
    If Not bRetorno Then

        GoTo ErrorHandler

    End If
    
    
    'GbFlagRN_SGAD_28616PAX
    'Recupera o Registro Processado do SAIP
    If (GbFlagRN_SGAD_28616PAX) Then
    
         sTabela = "INTEGRACAO_BLF"
         
        objDadosIntegracao.InformaBanco frmgenMDI_SGAPlus.objBanco
        objDadosIntegracao.InformaMensagem objMensagem
        
       'Verifica se a OTS de Topologia está ativada
        bRetorno = objDadosIntegracao.DesfazCorrecao(lIDServico, _
                                 objDadosCtrlCorrecao.RegistroSaip, _
                                 sTabela)
        
         'Tratar erro
        If (Not bRetorno) Then
            
            GoTo Finalizar
        
        End If
        
    End If

    ' Atualiza a mensagem de retorno como gravação realizada com sucesso
    objMensagem.TipoMsg = "A"
    objMensagem.Descricao = "Gravação realizada com sucesso"

    ' Finaliza a transação
    frmgenMDI_SGAPlus.objBanco.Finalizar_Transacao
    
    GravarJustificativaDevolucao = True
    
Finalizar:

    Screen.MousePointer = vbDefault
    
    Set objDadosCtrlCorrecao = Nothing
    
    'GbFlagRN_SGAD_28616PAX
    Set objDadosIntegracao = Nothing
    
    If objMensagem.TipoMsg <> "" Then
    
        objMensagem.ExibeMensagem
        
    End If
    
    Set objMensagem = Nothing
    
    On Error GoTo 0
    
    Exit Function
        
ErrorHandler:

    If FlagIniciouTransacao Then
    
       frmgenMDI_SGAPlus.objBanco.Cancelar_Transacao

    End If
    
    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: GravarAnotacao " & _
                            "da classe clsgenOC_SelecionarOTS"
    
    GoTo Finalizar

End Function

Public Property Get Familia_Produto() As String

    Familia_Produto = sFamilia_Produto

End Property

Public Property Let Familia_Produto(ByVal sNewValue As String)

    sFamilia_Produto = sNewValue
    
End Property

