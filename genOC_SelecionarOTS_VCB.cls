VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsgenOC_SelecionarOTS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
 Option Explicit

' Declara as classes utilizadas
Public objOC_Designacao          As New clsDNIOC_DesignacaoCCTO
Private objOC_TimeSlotDD         As New clsDniOC_TimeSlotDedicado
 
Private objOcAlocarCF            As clsInterOC_AlocarCF
Private objOC_AcessoEmbratel     As clsGenOC_TratarAcessoEmbratel
Private objOC_AcessoCliente      As clsgenOC_TratarAcessoCliente
'Trata acesso DSLAM - JAH
Private objOC_AcessoDSLAM        As clsgenOC_TratarAcessoDSLAM
Private objocConfiguracao        As clsInterOC_Configuracao
Private objocTratarCentroLocal   As clsInterOC_TratarCentroLocal
Private objVIPocTratarCL         As clsVIPOC_TratarCentroLocal
Private objocCoordCPE            As clsInterOC_CoordCPE
Private objocPreAtivacao         As clsInterOC_PreAtivacao
Private objocAtivacao            As Object
Private objocTratarBGP           As clsInterOC_TratarBGP
Private objOC_TratarIP           As clsVIPOC_TratarIPFixo
Public objMensagem               As clsgenMensagem
Private objCentroFuncional       As clsgenCentroFuncional
Private objServico               As clsGenServico
Private objCliente               As clsgenCliente

'JAH
Private objocConfiguracaoTel     As clsInterOC_ConfigNGN

'NET PHONE
Private objocConfigConftel       As clsInterOC_tratarConfigConftel

'VIPLINE IP
Private objocNGNAcesso           As clsInterOC_ConfigAcessoNGN
'Private objocTratarAcessoADE     As clsInterOC_TratarAcessoADE
Private objOCTratarAcessoEBT     As clsGenOC_TratarAcessoEmbratel
Private objocNGNNucleo           As clsInterOC_tratarNGNNucleo

Private objListaOTS              As New clsgenListaOTS
Private objFase                  As New clsgenFase

Private lCodigoFase              As Long
Private lProduto_ID              As Long
Private lFase_Id                 As Long
Private lServico_ID              As Long
Private lServico_ID_Anterior     As Long
Private lAcao_ID                 As Long
Private sFamilia_Produto         As String
Private sSigla_Acao              As String
Private sDataFimReal             As String
Private sDataInicioPrevisto      As String
Private iRegiao                  As Integer
Private lSequencia_Acao          As Long
Private lNumeroOE                As Long
Private lAnoOE                   As Long
Private lNumeroITemOE            As Long
Private lContratoID              As Long
Private lNrLocalInsta            As Long
Private lCircuitoID              As Long
Private sIndicaCCTOTemp          As String

' Declara os formulários usados nesta classe

Public objFormSelecionarOTS      As Form

Private sAnotacao                As String

' Declara variáveis
Private aListaCliente() As Variant
Private iSequenciaFase As Integer

Public Property Get Anotacao() As String

    Anotacao = sAnotacao

End Property

Public Property Let Anotacao(ByVal sNewValue As String)

    sAnotacao = sNewValue

End Property
Public Function ExportaListaExcel()

On Error GoTo ErrorHandler:
    
    Dim objExcel                 As Object
    Dim sDocExcel                As String
    
    Dim objMensagem              As clsgenMensagem
    Dim nLinha                   As Integer
    Dim i                        As Integer
    
    
    Screen.MousePointer = vbArrowHourglass
    
    ' Intância a classe de mensagem
    Set objMensagem = New clsgenMensagem

    ' Instancia o objeto do Word
    Set objExcel = CreateObject("Excel.Application")
    
    'Set objFormulario = frmgenSelecionarOTS
    
    sDocExcel = App.Path & "\RelacaoOTS.xlt"
    
    nLinha = 12

    ' Prepara os campos a serem preechidos
    With objExcel

        .Visible = True
        .Application.WindowState = 1
        .Workbooks.Open sDocExcel, , True
        .CommandBars("Standard").Visible = True
        
        .Cells(1, 7) = "Data: " & Format(Date, "dd/mm/yyyy")
        
        If objFormSelecionarOTS.CboServico.Text <> "" Then
            .Cells(7, 2) = objFormSelecionarOTS.CboServico.Text
        Else
            .Cells(7, 2) = "[Todos]"
        End If
        
        
        .Cells(8, 2) = objFormSelecionarOTS.Caption
        
        
        .Cells(9, 2) = objFormSelecionarOTS.dtpDataIni & " à " & objFormSelecionarOTS.dtpDataFim
        
        
        .Cells(9, 7) = objFormSelecionarOTS.cboCentroFunc.Text
        
        
        For i = 1 To objFormSelecionarOTS.spdCircuitos.MaxRows
        
            'Posiciona na linha do Grid
            objFormSelecionarOTS.spdCircuitos.Row = i
            
            'Captura na coluna do grid Acao
            objFormSelecionarOTS.spdCircuitos.Col = 1
            .Cells(nLinha, 1) = objFormSelecionarOTS.spdCircuitos.Text
            
            'Captura na coluna do grid Acao
            objFormSelecionarOTS.spdCircuitos.Col = 2
            .Cells(nLinha, 2) = objFormSelecionarOTS.spdCircuitos.Text
            
            'Captura na coluna do grid Cliente
            objFormSelecionarOTS.spdCircuitos.Col = 5
            .Cells(nLinha, 3) = objFormSelecionarOTS.spdCircuitos.Text
            
            'Captura na coluna do grid OTS
            objFormSelecionarOTS.spdCircuitos.Col = 3
            .Cells(nLinha, 4) = objFormSelecionarOTS.spdCircuitos.Text
            
            'Captura na coluna do grid Designação
            objFormSelecionarOTS.spdCircuitos.Col = 4
            .Cells(nLinha, 5) = objFormSelecionarOTS.spdCircuitos.Text
            
            'Captura na coluna do grid Data Conclusão
            objFormSelecionarOTS.spdCircuitos.Col = 6
            .Cells(nLinha, 6) = objFormSelecionarOTS.spdCircuitos.Text
        
            'Captura na coluna do grid Data Conclusão
            objFormSelecionarOTS.spdCircuitos.Col = 7
            .Cells(nLinha, 7) = objFormSelecionarOTS.spdCircuitos.Text
        
            nLinha = nLinha + 1
        
        Next

    End With

Finalizar:

    Screen.MousePointer = vbDefault
    
    On Error GoTo 0
    
    If objMensagem.TipoMsg <> "" Then
    
        objMensagem.ExibeMensagem
        
    End If
    
    Set objMensagem = Nothing
    Set objExcel = Nothing
    
    Exit Function
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: ExportaListaExcel " & _
                            "da classe clsGenOC_SelecionarOTS"
    GoTo Finalizar

End Function

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

Public Function TelaSelecionarOTS(ByVal psTitulo As String, _
                                  ByVal piSequenciaFase As Integer) As Boolean

On Error GoTo ErrorHandler:

    Dim bRetorno As Boolean
        
    'RN_SGAD_21148YBJ
    Dim vTodos As String
    vTodos = "[Todos]"
     
    Screen.MousePointer = vbArrowHourglass
     
    ' Atualiza a variável que guarda a sequencia
    iSequenciaFase = piSequenciaFase
    
    ' Intância a classe de mensagem
    Set objMensagem = New clsgenMensagem
     
    ' Instância o formulário de seleção de OTS
    
     If Familia_Produto = "D" Then  'PRODUTOS DE DADOS
        
        If CodigoFase = GbCodigoFaseDesignacao Then
        '#################
        'adeil RN-SGAD-14971ZKH 28/10/2003
        'se o form passado por parâmetro já estiver aberto, fecha-lo.
        'inicio
            If Verifica_Janela_Aberta("FrmDNI_SelecionarOTSporRegiao", "Produtos de Dados - Selecionar OTS" & psTitulo) Then
                GoTo Finalizar
            End If
            Set objFormSelecionarOTS = New FrmDNI_SelecionarOTSporRegiao

        Else
            If Verifica_Janela_Aberta("frmgenSelecionarOTS", "Produtos de Dados - Selecionar OTS" & psTitulo) Then
                GoTo Finalizar
            End If
            Set objFormSelecionarOTS = New frmgenSelecionarOTS
            
        End If
    'fim
       
       objFormSelecionarOTS.Caption = "Produtos de Dados - Selecionar OTS" & psTitulo

    Else                            'PRODUTOS DE INTERNET
    
       
        '#################
        'adeil RN-SGAD-14971ZKH 28/10/2003
        'se o form passado por parâmetro já estiver aberto, fecha-lo.
        'inicio
        If Verifica_Janela_Aberta("frmgenSelecionarOTS", "Produtos de Internet/IP - Selecionar OTS" & psTitulo) Then
            GoTo Finalizar
        End If
       
        Set objFormSelecionarOTS = New frmgenSelecionarOTS
        objFormSelecionarOTS.Caption = "Produtos de Internet/IP - Selecionar OTS" & psTitulo
    
    End If
   
    ' Define período
    objFormSelecionarOTS.dtpDataIni.Value = Format(Now - 10, "dd/mm/yyyy")
    objFormSelecionarOTS.dtpDataFim.Value = Format(Now, "dd/mm/yyyy")
    
    ' Executar o método de exibição dos centros funcionais
    bRetorno = ExibirListaCentroFuncional(objFormSelecionarOTS.cboCentroFunc, CodigoFase)
    
    If Not bRetorno Then
        
        Unload objFormSelecionarOTS
        
        GoTo Finalizar
        
    End If
    
    'RN_SGAD_21148YBJ
    If (GbFlagRN_SGAD_21148YBJ) Then
        
        objFormSelecionarOTS.cboCentroFunc.AddItem vTodos
        objFormSelecionarOTS.cboCentroFunc.ItemData(objFormSelecionarOTS.cboCentroFunc.NewIndex) = 0
        
    End If
    
    
    bRetorno = ExibirListaProduto(objFormSelecionarOTS.CboServico, Familia_Produto)
    
    objFormSelecionarOTS.CboServico.AddItem "", 0
    
    If Not bRetorno Then
        
        Unload objFormSelecionarOTS
        
        GoTo Finalizar
        
    End If
     
    objFormSelecionarOTS.cboCentroFunc.ListIndex = 0
        
    ' Informar o objeto controlador
    objFormSelecionarOTS.informa_OC Me
    
'   Call CarregaTituloProduto(objFormSelecionarOTS, psTitulo)
    
    '=====================================================
    '= Autor: Carlos Soares - 26/06/2003                                           =
    '= Ação: Na fase de Ativação ou pré-ativação,deve aparecer a coluna 49 da grid =
    '= Pendência tarefa 04                                                         =
    '=====================================================
    objFormSelecionarOTS.spdCircuitos.Col = 49
    objFormSelecionarOTS.spdCircuitos.ColWidth(49) = 50
    objFormSelecionarOTS.spdCircuitos.ColHidden = True
    
    If Familia_Produto = "I" And _
        (CodigoFase = GbCodigoFasePreAtivacao Or _
            CodigoFase = GbCodigoFaseAtivacao) Then

        objFormSelecionarOTS.spdCircuitos.Col = 49
        objFormSelecionarOTS.spdCircuitos.ColHidden = False

    End If
    
    '=====================================================
    '=====================================================

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
                            "da classe clsGenOC_SelecionarOTS"
    GoTo Finalizar

End Function

Public Function TelaFase_Diginet() As Boolean

On Error GoTo ErrorHandler:

    Dim bRetorno                As Boolean
    Dim sStatus                 As String
    Dim sPermiteExecucao        As String
    Dim sTipoAcessoCCTO         As String
    
    Screen.MousePointer = vbArrowHourglass
     
    'Intância a classe de mensagem
    Set objMensagem = New clsgenMensagem
     
    'Instância a classe de fase para buscar dados do cabeçalho
    Set objFase = New clsgenFase
    
    'Atribui a linha ativada
    objFormSelecionarOTS.spdCircuitos.Row = objFormSelecionarOTS.spdCircuitos.ActiveRow
    
    'Atribui os valores da lista da OTS selecionada nas propriedades da fase
    bRetorno = AtribuiPropriedadesFase
    
    'Atribui o valor do status da OTS
    objFormSelecionarOTS.spdCircuitos.Col = 2
    sStatus = objFormSelecionarOTS.spdCircuitos.Text
    
    objFormSelecionarOTS.spdCircuitos.Col = 7
    objFormSelecionarOTS.spdCircuitos.Row = objFormSelecionarOTS.spdCircuitos.ActiveRow
    
    sStatus = sStatus & " / " & objFormSelecionarOTS.spdCircuitos.Text
    
    'Verifica qual é a sequência da fase para pode instânciar a tela
    
    'Alocacao de Centro Funcional
    If CodigoFase = GbCodigoFaseAlocacao Then
        
        Set objOcAlocarCF = New clsInterOC_AlocarCF

        objOcAlocarCF.InformaOC_SelecionarOTS Me

        objOcAlocarCF.Fase_ID = lFase_Id
        objOcAlocarCF.Servico_ID = lServico_ID
        objOcAlocarCF.Servico_ID_Anterior = lServico_ID_Anterior
        objOcAlocarCF.Sigla_Acao = sSigla_Acao
        objOcAlocarCF.Status_Fase = sStatus
        objOcAlocarCF.Acao_ID = lAcao_ID
        objOcAlocarCF.Produto_ID = Produto_ID
        objOcAlocarCF.DataFinalizacao = sDataFimReal
        objOcAlocarCF.Sequencia_Acao = lSequencia_Acao
        objOcAlocarCF.CentroFuncionalID = objFormSelecionarOTS.cboCentroFunc.ItemData(objFormSelecionarOTS.cboCentroFunc.ListIndex)
        
        objFormSelecionarOTS.spdCircuitos.Col = 43
        objOcAlocarCF.Data_OTS = objFormSelecionarOTS.spdCircuitos.Text

        objFormSelecionarOTS.spdCircuitos.Col = 45
        objOcAlocarCF.Realocacao = objFormSelecionarOTS.spdCircuitos.Text

        objOcAlocarCF.CarregarForm
        
        If Not bRetorno Then

            GoTo Finalizar

        End If
    
    'Aprovisionamento de Acesso
    ElseIf CodigoFase = GbCodigoFaseAcesso Then
        
        'Atribui a linha selecionada
        objFormSelecionarOTS.spdCircuitos.Row = objFormSelecionarOTS.spdCircuitos.ActiveRow
        objFormSelecionarOTS.spdCircuitos.Col = 44
        
        sTipoAcessoCCTO = objFormSelecionarOTS.spdCircuitos.Text
        
        If sTipoAcessoCCTO = "E" Or sTipoAcessoCCTO = "T" Then
        
            Set objOC_AcessoEmbratel = New clsGenOC_TratarAcessoEmbratel
    
            'Método que recebe todos os dados necessários e que
            'efetivamente inicia o trabalho do Controlador
    
            objOC_AcessoEmbratel.InformaOC_SelecionarOTS Me
    
            objOC_AcessoEmbratel.Fase_ID = lFase_Id
            objOC_AcessoEmbratel.Servico_ID = lServico_ID
            objOC_AcessoEmbratel.Servico_ID_Anterior = lServico_ID_Anterior
            objOC_AcessoEmbratel.Sigla_Acao = sSigla_Acao
            objOC_AcessoEmbratel.Status_Fase = sStatus
            objOC_AcessoEmbratel.Acao_ID = lAcao_ID
            objOC_AcessoEmbratel.Produto_ID = Produto_ID
            objOC_AcessoEmbratel.DataFinalizacao = sDataFimReal
            objOC_AcessoEmbratel.Sequencia_Acao = lSequencia_Acao
            objOC_AcessoEmbratel.ContratoID = lContratoID
            objOC_AcessoEmbratel.NrLocalInsta = lNrLocalInsta
            objOC_AcessoEmbratel.AnoOE = lAnoOE
            objOC_AcessoEmbratel.NumeroOE = lNumeroOE
            objOC_AcessoEmbratel.NumeroItemOE = lNumeroITemOE
            objOC_AcessoEmbratel.Familia_Produto = Familia_Produto
            objOC_AcessoEmbratel.CentroFuncionalID = _
    objFormSelecionarOTS.cboCentroFunc.ItemData(objFormSelecionarOTS.cboCentroFunc.ListIndex)
            
            objOC_AcessoEmbratel.PermiteExecucao = IIf(sPermiteExecucao = "", "N", _
                                                         sPermiteExecucao)
            objOC_AcessoEmbratel.CarregarForm
        Else
    
            Set objOC_AcessoCliente = New clsgenOC_TratarAcessoCliente

            objOC_AcessoCliente.InformaOCSelecionarOTS Me

            objOC_AcessoCliente.Fase_ID = lFase_Id
            objOC_AcessoCliente.Servico_ID = lServico_ID
            objOC_AcessoCliente.Servico_ID_Anterior = lServico_ID_Anterior
            objOC_AcessoCliente.Sigla_Acao = sSigla_Acao
            objOC_AcessoCliente.Status_Fase = sStatus
            objOC_AcessoCliente.Acao_ID = lAcao_ID
            objOC_AcessoCliente.Produto_ID = Produto_ID
            objOC_AcessoCliente.DataFinalizacao = sDataFimReal
            objOC_AcessoCliente.Sequencia_Acao = lSequencia_Acao
            objOC_AcessoCliente.ContratoID = lContratoID
            objOC_AcessoCliente.NrLocalInsta = lNrLocalInsta
            objOC_AcessoCliente.AnoOE = lAnoOE
            objOC_AcessoCliente.NumeroOE = lNumeroOE
            objOC_AcessoCliente.NumeroItemOE = lNumeroITemOE
            objOC_AcessoCliente.Familia_Produto = Familia_Produto
            objOC_AcessoCliente.CentroFuncionalID = _
objFormSelecionarOTS.cboCentroFunc.ItemData(objFormSelecionarOTS.cboCentroFunc.ListIndex)
            
            objOC_AcessoCliente.PermiteExecucao = IIf(sPermiteExecucao = "", "N", _
                                                                   sPermiteExecucao)
            objOC_AcessoCliente.CarregarForm
        
        End If
        
        If Not bRetorno Then
    
            GoTo Finalizar
        
        End If
    
    'Designacao de Circuito
    ElseIf CodigoFase = GbCodigoFaseDesignacao Then
        
        ' Declara o objeto controle de Designacao de Circuito
        Set objOC_Designacao = New clsDNIOC_DesignacaoCCTO
        
        objOC_Designacao.Fase_ID = lFase_Id
        objOC_Designacao.Servico_ID = lServico_ID
        objOC_Designacao.Servico_ID_Anterior = lServico_ID_Anterior
        objOC_Designacao.Status_Fase = sStatus
        objOC_Designacao.Sigla_Acao = sSigla_Acao
        objOC_Designacao.Acao_ID = lAcao_ID
        objOC_Designacao.Produto_ID = Produto_ID
        objOC_Designacao.DataFinalizacao = sDataFimReal
        objOC_Designacao.Sequencia_Acao = lSequencia_Acao
        objOC_Designacao.InformaOC_SelecionarOTS Me
        objOC_Designacao.AnoOE = lAnoOE
        objOC_Designacao.NumeroOE = lNumeroOE
        objOC_Designacao.NumeroItemOE = lNumeroITemOE
        objOC_Designacao.CentroFuncionalID = objFormSelecionarOTS.cboCentroFunc.ItemData(objFormSelecionarOTS.cboCentroFunc.ListIndex)
        
        objOC_Designacao.CarregarForm

        If Not bRetorno Then
    
            GoTo Finalizar
        
        End If
    'Configuracao de Rede E1
    ElseIf CodigoFase = GbCodigoFaseConfigE1 Then
        
        ' Declara o objeto controle de Configuracao de Rede E1
        Dim objOC_ConfiguracaoRedeE1 As New clsDNIOC_ConfigRedeE1
        
        objOC_ConfiguracaoRedeE1.Fase_ID = lFase_Id
        objOC_ConfiguracaoRedeE1.Servico_ID = lServico_ID
        objOC_ConfiguracaoRedeE1.Servico_ID_Anterior = lServico_ID_Anterior
        objOC_ConfiguracaoRedeE1.Status_Fase = sStatus
        objOC_ConfiguracaoRedeE1.Sigla_Acao = sSigla_Acao
        objOC_ConfiguracaoRedeE1.Acao_ID = lAcao_ID
        objOC_ConfiguracaoRedeE1.Produto_ID = Produto_ID
        objOC_ConfiguracaoRedeE1.DataFinalizacao = sDataFimReal
        objOC_ConfiguracaoRedeE1.Sequencia_Acao = lSequencia_Acao
        objOC_ConfiguracaoRedeE1.InformaOC_SelecionarOTS Me
        objOC_ConfiguracaoRedeE1.AnoOE = lAnoOE
        objOC_ConfiguracaoRedeE1.NumeroOE = lNumeroOE
        objOC_ConfiguracaoRedeE1.NumeroItemOE = lNumeroITemOE
        objOC_ConfiguracaoRedeE1.CentroFuncionalID = objFormSelecionarOTS.cboCentroFunc.ItemData(objFormSelecionarOTS.cboCentroFunc.ListIndex)
        
        objOC_ConfiguracaoRedeE1.CarregarForm

        If Not bRetorno Then
    
            GoTo Finalizar
        
        End If
    
    'Configuracao de Rede Basica
    ElseIf CodigoFase = GbCodigoFaseConfigRB Then
        
        ' Declara o objeto controle de Configuracao de Rede E1
        Dim objOC_ConfiguracaoRB As New clsDNIOC_ConfigRedeBasica
        
        objOC_ConfiguracaoRB.Fase_ID = lFase_Id
        objOC_ConfiguracaoRB.Servico_ID = lServico_ID
        objOC_ConfiguracaoRB.Servico_ID_Anterior = lServico_ID_Anterior
        objOC_ConfiguracaoRB.Status_Fase = sStatus
        objOC_ConfiguracaoRB.Sigla_Acao = sSigla_Acao
        objOC_ConfiguracaoRB.Acao_ID = lAcao_ID
        objOC_ConfiguracaoRB.Produto_ID = Produto_ID
        objOC_ConfiguracaoRB.DataFinalizacao = sDataFimReal
        objOC_ConfiguracaoRB.Sequencia_Acao = lSequencia_Acao
        objOC_ConfiguracaoRB.InformaOC_SelecionarOTS Me
        objOC_ConfiguracaoRB.AnoOE = lAnoOE
        objOC_ConfiguracaoRB.NumeroOE = lNumeroOE
        objOC_ConfiguracaoRB.NumeroItemOE = lNumeroITemOE
        objOC_ConfiguracaoRB.CentroFuncionalID = objFormSelecionarOTS.cboCentroFunc.ItemData(objFormSelecionarOTS.cboCentroFunc.ListIndex)
        
        objOC_ConfiguracaoRB.CarregarForm

        If Not bRetorno Then
    
            GoTo Finalizar
        
        End If
    
    ElseIf CodigoFase = GbCodigoFaseTimeSlotDD Then
        
        ' Declara o objeto controle de Configuracao de Rede E1
        Set objOC_TimeSlotDD = New clsDniOC_TimeSlotDedicado
        
        objOC_TimeSlotDD.Fase_ID = lFase_Id
        objOC_TimeSlotDD.Servico_ID = lServico_ID
        objOC_TimeSlotDD.Servico_ID_Anterior = lServico_ID_Anterior
        objOC_TimeSlotDD.Status_Fase = sStatus
        objOC_TimeSlotDD.Sigla_Acao = sSigla_Acao
        objOC_TimeSlotDD.Acao_ID = lAcao_ID
        objOC_TimeSlotDD.Produto_ID = Produto_ID
        objOC_TimeSlotDD.DataFinalizacao = sDataFimReal
        objOC_TimeSlotDD.Sequencia_Acao = lSequencia_Acao
        objOC_TimeSlotDD.CircuitoID = lCircuitoID
        objOC_TimeSlotDD.InformaOC_SelecionarOTS Me
        objOC_TimeSlotDD.AnoOE = lAnoOE
        objOC_TimeSlotDD.NumeroOE = lNumeroOE
        objOC_TimeSlotDD.NumeroItemOE = lNumeroITemOE
        objOC_TimeSlotDD.CentroFuncionalID = objFormSelecionarOTS.cboCentroFunc.ItemData(objFormSelecionarOTS.cboCentroFunc.ListIndex)
        
        objOC_TimeSlotDD.CarregarForm

        If Not bRetorno Then
    
            GoTo Finalizar
        
        End If
    
    'Preparar Centro Local
    ElseIf CodigoFase = GbCodigoPrepararCL Then
        
        ' Declara o objeto controle de Preparacao de Centro Local
        Dim objOC_PrepararCL  As New clsDNIOC_PrepararCentroLocal
        
        objOC_PrepararCL.Fase_ID = lFase_Id
        objOC_PrepararCL.Servico_ID = lServico_ID
        objOC_PrepararCL.Servico_ID_Anterior = lServico_ID_Anterior
        objOC_PrepararCL.Status_Fase = sStatus
        objOC_PrepararCL.Sigla_Acao = sSigla_Acao
        objOC_PrepararCL.Acao_ID = lAcao_ID
        objOC_PrepararCL.Produto_ID = Produto_ID
        objOC_PrepararCL.DataFinalizacao = sDataFimReal
        objOC_PrepararCL.Sequencia_Acao = lSequencia_Acao
        objOC_PrepararCL.ContratoID = lContratoID
        objOC_PrepararCL.InformaOC_SelecionarOTS Me
        objOC_PrepararCL.AnoOE = lAnoOE
        objOC_PrepararCL.NumeroOE = lNumeroOE
        objOC_PrepararCL.NumeroItemOE = lNumeroITemOE
        objOC_PrepararCL.CentroFuncionalID = objFormSelecionarOTS.cboCentroFunc.ItemData(objFormSelecionarOTS.cboCentroFunc.ListIndex)
        
        objOC_PrepararCL.CarregarForm

        If Not bRetorno Then
    
            GoTo Finalizar
        
        End If
    
    'Data de Testes Internacionais
    ElseIf CodigoFase = GbCodigoFaseDataTestes Then
        
        ' Declara o objeto controle de Data de Testes Internacionais
        Dim objOC_DataTestes         As New clsDNIOC_DataTestesInternac
        
        objOC_DataTestes.Fase_ID = lFase_Id
        objOC_DataTestes.Servico_ID = lServico_ID
        objOC_DataTestes.Servico_ID_Anterior = lServico_ID_Anterior
        objOC_DataTestes.Status_Fase = sStatus
        objOC_DataTestes.Sigla_Acao = sSigla_Acao
        objOC_DataTestes.Acao_ID = lAcao_ID
        objOC_DataTestes.Produto_ID = Produto_ID
        objOC_DataTestes.DataFinalizacao = sDataFimReal
        objOC_DataTestes.Sequencia_Acao = lSequencia_Acao
        objOC_DataTestes.ContratoID = lContratoID
        objOC_DataTestes.InformaOC_SelecionarOTS Me
        objOC_DataTestes.AnoOE = lAnoOE
        objOC_DataTestes.NumeroOE = lNumeroOE
        objOC_DataTestes.NumeroItemOE = lNumeroITemOE
        objOC_DataTestes.CentroFuncionalID = objFormSelecionarOTS.cboCentroFunc.ItemData(objFormSelecionarOTS.cboCentroFunc.ListIndex)
        
        objOC_DataTestes.CarregarForm
        
        If Not bRetorno Then
    
            GoTo Finalizar
        
        End If
    
    'Pre-Ativação
    ElseIf CodigoFase = GbCodigoFasePreAtivacao Then
        
        ' Declara o objeto controle de Pre-Ativacao
        Dim objOC_PreAtivacao        As New clsDNIOC_PreAtivacao
        
        objOC_PreAtivacao.Fase_ID = lFase_Id
        objOC_PreAtivacao.Servico_ID = lServico_ID
        objOC_PreAtivacao.Servico_ID_Anterior = lServico_ID_Anterior
        objOC_PreAtivacao.Status_Fase = sStatus
        objOC_PreAtivacao.Sigla_Acao = sSigla_Acao
        objOC_PreAtivacao.Acao_ID = lAcao_ID
        objOC_PreAtivacao.Produto_ID = Produto_ID
        objOC_PreAtivacao.DataFinalizacao = sDataFimReal
        objOC_PreAtivacao.Sequencia_Acao = lSequencia_Acao
        objOC_PreAtivacao.InformaOC_SelecionarOTS Me
        objOC_PreAtivacao.AnoOE = lAnoOE
        objOC_PreAtivacao.NumeroOE = lNumeroOE
        objOC_PreAtivacao.NumeroItemOE = lNumeroITemOE
        objOC_PreAtivacao.CentroFuncionalID = objFormSelecionarOTS.cboCentroFunc.ItemData(objFormSelecionarOTS.cboCentroFunc.ListIndex)
        
        objOC_PreAtivacao.CarregarForm
        
        If Not bRetorno Then
    
            GoTo Finalizar
        
        End If
    
    'Programar CPE
    ElseIf CodigoFase = GbCodigoFaseProgramarCPE Then
        
        ' Declara o objeto controle Programar CPE
        Dim objOC_ProgramarCPE       As New clsDNIOC_ProgramarCPE
        objOC_ProgramarCPE.Fase_ID = lFase_Id
        objOC_ProgramarCPE.Servico_ID = lServico_ID
        objOC_ProgramarCPE.Servico_ID_Anterior = lServico_ID_Anterior
        objOC_ProgramarCPE.Status_Fase = sStatus
        objOC_ProgramarCPE.Sigla_Acao = sSigla_Acao
        objOC_ProgramarCPE.Acao_ID = lAcao_ID
        objOC_ProgramarCPE.Produto_ID = Produto_ID
        objOC_ProgramarCPE.DataFinalizacao = sDataFimReal
        objOC_ProgramarCPE.Sequencia_Acao = lSequencia_Acao
        objOC_ProgramarCPE.ContratoID = lContratoID
        objOC_ProgramarCPE.progChamador = "SELECIONAR"
        
        objOC_ProgramarCPE.AnoOE = lAnoOE
        objOC_ProgramarCPE.NumeroOE = lNumeroOE
        objOC_ProgramarCPE.NumeroItemOE = lNumeroITemOE
        
        objOC_ProgramarCPE.InformaOC_SelecionarOTS Me
        objOC_ProgramarCPE.CentroFuncionalID = objFormSelecionarOTS.cboCentroFunc.ItemData(objFormSelecionarOTS.cboCentroFunc.ListIndex)
        
        objOC_ProgramarCPE.CarregarForm

        If Not bRetorno Then
    
            GoTo Finalizar
        
        End If
    
    'Ativacao
    ElseIf CodigoFase = GbCodigoFaseAtivacao Then
        
        ' Declara o objeto controle de Data de Testes Internacionais
        Dim objOC_AtivacaoCCTO       As New clsDNIOC_AtivacaoCCTO
        
        objOC_AtivacaoCCTO.Fase_ID = lFase_Id
        objOC_AtivacaoCCTO.Servico_ID = lServico_ID
        objOC_AtivacaoCCTO.Servico_ID_Anterior = lServico_ID_Anterior
        objOC_AtivacaoCCTO.Status_Fase = sStatus
        objOC_AtivacaoCCTO.Sigla_Acao = sSigla_Acao
        objOC_AtivacaoCCTO.Acao_ID = lAcao_ID
        objOC_AtivacaoCCTO.Produto_ID = Produto_ID
        objOC_AtivacaoCCTO.DataInicioPrevisto = sDataInicioPrevisto
        objOC_AtivacaoCCTO.DataFinalizacao = sDataFimReal
        objOC_AtivacaoCCTO.Sequencia_Acao = lSequencia_Acao
        
        objOC_AtivacaoCCTO.AnoOE = lAnoOE
        objOC_AtivacaoCCTO.NumeroOE = lNumeroOE
        objOC_AtivacaoCCTO.NumeroItemOE = lNumeroITemOE
        objOC_AtivacaoCCTO.ContratoID = lContratoID
        
        objOC_AtivacaoCCTO.InformaOC_SelecionarOTS Me
        objOC_AtivacaoCCTO.CentroFuncionalID = objFormSelecionarOTS.cboCentroFunc.ItemData(objFormSelecionarOTS.cboCentroFunc.ListIndex)
        
        objOC_AtivacaoCCTO.CarregarForm

        If Not bRetorno Then
    
            GoTo Finalizar
        
        End If
    End If
    
    GoTo Finalizar
       
Finalizar:

    On Error GoTo 0
    
    Screen.MousePointer = vbDefault
    
    If objMensagem.TipoMsg <> "" Then
    
        objMensagem.ExibeMensagem
        
    End If
    
    Set objMensagem = Nothing
    
    Exit Function
        
ErrorHandler:


    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: TelaFase_Diginet " & _
                            "da classe clsGenOC_SelecionarOTS"
    GoTo Finalizar

End Function
Public Function TelaFase_BlFlex() As Boolean

On Error GoTo ErrorHandler:

    Dim bRetorno                As Boolean
    Dim sStatus                 As String
    Dim lIDServico              As Long
    Dim lIDRedundante           As Long
    Dim vDesigOTS()             As Variant
    Dim objdadosServico         As clsGenServico
    Dim sCNL                    As String
    Dim lNumeroOTS              As Long
    Dim iAnoOTS                 As Integer
    Dim sPermiteExecucao        As String
    Dim vsTipoAcesso            As String
    Dim sCentroFuncional        As String
    Dim objDadosTecnicos        As Object
    Dim sNumeroContrato         As String
    Dim sAnoContrato            As String
    Dim sVersaoContrato         As String
    Dim sCodigoProduto          As String
    
    Dim sPeer_BGP               As String
    
    'RN_SGAD_21148YBJ
    Dim lCentroFuncionalID      As Long
    
    'JAH
    Dim objContrato             As clsgenDadosContrato
    Dim sTabela_Contrato        As String
    
    Screen.MousePointer = vbArrowHourglass
    
    'Intância a classe de mensagem
    Set objMensagem = New clsgenMensagem

    'Instância a classe de fase para buscar dados do cabeçalho
    Set objFase = New clsgenFase

    'Atribui os valores da lista da OTS selecionada nas propriedades da fase
    bRetorno = AtribuiPropriedadesFase

    'Atribui a linha ativada
    objFormSelecionarOTS.spdCircuitos.Row = objFormSelecionarOTS.spdCircuitos.ActiveRow

    'Atribui o valor do status da OTS

    objFormSelecionarOTS.spdCircuitos.Col = 2
    sStatus = objFormSelecionarOTS.spdCircuitos.Text

    'RN_SGAD_21148YBJ
    If (GbFlagRN_SGAD_21148YBJ) Then
    
        objFormSelecionarOTS.spdCircuitos.Row = objFormSelecionarOTS.spdCircuitos.ActiveRow
        objFormSelecionarOTS.spdCircuitos.Col = 52
        sCentroFuncional = objFormSelecionarOTS.spdCircuitos.Text
    
    Else
    
        sCentroFuncional = objFormSelecionarOTS.cboCentroFunc.Text
    
    End If


    objFormSelecionarOTS.spdCircuitos.Col = 7
    objFormSelecionarOTS.spdCircuitos.Row = objFormSelecionarOTS.spdCircuitos.ActiveRow

    sStatus = sStatus & " / " & objFormSelecionarOTS.spdCircuitos.Text

    'Monta dados do Cabecalho
    If objBlfCabecalho Is Nothing Then

        Set objBlfCabecalho = New clsBlf_Cabecalho

    End If

    objFormSelecionarOTS.spdCircuitos.Col = 28
    
    Select Case objFormSelecionarOTS.spdCircuitos.Text
           Case "ATV"
                objBlfCabecalho.Acao = "Ativação"
           Case "ALT"
                objBlfCabecalho.Acao = "Alteração"
           Case "ATG"
                objBlfCabecalho.Acao = "Correção de Ativação"
           Case "ALG"
                objBlfCabecalho.Acao = "Alteração de Ativação"
           Case "DES"
                objBlfCabecalho.Acao = "Desativação"
           Case "CAN"
                objBlfCabecalho.Acao = "Cancelamento"
    End Select

    Set objdadosServico = New clsGenServico

    objdadosServico.InformaBanco frmgenMDI_SGAPlus.objBanco
    objdadosServico.InformaMensagem objMensagem
    
    'Resgata os dados do Serviço
    bRetorno = objdadosServico.RecuperarDados(lServico_ID)
    
    If (objMensagem.TipoMsg <> "") Then

       GoTo ErrorHandler

    End If
    
    
    objFormSelecionarOTS.spdCircuitos.Col = 22
    
    'IP VPN
    If (lProduto_ID <> GbIDProdutoBIV) Then
    
        Select Case objFormSelecionarOTS.spdCircuitos.Text
                Case GbTIERED
                    objBlfCabecalho.TipoCircuito = "TIERED"
                Case GbBURSTABLE
                    objBlfCabecalho.TipoCircuito = "BURSTABLE"
                Case GbSHADOW
                    objBlfCabecalho.TipoCircuito = "SHADOW"
                Case GbCOMUM
                    objBlfCabecalho.TipoCircuito = "COMUM"
                Case GbDIVERSE
                    objBlfCabecalho.TipoCircuito = "DIVERSE"
                Case GbDOUBLE
                    objBlfCabecalho.TipoCircuito = "DOUBLE"
        End Select
    End If
    
    '#BLD
    If (lProduto_ID = GbIDProdutoBLD Or _
            lProduto_ID = GbIDProdutoVOIP) Then
    
        objBlfCabecalho.TipoCircuito = "COMUM"
        
    End If
    
    objFormSelecionarOTS.spdCircuitos.Col = 9
    lIDServico = CLng("0" & objFormSelecionarOTS.spdCircuitos.Text)
    
    objFormSelecionarOTS.spdCircuitos.Col = 10
    lIDRedundante = CLng("0" & objFormSelecionarOTS.spdCircuitos.Text)
    
    If lIDRedundante > 0 Then
    
        objBlfCabecalho.OTSRedundante = True
        
    Else
    
        objBlfCabecalho.OTSRedundante = False
        
    End If
    
    bRetorno = objdadosServico.DevolveDesignacaoOTS(lIDServico, lProduto_ID, True)

    If objMensagem.TipoMsg <> "" Then

       GoTo ErrorHandler

    End If
    
    objBlfCabecalho.OtsPrincipal = objdadosServico.OTS
    objBlfCabecalho.DesignacaoPrincipal = objdadosServico.DesignacaoCircuito
        

    objFormSelecionarOTS.spdCircuitos.Col = 39
    sPermiteExecucao = objFormSelecionarOTS.spdCircuitos.Text

'    objFormSelecionarOTS.spdCircuitos.Col = 10
'    lIDServico = CLng("0" & objFormSelecionarOTS.spdCircuitos.Text)
    
    bRetorno = objdadosServico.DevolveDesignacaoOTS(lIDRedundante, lProduto_ID, True)

    If objMensagem.TipoMsg <> "" Then

       GoTo ErrorHandler

    End If

    objBlfCabecalho.OTSAssociada = objdadosServico.OTS
    
    objBlfCabecalho.DesignacaoSecundaria = objdadosServico.DesignacaoCircuito
    
    objFormSelecionarOTS.spdCircuitos.Col = 22
    
    If (objFormSelecionarOTS.spdCircuitos.Text = GbSHADOW) Or _
       (objFormSelecionarOTS.spdCircuitos.Text = GbDIVERSE) Or _
       (objFormSelecionarOTS.spdCircuitos.Text = GbDOUBLE) Then
    
        If lIDRedundante = 0 Then
        
            ' Buscar o ID principal
            bRetorno = objdadosServico.BuscarPrincipal(lIDServico, _
                                                       lIDRedundante)
    
            If objMensagem.TipoMsg <> "" Then
        
               GoTo ErrorHandler
        
            End If
            
            bRetorno = objdadosServico.DevolveDesignacaoOTS(lIDRedundante, lProduto_ID, True)
        
            If objMensagem.TipoMsg <> "" Then
        
               GoTo ErrorHandler
        
            End If
        
            objBlfCabecalho.OTSAssociada = objdadosServico.OTS
            
            objBlfCabecalho.DesignacaoSecundaria = objdadosServico.DesignacaoCircuito
            
        End If
        
    End If
        
    'OBTEM A OTS ANTERIOR, CASO EXISTA
    
    objFormSelecionarOTS.spdCircuitos.Col = 11
    
    lIDServico = CLng("0" & objFormSelecionarOTS.spdCircuitos.Text)
    
    bRetorno = objdadosServico.DevolveDesignacaoOTS(lIDServico, lProduto_ID, True)

    If objMensagem.TipoMsg <> "" Then

       GoTo ErrorHandler

    End If

    objBlfCabecalho.OtsAnterior = objdadosServico.OTS
    
    
    objFormSelecionarOTS.spdCircuitos.Col = 37
    objBlfCabecalho.ContaCorrente = objFormSelecionarOTS.spdCircuitos.Text
    
    objFormSelecionarOTS.spdCircuitos.Col = 38
    objBlfCabecalho.NomeCliente = objFormSelecionarOTS.spdCircuitos.Text
    
    'Nome Guerra. IP VPN
    objFormSelecionarOTS.spdCircuitos.Col = 50
    objBlfCabecalho.NomeFantasia = objFormSelecionarOTS.spdCircuitos.Text
   
    '
    'objBlfCabecalho.NomeFantasia = objFormSelecionarOTS.spdCircuitos.Text
              
    objFormSelecionarOTS.spdCircuitos.Col = 16

    If objFormSelecionarOTS.spdCircuitos.Text <> "" Then

       objBlfCabecalho.DataFinalizacao = Format(objFormSelecionarOTS.spdCircuitos.Text, "dd/mm/yyyy")

    Else

        objBlfCabecalho.DataFinalizacao = ""

    End If

    objFormSelecionarOTS.spdCircuitos.Col = 17
    objBlfCabecalho.NumeroOE = objFormSelecionarOTS.spdCircuitos.Text

    objFormSelecionarOTS.spdCircuitos.Col = 19
    objBlfCabecalho.AnoOE = objFormSelecionarOTS.spdCircuitos.Text

    objFormSelecionarOTS.spdCircuitos.Col = 20
    objBlfCabecalho.ItemOE = objFormSelecionarOTS.spdCircuitos.Text

    objFormSelecionarOTS.spdCircuitos.Col = 40
    sNumeroContrato = objFormSelecionarOTS.spdCircuitos.Text
    
    objFormSelecionarOTS.spdCircuitos.Col = 41
    sAnoContrato = objFormSelecionarOTS.spdCircuitos.Text

    objFormSelecionarOTS.spdCircuitos.Col = 42
    sVersaoContrato = objFormSelecionarOTS.spdCircuitos.Text

    objBlfCabecalho.Contrato = sNumeroContrato & "/" & sAnoContrato & "-" & sVersaoContrato

    objFormSelecionarOTS.spdCircuitos.Col = 33
    objBlfCabecalho.CNL = objFormSelecionarOTS.spdCircuitos.Text
    
    
    'IP VPN e JAH
    objFormSelecionarOTS.spdCircuitos.Col = 6
    objBlfCabecalho.PrevConclusaoFase = objFormSelecionarOTS.spdCircuitos.Text
    
    
    objFormSelecionarOTS.spdCircuitos.Col = 48
    objBlfCabecalho.PrevATVContrato = objFormSelecionarOTS.spdCircuitos.Text
    
    'Tipo de Item. IP VPN
    objFormSelecionarOTS.spdCircuitos.Col = 47
    If objFormSelecionarOTS.spdCircuitos.Text = GbItem_SITE Then
        
        objBlfCabecalho.TipoItem = "Site"
        
    ElseIf objFormSelecionarOTS.spdCircuitos.Text = GbItem_TOPOLOGIA Then
        
        objBlfCabecalho.TipoItem = "Topologia"
    
    ElseIf objFormSelecionarOTS.spdCircuitos.Text = GbItem_EXTRNET Then
        
        objBlfCabecalho.TipoItem = "Extranet"
    
    'JAH
    ElseIf objFormSelecionarOTS.spdCircuitos.Text = GbDADOSVOZ Then
        
        objBlfCabecalho.TipoItem = "VOZ"
        objBlfCabecalho.TipoItem_Codigo = GbDADOSVOZ
    
    ElseIf objFormSelecionarOTS.spdCircuitos.Text = GbDADOSBANDALARGA Then
        
        objBlfCabecalho.TipoItem = "BANDA LARGA"
        objBlfCabecalho.TipoItem_Codigo = GbDADOSBANDALARGA
    
    ElseIf objFormSelecionarOTS.spdCircuitos.Text = GbDADOSVOZ_BANDALARGA Then
        
        objBlfCabecalho.TipoItem = "VOZ E BANDA LARGA"
        objBlfCabecalho.TipoItem_Codigo = GbDADOSVOZ_BANDALARGA
    
    End If
    
    
    'RN_SGAD_21148YBJ
    'Obtém o Centro Funcional ID
    If (GbFlagRN_SGAD_21148YBJ) Then

        'MTLP: inclusão do ID do Centro Funcional da Fase
        'objFormSelecionarOTS.spdCircuitos.Col = 51 'coluna retorna o ID do CF da OTS
        objFormSelecionarOTS.spdCircuitos.Col = 53  'coluna retorna o ID do CF da Fase
        lCentroFuncionalID = objFormSelecionarOTS.spdCircuitos.Text

    Else
    
        lCentroFuncionalID = objFormSelecionarOTS.cboCentroFunc.ItemData(objFormSelecionarOTS.cboCentroFunc.ListIndex)
    
    End If
    
    
    ' Instância a classe de dados técnicos
    If (Produto_ID = GbIDProdutoBLF) Then
        
        ' BLFlex
        Set objDadosTecnicos = New clsblfDadosTecnicos
        
    'IP VPN
    ElseIf (Produto_ID = GbIDProdutoBIV) Then
        
        ' BLFlex
        Set objDadosTecnicos = New clsivnDadosTecnicos
        
        sCodigoProduto = GbCodigo_Produto_IPVPN
    
    'JAH
    ElseIf (Produto_ID = GbIDProdutoJAH) Then
        
        Set objDadosTecnicos = New clsjahDadosTecnicos
        
        sCodigoProduto = GbCodigo_Produto_JAH
    
    
    ElseIf (Produto_ID = GbIDProdutoWID) Or _
           (Produto_ID = GbIDProdutoVPN_RA) Or _
           (Produto_ID = GbIDProdutoVPN_DIRECT) Then
    
        ' Dedicated
        Set objDadosTecnicos = New clswidDadosTecnicos
    
    ElseIf (Produto_ID = GbIDProdutoBLD Or _
                lProduto_ID = GbIDProdutoVOIP) Then 'BLD
    
        Set objDadosTecnicos = New clsbldDadosTecnicos
    
    End If

    objDadosTecnicos.InformaBanco frmgenMDI_SGAPlus.objBanco
    objDadosTecnicos.InformaMensagem objMensagem
    
    'Obtém os dados técnicos do produto IP VPN
    If (Produto_ID = GbIDProdutoBIV) Then
    
        bRetorno = objDadosTecnicos.RecuperarDados_Site(lContratoID, _
                                lNrLocalInsta, _
                                "" _
                                , sCodigoProduto)
                                
        sPeer_BGP = objDadosTecnicos.IPPeerBPG
    
    'Obtém os dados técnicos do produto JAH
    'e os dados de Contrato - JAH
    
    ElseIf (Produto_ID = GbIDProdutoJAH) Then
    
        'Obtém dados técnicos
        bRetorno = objDadosTecnicos.RecuperarDados(lContratoID, _
                                lNrLocalInsta, _
                                objdadosServico.AcaoSOE, _
                                sCodigoProduto)
                                
        'Trata erro
        If (objMensagem.TipoMsg <> "") Then
    
            GoTo ErrorHandler
        
        End If

        'Dados de contrato
        sTabela_Contrato = "ATEN_CONTR_PADRA"
        
        Set objContrato = New clsgenDadosContrato
        
        objContrato.InformaBanco frmgenMDI_SGAPlus.objBanco
        objContrato.InformaMensagem objMensagem
    
        bRetorno = objContrato.RecuperarDados(lContratoID, _
                            sTabela_Contrato)
                                
                                
        If (Not bRetorno) Then
            
            'Tratar Contrato não obtido
            objMensagem.TipoMsg = "E"
            objMensagem.Descricao = "Dados do Contrato não obtidos no ATEN!"
                                
            GoTo Finalizar
        End If

        'Complementa o cabeçalho com informações técnicas específicas do JAH
        objBlfCabecalho.UserName = objDadosTecnicos.UserName
        objBlfCabecalho.NulTelefone = objDadosTecnicos.NumTelBandaLarga
        
        objBlfCabecalho.TipoPessoa = objContrato.TipoPessoa
        objBlfCabecalho.Cod_Velocidade = objDadosTecnicos.Cod_Velocidade

    ElseIf (Produto_ID <> GbIDProdutoBLD And _
                lProduto_ID <> GbIDProdutoVOIP) Then
        
        bRetorno = objDadosTecnicos.RecuperarDados(lContratoID, _
                    lNrLocalInsta)

    Else
        
        bRetorno = objDadosTecnicos.RecuperarDados(lNumeroOE, _
                        lAnoOE, _
                        1, _
                        lNumeroITemOE)
        
    End If
    
    'Trata erro
    If (objMensagem.TipoMsg <> "") Then

        GoTo ErrorHandler
    
    End If
    
    'BLD
    If (Produto_ID <> GbIDProdutoBLD And _
            lProduto_ID <> GbIDProdutoVOIP) Then

        objBlfCabecalho.Velocidade = objDadosTecnicos.Velocidade_WebIP

    Else

        objBlfCabecalho.Velocidade = objDadosTecnicos.Velocidade

    End If


    Set objDadosTecnicos = Nothing
    
    'Verifica qual é a sequência da fase para pode instânciar a tela

    'Alocacao de Centro Funcional
    If CodigoFase = GbCodigoFaseAlocacao Then

        Set objOcAlocarCF = New clsInterOC_AlocarCF

        objOcAlocarCF.InformaOC_SelecionarOTS Me
 
        objOcAlocarCF.Fase_ID = lFase_Id
        objOcAlocarCF.Servico_ID = lServico_ID
        objOcAlocarCF.Servico_ID_Anterior = lServico_ID_Anterior
        objOcAlocarCF.Sigla_Acao = sSigla_Acao
        objOcAlocarCF.Status_Fase = sStatus
        objOcAlocarCF.Acao_ID = lAcao_ID
        objOcAlocarCF.Produto_ID = Produto_ID
        objOcAlocarCF.DataFinalizacao = sDataFimReal
        objOcAlocarCF.Sequencia_Acao = lSequencia_Acao
        
        'RN_SGAD_21148YBJ
        'objOcAlocarCF.CentroFuncionalID = objFormSelecionarOTS.cboCentroFunc.ItemData(objFormSelecionarOTS.cboCentroFunc.ListIndex)
        objOcAlocarCF.CentroFuncionalID = lCentroFuncionalID
        
        objFormSelecionarOTS.spdCircuitos.Col = 43
        objOcAlocarCF.Data_OTS = objFormSelecionarOTS.spdCircuitos.Text

        objFormSelecionarOTS.spdCircuitos.Col = 45
        objOcAlocarCF.Realocacao = objFormSelecionarOTS.spdCircuitos.Text

        objOcAlocarCF.CarregarForm

    'Aprovisionamento de Acesso
    ElseIf CodigoFase = GbCodigoFaseAcesso Then

            'Dim objocTratarAcesso As clsblfOC_TratarAcesso
            'Set objocTratarAcesso = New clsblfOC_TratarAcesso

            objFormSelecionarOTS.spdCircuitos.Col = 44
            vsTipoAcesso = objFormSelecionarOTS.spdCircuitos.Text

            'Verifica se o Acesso é Embratel ou Cliente

            If vsTipoAcesso = "E" Or vsTipoAcesso = "T" Or vsTipoAcesso = "N" Then

                Set objOC_AcessoEmbratel = New clsGenOC_TratarAcessoEmbratel

                'Método que recebe todos os dados necessários e que
                'efetivamente inicia o trabalho do Controlador

                objOC_AcessoEmbratel.InformaOC_SelecionarOTS Me

                objOC_AcessoEmbratel.Fase_ID = lFase_Id
                objOC_AcessoEmbratel.Servico_ID = lServico_ID
                objOC_AcessoEmbratel.Servico_ID_Anterior = lServico_ID_Anterior
                objOC_AcessoEmbratel.Sigla_Acao = sSigla_Acao
                objOC_AcessoEmbratel.Status_Fase = sStatus
                objOC_AcessoEmbratel.Acao_ID = lAcao_ID
                objOC_AcessoEmbratel.Produto_ID = Produto_ID
                objOC_AcessoEmbratel.DataFinalizacao = sDataFimReal
                objOC_AcessoEmbratel.Sequencia_Acao = lSequencia_Acao
                objOC_AcessoEmbratel.ContratoID = lContratoID
                objOC_AcessoEmbratel.NrLocalInsta = lNrLocalInsta
                objOC_AcessoEmbratel.Familia_Produto = Familia_Produto
                
                'RN_SGAD_21148YBJ
                'objOC_AcessoEmbratel.CentroFuncionalID = objFormSelecionarOTS.cboCentroFunc.ItemData (objFormSelecionarOTS.cboCentroFunc.ListIndex)
                objOC_AcessoEmbratel.CentroFuncionalID = lCentroFuncionalID
                
                
                objOC_AcessoEmbratel.PermiteExecucao = IIf(sPermiteExecucao = "", "N", _
                                                             sPermiteExecucao)
                objOC_AcessoEmbratel.CarregarForm

            ElseIf vsTipoAcesso = "C" Then  'Tipo Acesso Cliente

                Set objOC_AcessoCliente = New clsgenOC_TratarAcessoCliente

                objOC_AcessoCliente.InformaOCSelecionarOTS Me

                objOC_AcessoCliente.Fase_ID = lFase_Id
                objOC_AcessoCliente.Servico_ID = lServico_ID
                objOC_AcessoCliente.Servico_ID_Anterior = lServico_ID_Anterior
                objOC_AcessoCliente.Sigla_Acao = sSigla_Acao
                objOC_AcessoCliente.Status_Fase = sStatus
                objOC_AcessoCliente.Acao_ID = lAcao_ID
                objOC_AcessoCliente.Produto_ID = Produto_ID
                objOC_AcessoCliente.DataFinalizacao = sDataFimReal
                objOC_AcessoCliente.Sequencia_Acao = lSequencia_Acao
                objOC_AcessoCliente.ContratoID = lContratoID
                objOC_AcessoCliente.NrLocalInsta = lNrLocalInsta
                objOC_AcessoCliente.Familia_Produto = Familia_Produto
                
                
                'RN_SGAD_21148YBJ
                'objOC_AcessoCliente.CentroFuncionalID = objFormSelecionarOTS.cboCentroFunc.ItemData(objFormSelecionarOTS.cboCentroFunc.ListIndex)
                objOC_AcessoCliente.CentroFuncionalID = lCentroFuncionalID
                
                objOC_AcessoCliente.PermiteExecucao = IIf(sPermiteExecucao = "", "N", _
                                                                       sPermiteExecucao)
                objOC_AcessoCliente.CarregarForm
            
            '''
            ElseIf (vsTipoAcesso = "D") Then  'Tipo Acesso DSLAM - JAH

                Set objOC_AcessoDSLAM = New clsgenOC_TratarAcessoDSLAM
                
                objOC_AcessoDSLAM.InformaOCSelecionarOTS Me

                objOC_AcessoDSLAM.Fase_ID = lFase_Id
                objOC_AcessoDSLAM.Servico_ID = lServico_ID
                objOC_AcessoDSLAM.Servico_ID_Anterior = lServico_ID_Anterior
                objOC_AcessoDSLAM.Sigla_Acao = sSigla_Acao
                objOC_AcessoDSLAM.Status_Fase = sStatus
                objOC_AcessoDSLAM.Acao_ID = lAcao_ID
                objOC_AcessoDSLAM.Produto_ID = Produto_ID
                objOC_AcessoDSLAM.DataFinalizacao = sDataFimReal
                objOC_AcessoDSLAM.Sequencia_Acao = lSequencia_Acao
                objOC_AcessoDSLAM.ContratoID = lContratoID
                objOC_AcessoDSLAM.NrLocalInsta = lNrLocalInsta
                objOC_AcessoDSLAM.Familia_Produto = Familia_Produto
                
                'RN_SGAD_21148YBJ
                objOC_AcessoDSLAM.CentroFuncionalID = lCentroFuncionalID
                
                objOC_AcessoDSLAM.PermiteExecucao = IIf(sPermiteExecucao = "", "N", _
                                                                       sPermiteExecucao)
                objOC_AcessoDSLAM.CarregarForm
            
            Else    'Tipo Acesso não cadastrado na tabela Serviço

                objMensagem.TipoMsg = "E"
                objMensagem.Descricao = "Tipo do acesso inválido para esse circuito"

                GoTo ErrorHandler

            End If

            'E - Acesso Embratel  T - Acesso Terceiros  C - Acesso Cliente

    ElseIf CodigoFase = GbCodigoFaseConfiguracao Then

            Set objocConfiguracao = New clsInterOC_Configuracao

            'Método que recebe todos os dados necessários e que efetivamente inicia o trabalho do Controlador
            objocConfiguracao.InformaOCSelecionarOTS Me

            objocConfiguracao.Fase_ID = lFase_Id
            objocConfiguracao.Servico_ID = lServico_ID
            objocConfiguracao.Servico_ID_Anterior = lServico_ID_Anterior
            objocConfiguracao.Sigla_Acao = sSigla_Acao
            objocConfiguracao.Acao_ID = lAcao_ID
            objocConfiguracao.Status_Fase = sStatus
            objocConfiguracao.Produto_ID = Produto_ID
            objocConfiguracao.DataFinalizacao = sDataFimReal
            objocConfiguracao.Sequencia_Acao = lSequencia_Acao
            objocConfiguracao.ContratoID = lContratoID
            objocConfiguracao.NrLocalInsta = lNrLocalInsta
            
            'RN_SGAD_21148YBJ
            'objocConfiguracao.CentroFuncionalID = objFormSelecionarOTS.cboCentroFunc.ItemData(objFormSelecionarOTS.cboCentroFunc.ListIndex)
            objocConfiguracao.CentroFuncionalID = lCentroFuncionalID
            
            objocConfiguracao.PermiteExecucao = IIf(sPermiteExecucao = "", "N", _
                                                         sPermiteExecucao)
            objocConfiguracao.CarregarForm
    ''
    ElseIf CodigoFase = GbCodigoFaseConfiguracaoNGN Then

            Set objocConfiguracaoTel = New clsInterOC_ConfigNGN

            'Método que recebe todos os dados necessários e que efetivamente inicia o trabalho do Controlador
            objocConfiguracaoTel.InformaOCSelecionarOTS Me

            objocConfiguracaoTel.Fase_ID = lFase_Id
            objocConfiguracaoTel.Servico_ID = lServico_ID
            objocConfiguracaoTel.Servico_ID_Anterior = lServico_ID_Anterior
            objocConfiguracaoTel.Sigla_Acao = sSigla_Acao
            objocConfiguracaoTel.Acao_ID = lAcao_ID
            objocConfiguracaoTel.Status_Fase = sStatus
            objocConfiguracaoTel.Produto_ID = Produto_ID
            objocConfiguracaoTel.DataFinalizacao = sDataFimReal
            objocConfiguracaoTel.Sequencia_Acao = lSequencia_Acao
            objocConfiguracaoTel.ContratoID = lContratoID
            objocConfiguracaoTel.NrLocalInsta = lNrLocalInsta
            
            'RN_SGAD_21148YBJ
            'objocTratarCentroLocal.CentroFuncionalID = objFormSelecionarOTS.cboCentroFunc.ItemData(objFormSelecionarOTS.cboCentroFunc.ListIndex)
            objocConfiguracaoTel.CentroFuncionalID = lCentroFuncionalID
            
            objocConfiguracaoTel.PermiteExecucao = IIf(sPermiteExecucao = "", "N", _
                                                         sPermiteExecucao)
            objocConfiguracaoTel.DescricaoCF = sCentroFuncional
            
            objocConfiguracaoTel.CarregarForm
    
    ElseIf CodigoFase = GbCodigoFaseConfiguracaoConftel Then

            Set objocConfiguracaoTel = New clsInterOC_ConfigNGN

            'Método que recebe todos os dados necessários e que efetivamente inicia o trabalho do Controlador
            objocConfiguracaoTel.InformaOCSelecionarOTS Me

            objocConfiguracaoTel.Fase_ID = lFase_Id
            objocConfiguracaoTel.Servico_ID = lServico_ID
            objocConfiguracaoTel.Servico_ID_Anterior = lServico_ID_Anterior
            objocConfiguracaoTel.Sigla_Acao = sSigla_Acao
            objocConfiguracaoTel.Acao_ID = lAcao_ID
            objocConfiguracaoTel.Status_Fase = sStatus
            objocConfiguracaoTel.Produto_ID = Produto_ID
            objocConfiguracaoTel.DataFinalizacao = sDataFimReal
            objocConfiguracaoTel.Sequencia_Acao = lSequencia_Acao
            objocConfiguracaoTel.ContratoID = lContratoID
            objocConfiguracaoTel.NrLocalInsta = lNrLocalInsta
            
            'RN_SGAD_21148YBJ
            'objocTratarCentroLocal.CentroFuncionalID = objFormSelecionarOTS.cboCentroFunc.ItemData(objFormSelecionarOTS.cboCentroFunc.ListIndex)
            objocConfiguracaoTel.CentroFuncionalID = lCentroFuncionalID
            
            objocConfiguracaoTel.PermiteExecucao = IIf(sPermiteExecucao = "", "N", _
                                                         sPermiteExecucao)
            objocConfiguracaoTel.DescricaoCF = sCentroFuncional
            
            objocConfiguracaoTel.CarregarForm
    
    ElseIf CodigoFase = GbCodigoFaseCentroLocal Then

            Set objocTratarCentroLocal = New clsInterOC_TratarCentroLocal

            'Método que recebe todos os dados necessários e que efetivamente inicia o trabalho do Controlador
            objocTratarCentroLocal.InformaOCSelecionarOTS Me

            objocTratarCentroLocal.Fase_ID = lFase_Id
            objocTratarCentroLocal.Servico_ID = lServico_ID
            objocTratarCentroLocal.Servico_ID_Anterior = lServico_ID_Anterior
            objocTratarCentroLocal.Sigla_Acao = sSigla_Acao
            objocTratarCentroLocal.Acao_ID = lAcao_ID
            objocTratarCentroLocal.Status_Fase = sStatus
            objocTratarCentroLocal.Produto_ID = Produto_ID
            objocTratarCentroLocal.DataFinalizacao = sDataFimReal
            objocTratarCentroLocal.Sequencia_Acao = lSequencia_Acao
            objocTratarCentroLocal.ContratoID = lContratoID
            objocTratarCentroLocal.NrLocalInsta = lNrLocalInsta
            
            'Caso o usuário selecione o Centro Funcional Todos o sistema
            'não possibilita gravação. #BLD 09/07/2004
            If (objFormSelecionarOTS.cboCentroFunc.ItemData(objFormSelecionarOTS.cboCentroFunc.ListIndex) > 0) Then
            
                'RN_SGAD_21148YBJ
                'objocTratarCentroLocal.CentroFuncionalID = objFormSelecionarOTS.cboCentroFunc.ItemData(objFormSelecionarOTS.cboCentroFunc.ListIndex)
                objocTratarCentroLocal.CentroFuncionalID = lCentroFuncionalID
            
                objocTratarCentroLocal.DescricaoCF = sCentroFuncional

            End If
            
            objocTratarCentroLocal.PermiteExecucao = IIf(sPermiteExecucao = "", "N", _
                                                         sPermiteExecucao)
            
            objocTratarCentroLocal.CarregarForm

    ElseIf CodigoFase = GbCodigoFaseProgramarCPE Then

            Set objocCoordCPE = New clsInterOC_CoordCPE

            objocCoordCPE.InformaOCSelecionarOTS Me

            objocCoordCPE.Fase_ID = lFase_Id
            objocCoordCPE.Servico_ID = lServico_ID
            objocCoordCPE.Servico_ID_Anterior = lServico_ID_Anterior
            objocCoordCPE.Sigla_Acao = sSigla_Acao
            objocCoordCPE.Status_Fase = sStatus
            objocCoordCPE.Acao_ID = lAcao_ID
            objocCoordCPE.Produto_ID = Produto_ID
            objocCoordCPE.DataFinalizacao = sDataFimReal
            objocCoordCPE.Sequencia_Acao = lSequencia_Acao
            objocCoordCPE.ContratoID = lContratoID
            objocCoordCPE.NrLocalInsta = lNrLocalInsta
            
            objFormSelecionarOTS.spdCircuitos.Col = 43
            objocCoordCPE.Data_OTS = objFormSelecionarOTS.spdCircuitos.Text
            
            objocCoordCPE.PermiteExecucao = IIf(sPermiteExecucao = "", "N", sPermiteExecucao)
            
            'RN_SGAD_21148YBJ
            'objocCoordCPE.CentroFuncionalID = objFormSelecionarOTS.cboCentroFunc.ItemData(objFormSelecionarOTS.cboCentroFunc.ListIndex)
            objocCoordCPE.CentroFuncionalID = lCentroFuncionalID

            objocCoordCPE.CarregarForm

    ElseIf CodigoFase = GbCodigoFasePreAtivacao Then

            Set objocPreAtivacao = New clsInterOC_PreAtivacao

            objocPreAtivacao.InformaOCSelecionarOTS Me

            objocPreAtivacao.Fase_ID = lFase_Id
            objocPreAtivacao.Servico_ID = lServico_ID
            objocPreAtivacao.Servico_ID_Anterior = lServico_ID_Anterior
            objocPreAtivacao.Sigla_Acao = sSigla_Acao
            objocPreAtivacao.Status_Fase = sStatus
            objocPreAtivacao.Acao_ID = lAcao_ID
            objocPreAtivacao.Produto_ID = Produto_ID
            objocPreAtivacao.DataFinalizacao = sDataFimReal
            objocPreAtivacao.Sequencia_Acao = lSequencia_Acao
            objocPreAtivacao.ContratoID = lContratoID
            objocPreAtivacao.NrLocalInsta = lNrLocalInsta
            objocPreAtivacao.PermiteExecucao = IIf(sPermiteExecucao = "", "N", sPermiteExecucao)
            
            'RN_SGAD_21148YBJ
            'objocPreAtivacao.CentroFuncionalID = objFormSelecionarOTS.cboCentroFunc.ItemData(objFormSelecionarOTS.cboCentroFunc.ListIndex)
            objocPreAtivacao.CentroFuncionalID = lCentroFuncionalID

            objocPreAtivacao.CarregarForm

    'Ativacao
    ElseIf CodigoFase = GbCodigoFaseAtivacao Then

        Set objocAtivacao = New clsInterOC_AtivacaoFinalizacao

        objocAtivacao.InformaOCSelecionarOTS Me

        objocAtivacao.Fase_ID = lFase_Id
        objocAtivacao.Servico_ID = lServico_ID
        objocAtivacao.Servico_ID_Anterior = lServico_ID_Anterior
        objocAtivacao.Sigla_Acao = sSigla_Acao
        objocAtivacao.Status_Fase = sStatus
        objocAtivacao.Acao_ID = lAcao_ID
        objocAtivacao.Produto_ID = Produto_ID
        objocAtivacao.DataFinalizacao = sDataFimReal
        objocAtivacao.Sequencia_Acao = lSequencia_Acao
        objocAtivacao.ContratoID = lContratoID
        objocAtivacao.NrLocalInsta = lNrLocalInsta
        objocAtivacao.PermiteExecucao = IIf(sPermiteExecucao = "", "N", sPermiteExecucao)

        'RN_SGAD_21148YBJ
        'objocAtivacao.CentroFuncionalID = objFormSelecionarOTS.cboCentroFunc.ItemData(objFormSelecionarOTS.cboCentroFunc.ListIndex)
        objocAtivacao.CentroFuncionalID = lCentroFuncionalID

        objocAtivacao.CarregarForm
    'Tratar BGP
    ElseIf CodigoFase = GbCodigoFaseTratarBGP Then

        Set objocTratarBGP = New clsInterOC_TratarBGP

        objocTratarBGP.InformaOCSelecionarOTS Me

        objocTratarBGP.Fase_ID = lFase_Id
        objocTratarBGP.Servico_ID = lServico_ID
        objocTratarBGP.Servico_ID_Anterior = lServico_ID_Anterior
        objocTratarBGP.Sigla_Acao = sSigla_Acao
        objocTratarBGP.Status_Fase = sStatus
        objocTratarBGP.Acao_ID = lAcao_ID
        objocTratarBGP.Produto_ID = Produto_ID
        objocTratarBGP.DataFinalizacao = sDataFimReal
        objocTratarBGP.Sequencia_Acao = lSequencia_Acao
        objocTratarBGP.ContratoID = lContratoID
        objocTratarBGP.NrLocalInsta = lNrLocalInsta
        objocTratarBGP.PermiteExecucao = IIf(sPermiteExecucao = "", "N", sPermiteExecucao)
        objocTratarBGP.PeerBGP = sPeer_BGP

        objocTratarBGP.CarregarForm

    End If


    'Este objeto será eliminado no chamador
'    Set objMensagem = Nothing

    GoTo Finalizar
       
Finalizar:

    Set objdadosServico = Nothing
    'JAH
    Set objContrato = Nothing

    On Error GoTo 0
    
    Screen.MousePointer = vbDefault
    
    Exit Function
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    
    If (objMensagem.Descricao = "") Then
    
        objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: TelaFase_BlFlex " & _
                            "da classe clsGenOC_SelecionarOTS"
    
    Else
    
        objMensagem.ExibeMensagem
        
    End If
    
    GoTo Finalizar

End Function
Public Function TelaFase_NetPhone() As Boolean

On Error GoTo ErrorHandler:

    Dim bRetorno                As Boolean
    Dim sStatus                 As String
    Dim lIDServico              As Long
    Dim vDesigOTS()             As Variant
    Dim objdadosServico         As clsGenServico
    Dim sCNL                    As String
    Dim lNumeroOTS              As Long
    Dim iAnoOTS                 As Integer
    Dim sPermiteExecucao        As String
    Dim sCentroFuncional        As String
    Dim sCodigoProduto          As String
    Dim lCentroFuncionalID      As Long
    Dim sTabela_Contrato        As String
    Dim sTelefone               As String
    Screen.MousePointer = vbArrowHourglass
    
    'Intância a classe de mensagem
    Set objMensagem = New clsgenMensagem

    'Instância a classe de fase para buscar dados do cabeçalho
    Set objFase = New clsgenFase

    'Atribui os valores da lista da OTS selecionada nas propriedades da fase
    bRetorno = AtribuiPropriedadesFase

    'Atribui a linha ativada
    objFormSelecionarOTS.spdCircuitos.Row = objFormSelecionarOTS.spdCircuitos.ActiveRow

    'Atribui o valor do status da OTS

    objFormSelecionarOTS.spdCircuitos.Col = 2
    sStatus = objFormSelecionarOTS.spdCircuitos.Text

    'RN_SGAD_21148YBJ
    If (GbFlagRN_SGAD_21148YBJ) Then
    
        objFormSelecionarOTS.spdCircuitos.Row = objFormSelecionarOTS.spdCircuitos.ActiveRow
        objFormSelecionarOTS.spdCircuitos.Col = 52
        sCentroFuncional = objFormSelecionarOTS.spdCircuitos.Text
    
    Else
    
        sCentroFuncional = objFormSelecionarOTS.cboCentroFunc.Text
    
    End If


    objFormSelecionarOTS.spdCircuitos.Col = 7
    objFormSelecionarOTS.spdCircuitos.Row = objFormSelecionarOTS.spdCircuitos.ActiveRow

    sStatus = sStatus & " / " & objFormSelecionarOTS.spdCircuitos.Text

    'Monta dados do Cabecalho
    If objBlfCabecalho Is Nothing Then

        Set objBlfCabecalho = New clsBlf_Cabecalho

    End If

    objFormSelecionarOTS.spdCircuitos.Col = 28
    
    Select Case objFormSelecionarOTS.spdCircuitos.Text
           Case "ATV"
                objBlfCabecalho.Acao = "Ativação"
           Case "ALT"
                objBlfCabecalho.Acao = "Alteração"
           Case "ATG"
                objBlfCabecalho.Acao = "Correção de Ativação"
           Case "ALG"
                objBlfCabecalho.Acao = "Alteração de Ativação"
           Case "DES"
                objBlfCabecalho.Acao = "Desativação"
           Case "CAN"
                objBlfCabecalho.Acao = "Cancelamento"
    End Select

    Set objdadosServico = New clsGenServico

    objdadosServico.InformaBanco frmgenMDI_SGAPlus.objBanco
    objdadosServico.InformaMensagem objMensagem
    
    'Resgata os dados do Serviço
    bRetorno = objdadosServico.RecuperarDados(lServico_ID)
    
    If (objMensagem.TipoMsg <> "") Then

       GoTo ErrorHandler

    End If
    
    objBlfCabecalho.TipoCircuito = "VOZ LOCAL"
    
    objFormSelecionarOTS.spdCircuitos.Col = 9
    lIDServico = CLng("0" & objFormSelecionarOTS.spdCircuitos.Text)
    
    bRetorno = objdadosServico.DevolveDesignacaoOTS(lIDServico, lProduto_ID, True)

    If objMensagem.TipoMsg <> "" Then

       GoTo ErrorHandler

    End If
    
    objBlfCabecalho.OtsPrincipal = objdadosServico.OTS
    
    'OBTEM A OTS ANTERIOR, CASO EXISTA
    objFormSelecionarOTS.spdCircuitos.Col = 11
    
    lIDServico = CLng("0" & objFormSelecionarOTS.spdCircuitos.Text)
    
    bRetorno = objdadosServico.DevolveDesignacaoOTS(lIDServico, lProduto_ID, True)

    If objMensagem.TipoMsg <> "" Then

       GoTo ErrorHandler

    End If

    objBlfCabecalho.OtsAnterior = objdadosServico.OTS
    
    objFormSelecionarOTS.spdCircuitos.Col = 37
    objBlfCabecalho.ContaCorrente = objFormSelecionarOTS.spdCircuitos.Text
    
    objFormSelecionarOTS.spdCircuitos.Col = 38
    objBlfCabecalho.NomeCliente = objFormSelecionarOTS.spdCircuitos.Text
              
    objFormSelecionarOTS.spdCircuitos.Col = 16
    
    If objFormSelecionarOTS.spdCircuitos.Text <> "" Then
       objBlfCabecalho.DataFinalizacao = Format(objFormSelecionarOTS.spdCircuitos.Text, "dd/mm/yyyy")
    Else
        objBlfCabecalho.DataFinalizacao = ""
    End If

    objFormSelecionarOTS.spdCircuitos.Col = 53  'coluna retorna o ID do CF da Fase
    lCentroFuncionalID = objFormSelecionarOTS.spdCircuitos.Text

    objFormSelecionarOTS.spdCircuitos.Col = 54
    sTelefone = "(" & objFormSelecionarOTS.spdCircuitos.Text & ") "

    objFormSelecionarOTS.spdCircuitos.Col = 55
    sTelefone = sTelefone & objFormSelecionarOTS.spdCircuitos.Text
    
    'Numero item do servico gerado pelo ATEN
    objFormSelecionarOTS.spdCircuitos.Col = 56
    objBlfCabecalho.NumeroOE = objFormSelecionarOTS.spdCircuitos.Text
    
    'Ano item do servico gerado pelo ATEN
    objFormSelecionarOTS.spdCircuitos.Col = 57
    objBlfCabecalho.NumeroOE = objBlfCabecalho.NumeroOE & IIf(Len(Trim(objFormSelecionarOTS.spdCircuitos.Text)) = 0, "", "/" & objFormSelecionarOTS.spdCircuitos.Text)
    
    objBlfCabecalho.NulTelefone = sTelefone
    
    'Configuracao conftel
    If CodigoFase = GbCodigoFaseConfiguracaoConftel Then

        Set objocConfigConftel = New clsInterOC_tratarConfigConftel

        objocConfigConftel.InformaOCSelecionarOTS Me
 
        objocConfigConftel.Fase_ID = lFase_Id
        objocConfigConftel.Servico_ID = lServico_ID
        objocConfigConftel.Servico_ID_Anterior = lServico_ID_Anterior
        objocConfigConftel.Sigla_Acao = sSigla_Acao
        objocConfigConftel.Status_Fase = sStatus
        objocConfigConftel.Acao_ID = lAcao_ID
        objocConfigConftel.Produto_ID = Produto_ID
        objocConfigConftel.DataFinalizacao = sDataFimReal
        objocConfigConftel.Sequencia_Acao = lSequencia_Acao
        objocConfigConftel.CentroFuncionalID = lCentroFuncionalID
        
        objFormSelecionarOTS.spdCircuitos.Col = 43
'        objocConfigConftel.Data_OTS = objFormSelecionarOTS.spdCircuitos.Text


        objocConfigConftel.CarregarForm

    End If
       
Finalizar:

    Set objdadosServico = Nothing

    On Error GoTo 0
    
    Screen.MousePointer = vbDefault
    
    Exit Function
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    
    If (objMensagem.Descricao = "") Then
    
        objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: TelaFase_NetPhone " & _
                            "da classe clsGenOC_SelecionarOTS"
    
    Else
    
        objMensagem.ExibeMensagem
        
    End If
    
    GoTo Finalizar

End Function

Public Function TelaFase_VIPLINEIP() As Boolean
 
On Error GoTo ErrorHandler:

    Dim bRetorno                As Boolean
    Dim sStatus                 As String
    Dim lIDServico              As Long
    Dim vDesigOTS()             As Variant
    Dim objdadosServico         As clsGenServico
    Dim objAtenItemServico      As VIP_Aten_Item_Servico
    Dim sCNL                    As String
    Dim lNumeroOTS              As Long
    Dim iAnoOTS                 As Integer
    Dim sPermiteExecucao        As String
    Dim sCentroFuncional        As String
    Dim sCodigoProduto          As String
    Dim lCentroFuncionalID      As Long
    Dim sTabela_Contrato        As String
    Dim sTelefone               As String
    Dim objDadosTecVoz          As VIP_Aten_Tec_Voz
    Dim objDadosContrato        As VIP_Aten_Contrato
    Dim objDadosTecBanda        As VIP_Aten_Dados_Tec_Banda
    
    Screen.MousePointer = vbArrowHourglass
    
    'Intância a classe de mensagem
    Set objMensagem = New clsgenMensagem

    'Instância a classe de fase para buscar dados do cabeçalho
    Set objFase = New clsgenFase

    'Atribui os valores da lista da OTS selecionada nas propriedades da fase
    bRetorno = AtribuiPropriedadesFase

    'Atribui a linha ativada
    objFormSelecionarOTS.spdCircuitos.Row = objFormSelecionarOTS.spdCircuitos.ActiveRow

    'Atribui o valor do status da OTS

    objFormSelecionarOTS.spdCircuitos.Col = 2
    sStatus = objFormSelecionarOTS.spdCircuitos.Text

    objFormSelecionarOTS.spdCircuitos.Row = objFormSelecionarOTS.spdCircuitos.ActiveRow
    objFormSelecionarOTS.spdCircuitos.Col = 52
    sCentroFuncional = objFormSelecionarOTS.spdCircuitos.Text

    objFormSelecionarOTS.spdCircuitos.Col = 7
    objFormSelecionarOTS.spdCircuitos.Row = objFormSelecionarOTS.spdCircuitos.ActiveRow

    sStatus = sStatus & " / " & objFormSelecionarOTS.spdCircuitos.Text

    'Monta dados do Cabecalho
    If objBlfCabecalho Is Nothing Then

        Set objBlfCabecalho = New clsBlf_Cabecalho

    End If

    objFormSelecionarOTS.spdCircuitos.Col = 28
    
    Select Case objFormSelecionarOTS.spdCircuitos.Text
           Case "ATV"
                objBlfCabecalho.Acao = "Ativação"
           Case "ALT"
                objBlfCabecalho.Acao = "Alteração"
           Case "ATG"
                objBlfCabecalho.Acao = "Correção de Ativação"
           Case "ALG"
                objBlfCabecalho.Acao = "Alteração de Ativação"
           Case "DES"
                objBlfCabecalho.Acao = "Desativação"
           Case "CAN"
                objBlfCabecalho.Acao = "Cancelamento"
    End Select

    Set objdadosServico = New clsGenServico

    objdadosServico.InformaBanco frmgenMDI_SGAPlus.objBanco
    objdadosServico.InformaMensagem objMensagem
    
    'Resgata os dados do Serviço
    bRetorno = objdadosServico.RecuperarDados(lServico_ID)
    
    If (objMensagem.TipoMsg <> "") Then

       GoTo ErrorHandler

    End If
    
    Set objAtenItemServico = New VIP_Aten_Item_Servico

    objAtenItemServico.InformaBanco frmgenMDI_SGAPlus.objBanco
    objAtenItemServico.InformaMensagem objMensagem
    
    bRetorno = objAtenItemServico.RecuperarDados(lServico_ID)

    If (objMensagem.TipoMsg <> "") Then

       GoTo ErrorHandler

    End If
    
    If objAtenItemServico.Tipo_Servico = "V" Then
        
        objBlfCabecalho.TipoCircuito = "VOZ LOCAL"
        objBlfCabecalho.TipoItem_Codigo = "V"
    
    Else
        
        objBlfCabecalho.TipoCircuito = "BANDA LARGA"
        objBlfCabecalho.TipoItem_Codigo = "I"
    
    End If
    
    objFormSelecionarOTS.spdCircuitos.Col = 9
    lIDServico = CLng("0" & objFormSelecionarOTS.spdCircuitos.Text)
    
    bRetorno = objdadosServico.DevolveDesignacaoOTS(lIDServico, lProduto_ID, True)

    If objMensagem.TipoMsg <> "" Then

       GoTo ErrorHandler

    End If
    
    objBlfCabecalho.OtsPrincipal = objdadosServico.OTS
    objBlfCabecalho.DesignacaoPrincipal = objdadosServico.DesignacaoCircuito
    
    'OBTEM A OTS ANTERIOR, CASO EXISTA
    objFormSelecionarOTS.spdCircuitos.Col = 11
    
    lIDServico = CLng("0" & objFormSelecionarOTS.spdCircuitos.Text)
    
    bRetorno = objdadosServico.DevolveDesignacaoOTS(lIDServico, lProduto_ID, True)

    If objMensagem.TipoMsg <> "" Then

       GoTo ErrorHandler

    End If

    objBlfCabecalho.OtsAnterior = objdadosServico.OTS
    
    objFormSelecionarOTS.spdCircuitos.Col = 37
    objBlfCabecalho.ContaCorrente = objFormSelecionarOTS.spdCircuitos.Text
     
    objFormSelecionarOTS.spdCircuitos.Col = 38
    objBlfCabecalho.NomeCliente = objFormSelecionarOTS.spdCircuitos.Text
    
    Set objDadosContrato = New VIP_Aten_Contrato

    objDadosContrato.InformaBanco frmgenMDI_SGAPlus.objBanco
    objDadosContrato.InformaMensagem objMensagem
    
    bRetorno = objDadosContrato.RecuperarDados(objAtenItemServico.ContratoID)

    If (objMensagem.TipoMsg <> "") Then

       GoTo ErrorHandler

    End If
    
    objBlfCabecalho.NomeFantasia = objDadosContrato.Nm_Fantasia
    
    objBlfCabecalho.Contrato = objDadosContrato.NumeroContrato & "/" & _
                               objDadosContrato.AnoContrato & "-" & _
                               objDadosContrato.NumeroAditivo
    
    objFormSelecionarOTS.spdCircuitos.Col = 16
    
    If objFormSelecionarOTS.spdCircuitos.Text <> "" Then
       
       objBlfCabecalho.DataFinalizacao = Format(objFormSelecionarOTS.spdCircuitos.Text, "dd/mm/yyyy")
       
    Else
    
        objBlfCabecalho.DataFinalizacao = ""
        
    End If

    
    objFormSelecionarOTS.spdCircuitos.Col = 53  'coluna retorna o ID do CF da Fase
    lCentroFuncionalID = objFormSelecionarOTS.spdCircuitos.Text

    If objAtenItemServico.Tipo_Servico = "V" Then

        Set objDadosTecVoz = New VIP_Aten_Tec_Voz

        objDadosTecVoz.InformaBanco frmgenMDI_SGAPlus.objBanco
        objDadosTecVoz.InformaMensagem objMensagem
        
        bRetorno = objDadosTecVoz.RecuperarDados(objAtenItemServico.ID)
        
        If (objMensagem.TipoMsg <> "") Then
    
           GoTo ErrorHandler
    
        End If

        sTelefone = "(" & objDadosTecVoz.CodigoAreaTerminal & ") "
        sTelefone = sTelefone & objDadosTecVoz.NumeroTerminal
        
        objBlfCabecalho.NulTelefone = sTelefone
        
        If objDadosTecVoz.IdentificadorLinha = "S" Then
            
            objBlfCabecalho.TipoCircuito = "MULTILINHA"
        
        ElseIf objDadosTecVoz.IdentificadorLinha = "I" Then
            
            objBlfCabecalho.TipoCircuito = "INDIVIDUAL"
        
        ElseIf objDadosTecVoz.IdentificadorLinha = "C" Then
            
            objBlfCabecalho.TipoCircuito = "PABX VIRTUAL (CENTREX)"
        
        End If
        
        sTelefone = ""
        
        Set objDadosTecBanda = New VIP_Aten_Dados_Tec_Banda

        objDadosTecBanda.InformaBanco frmgenMDI_SGAPlus.objBanco
        objDadosTecBanda.InformaMensagem objMensagem
        
        bRetorno = objDadosTecBanda.RecuperarDados(objAtenItemServico.ID)
        
        If (objMensagem.TipoMsg <> "") Then
    
           GoTo ErrorHandler
    
        End If
        
        objBlfCabecalho.Velocidade = objDadosTecBanda.VelocidadeDownload
        
    End If
    
    objFormSelecionarOTS.spdCircuitos.Col = 33
    objBlfCabecalho.CNL = objFormSelecionarOTS.spdCircuitos.Text
    
    'Numero item do servico gerado pelo ATEN
    objBlfCabecalho.NumeroOE = objAtenItemServico.NumeroItemSite & "/" & objAtenItemServico.AnoItemAtenID
    objBlfCabecalho.ItemOE = objAtenItemServico.TotalItem
    
    If CodigoFase = GbCodigoFaseAlocacao Then

        Set objOcAlocarCF = New clsInterOC_AlocarCF

        objOcAlocarCF.InformaOC_SelecionarOTS Me
 
        objOcAlocarCF.Fase_ID = lFase_Id
        objOcAlocarCF.Servico_ID = lServico_ID
        objOcAlocarCF.Servico_ID_Anterior = lServico_ID_Anterior
        objOcAlocarCF.Sigla_Acao = sSigla_Acao
        objOcAlocarCF.Status_Fase = sStatus
        objOcAlocarCF.Acao_ID = lAcao_ID
        objOcAlocarCF.Produto_ID = Produto_ID
        objOcAlocarCF.DataFinalizacao = sDataFimReal
        objOcAlocarCF.Sequencia_Acao = lSequencia_Acao
        
        'RN_SGAD_21148YBJ
        'objOcAlocarCF.CentroFuncionalID = objFormSelecionarOTS.cboCentroFunc.ItemData(objFormSelecionarOTS.cboCentroFunc.ListIndex)
        objOcAlocarCF.CentroFuncionalID = lCentroFuncionalID
        
        objFormSelecionarOTS.spdCircuitos.Col = 43
        objOcAlocarCF.Data_OTS = objFormSelecionarOTS.spdCircuitos.Text

        objFormSelecionarOTS.spdCircuitos.Col = 45
        objOcAlocarCF.Realocacao = objFormSelecionarOTS.spdCircuitos.Text

        objOcAlocarCF.CarregarForm
    
    ElseIf CodigoFase = GbCodigoFaseTratarIP Then

        Set objOC_TratarIP = New clsVIPOC_TratarIPFixo

        objOC_TratarIP.InformaOCSelecionarOTS Me
 
        objOC_TratarIP.Fase_ID = lFase_Id
        objOC_TratarIP.Servico_ID = lServico_ID
        objOC_TratarIP.Servico_ID_Anterior = lServico_ID_Anterior
        objOC_TratarIP.Sigla_Acao = sSigla_Acao
        objOC_TratarIP.Status_Fase = sStatus
        objOC_TratarIP.Acao_ID = lAcao_ID
        objOC_TratarIP.Produto_ID = Produto_ID
        objOC_TratarIP.DataFinalizacao = sDataFimReal
        objOC_TratarIP.Sequencia_Acao = lSequencia_Acao
        objOC_TratarIP.CentroFuncionalID = lCentroFuncionalID
        
'        objFormSelecionarOTS.spdCircuitos.Col = 43
'        objOC_TratarIP.Data_OTS = objFormSelecionarOTS.spdCircuitos.Text

        objOC_TratarIP.CarregarForm
    
    ElseIf CodigoFase = GbCodigoFaseAcesso Then

        Set objOC_AcessoEmbratel = New clsGenOC_TratarAcessoEmbratel

        'Método que recebe todos os dados necessários e que
        'efetivamente inicia o trabalho do Controlador

        objOC_AcessoEmbratel.InformaOC_SelecionarOTS Me

        objOC_AcessoEmbratel.Fase_ID = lFase_Id
        objOC_AcessoEmbratel.Servico_ID = lServico_ID
        objOC_AcessoEmbratel.Servico_ID_Anterior = lServico_ID_Anterior
        objOC_AcessoEmbratel.Sigla_Acao = sSigla_Acao
        objOC_AcessoEmbratel.Status_Fase = sStatus
        objOC_AcessoEmbratel.Acao_ID = lAcao_ID
        objOC_AcessoEmbratel.Produto_ID = Produto_ID
        objOC_AcessoEmbratel.DataFinalizacao = sDataFimReal
        objOC_AcessoEmbratel.Sequencia_Acao = lSequencia_Acao
        objOC_AcessoEmbratel.ContratoID = objAtenItemServico.ContratoID
        objOC_AcessoEmbratel.EnderecoID = objAtenItemServico.EnderecoID
        objOC_AcessoEmbratel.ItemServicoID = objAtenItemServico.ID
        objOC_AcessoEmbratel.Familia_Produto = Familia_Produto
        objOC_AcessoEmbratel.CentroFuncionalID = lCentroFuncionalID
        objOC_AcessoEmbratel.PermiteExecucao = IIf(sPermiteExecucao = "", "N", _
                                                     sPermiteExecucao)
        objOC_AcessoEmbratel.CarregarForm
    
    ElseIf CodigoFase = GbCodigoFaseCentroLocal Then

            Set objVIPocTratarCL = New clsVIPOC_TratarCentroLocal

            'Método que recebe todos os dados necessários e que efetivamente inicia o trabalho do Controlador
            objVIPocTratarCL.InformaOCSelecionarOTS Me

            objVIPocTratarCL.Fase_ID = lFase_Id
            objVIPocTratarCL.Servico_ID = lServico_ID
            objVIPocTratarCL.Servico_ID_Anterior = lServico_ID_Anterior
            objVIPocTratarCL.Sigla_Acao = sSigla_Acao
            objVIPocTratarCL.Acao_ID = lAcao_ID
            objVIPocTratarCL.Status_Fase = sStatus
            objVIPocTratarCL.Produto_ID = Produto_ID
            objVIPocTratarCL.DataFinalizacao = sDataFimReal
            objVIPocTratarCL.Sequencia_Acao = lSequencia_Acao
            objVIPocTratarCL.ContratoID = objAtenItemServico.ContratoID
            objVIPocTratarCL.CentroFuncionalID = lCentroFuncionalID
            objVIPocTratarCL.PermiteExecucao = IIf(sPermiteExecucao = "", "N", _
                                                         sPermiteExecucao)
            
            objVIPocTratarCL.CarregarForm
    
    ElseIf CodigoFase = GbCodigoFasePreAtivacao Then

            Set objocPreAtivacao = New clsInterOC_PreAtivacao

            objocPreAtivacao.InformaOCSelecionarOTS Me

            objocPreAtivacao.Fase_ID = lFase_Id
            objocPreAtivacao.Servico_ID = lServico_ID
            objocPreAtivacao.Servico_ID_Anterior = lServico_ID_Anterior
            objocPreAtivacao.Sigla_Acao = sSigla_Acao
            objocPreAtivacao.Status_Fase = sStatus
            objocPreAtivacao.Acao_ID = lAcao_ID
            objocPreAtivacao.Produto_ID = Produto_ID
            objocPreAtivacao.DataFinalizacao = sDataFimReal
            objocPreAtivacao.Sequencia_Acao = lSequencia_Acao
            objocPreAtivacao.ContratoID = objAtenItemServico.ContratoID
            objocPreAtivacao.EnderecoID = objAtenItemServico.EnderecoID
            objocPreAtivacao.PermiteExecucao = IIf(sPermiteExecucao = "", "N", sPermiteExecucao)
            
            'RN_SGAD_21148YBJ
            'objocPreAtivacao.CentroFuncionalID = objFormSelecionarOTS.cboCentroFunc.ItemData(objFormSelecionarOTS.cboCentroFunc.ListIndex)
            objocPreAtivacao.CentroFuncionalID = lCentroFuncionalID

            objocPreAtivacao.CarregarForm
    'Configuracao NGN Acesso
    ElseIf CodigoFase = GbCodigoFaseConfiguracaoNGN Then

            Set objocNGNAcesso = New clsInterOC_ConfigAcessoNGN

            objocNGNAcesso.InformaOCSelecionarOTS Me

            objocNGNAcesso.Fase_ID = lFase_Id
            objocNGNAcesso.Servico_ID = lServico_ID
            objocNGNAcesso.Servico_ID_Anterior = lServico_ID_Anterior
            objocNGNAcesso.Sigla_Acao = sSigla_Acao
            objocNGNAcesso.Status_Fase = sStatus
            objocNGNAcesso.Acao_ID = lAcao_ID
            objocNGNAcesso.Produto_ID = Produto_ID
            objocNGNAcesso.DataFinalizacao = sDataFimReal
            objocNGNAcesso.Sequencia_Acao = lSequencia_Acao
            objocNGNAcesso.ContratoID = objAtenItemServico.ContratoID
            objocNGNAcesso.PermiteExecucao = IIf(sPermiteExecucao = "", "N", sPermiteExecucao)
            objocNGNAcesso.CentroFuncionalID = lCentroFuncionalID

            objocNGNAcesso.CarregarForm
    
    ElseIf CodigoFase = GbCodigoFaseConfiguracaoConftel Then

            Set objocNGNNucleo = New clsInterOC_tratarNGNNucleo

            objocNGNNucleo.InformaOCSelecionarOTS Me

            objocNGNNucleo.Fase_ID = lFase_Id
            objocNGNNucleo.Servico_ID = lServico_ID
            objocNGNNucleo.Servico_ID_Anterior = lServico_ID_Anterior
            objocNGNNucleo.Sigla_Acao = sSigla_Acao
            objocNGNNucleo.Status_Fase = sStatus
            objocNGNNucleo.Acao_ID = lAcao_ID
            objocNGNNucleo.Produto_ID = Produto_ID
            objocNGNNucleo.DataFinalizacao = sDataFimReal
            objocNGNNucleo.Sequencia_Acao = lSequencia_Acao
            objocNGNNucleo.ContratoID = Val(objAtenItemServico.ContratoID)
            objocNGNNucleo.EnderecoInstalacaoID = Val(objAtenItemServico.EnderecoID)
            objocNGNNucleo.CentrexID = Val(objAtenItemServico.CentrexID)
            objocNGNNucleo.ItemServicoID = Val(objAtenItemServico.ID)
            
            objocNGNNucleo.PermiteExecucao = IIf(sPermiteExecucao = "", "N", sPermiteExecucao)
            
            objocNGNNucleo.CentroFuncionalID = lCentroFuncionalID

            objocNGNNucleo.CarregarForm
    
    ElseIf CodigoFase = GbCodigoFaseAtivacao Then

        Set objocAtivacao = New clsVIPOC_AtivacaoFinalizacao

        objocAtivacao.InformaOCSelecionarOTS Me

        objocAtivacao.Fase_ID = lFase_Id
        objocAtivacao.Servico_ID = lServico_ID
        objocAtivacao.Servico_ID_Anterior = lServico_ID_Anterior
        objocAtivacao.Sigla_Acao = sSigla_Acao
        objocAtivacao.Status_Fase = sStatus
        objocAtivacao.Acao_ID = lAcao_ID
        objocAtivacao.Produto_ID = Produto_ID
        objocAtivacao.DataFinalizacao = sDataFimReal
        objocAtivacao.Sequencia_Acao = lSequencia_Acao
        objocAtivacao.ContratoID = lContratoID
        objocAtivacao.NrLocalInsta = lNrLocalInsta
        objocAtivacao.PermiteExecucao = IIf(sPermiteExecucao = "", "N", sPermiteExecucao)

        'RN_SGAD_21148YBJ
        'objocAtivacao.CentroFuncionalID = objFormSelecionarOTS.cboCentroFunc.ItemData(objFormSelecionarOTS.cboCentroFunc.ListIndex)
        objocAtivacao.CentroFuncionalID = lCentroFuncionalID

        objocAtivacao.CarregarForm

    End If
       
Finalizar:
    
    Set objdadosServico = Nothing

    On Error GoTo 0
    
    Screen.MousePointer = vbDefault
    
    Exit Function
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    
    If (objMensagem.Descricao = "") Then
    
        objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: TelaFase_VIPLINEIP " & _
                            "da classe clsGenOC_SelecionarOTS"
    
        
    End If
        
    objMensagem.ExibeMensagem
    
    GoTo Finalizar

End Function

Public Sub TelaAnotacoes(ByVal pbChamouDaSelecionarOTS As Boolean)

On Error GoTo ErrorHandler:

    Dim bRetorno                As Boolean
    Dim objFormAnotacoes        As frmgenAnotacoes
    Dim lProdutoLinha           As Long
    
    If objFormSelecionarOTS.spdCircuitos.ActiveRow <= 0 Then
        
        Exit Sub
        
    End If
        
    objFormSelecionarOTS.spdCircuitos.Row = objFormSelecionarOTS.spdCircuitos.ActiveRow
    objFormSelecionarOTS.spdCircuitos.Col = 26
    
    lProdutoLinha = CLng(objFormSelecionarOTS.spdCircuitos.Text)
    
    Screen.MousePointer = vbArrowHourglass
    
    ' Intância a classe de mensagem
    Set objMensagem = New clsgenMensagem

    ' Instância o formulário de que exibe os números de designação e OTS
    Set objFormAnotacoes = New frmgenAnotacoes
    
    ' Atribui o valor da propriedade no objeto da tela correspondente ao texto
    objFormAnotacoes.txtTexto = objFase.Anotacao
    
    ' Atribui o título da tela
    
    If lProdutoLinha = GbIDProdutoWID Then
    
        ' Dedicated
        objFormAnotacoes.Caption = "Worldcom Internet Dedicated"
        
    ElseIf lProdutoLinha = GbIDProdutoDiginet Then
    
        ' Diginet
        objFormAnotacoes.Caption = "Diginet"
    
    ElseIf lProdutoLinha = GbIDProdutoEmbratelIPL Then
    
        ' Embratel IPL
        objFormAnotacoes.Caption = "Embratel IPL"
    
    ElseIf lProdutoLinha = GbIDProdutoBLF Then
    
        ' BLF
        objFormAnotacoes.Caption = "Business Link Flex"

    End If
     
    ' Se a chamada for da tela se seleção de OTS, colocar o botão de gravação visível
    If pbChamouDaSelecionarOTS Then
    
        objFormAnotacoes.cmdGravar.Enabled = True
        objFormAnotacoes.cmdGravar.Visible = True
        
        objFormSelecionarOTS.spdCircuitos.Row = objFormSelecionarOTS.spdCircuitos.ActiveRow
        objFormSelecionarOTS.spdCircuitos.Col = 3
        objFormAnotacoes.lblCabecalho.Caption = "Anotação OTS: " & Mid(objFormSelecionarOTS.spdCircuitos.Text, 2)
        
    Else
        
        objFormAnotacoes.cmdGravar.Enabled = False
        objFormAnotacoes.cmdGravar.Visible = False
        
        objFormAnotacoes.lblCabecalho.Caption = "Anotação OTS: " & objFase.OTS
    
    End If
    
    ' Informar o objeto controlador
    objFormAnotacoes.informa_OC Me
    
    ' Informa o banco para a classe de fase
    objFase.InformaBanco frmgenMDI_SGAPlus.objBanco
    
    ' Informa a mensagem para a classe de fase
    objFase.InformaMensagem objMensagem
    
    ' Executa o método para recuperar os de anotação da fase
    'Coluna 8 - ID FASE
    objFormSelecionarOTS.spdCircuitos.Row = objFormSelecionarOTS.spdCircuitos.ActiveRow
    objFormSelecionarOTS.spdCircuitos.Col = 8
    
    bRetorno = objFase.RecuperarAnotacao(CLng(objFormSelecionarOTS.spdCircuitos.Text))
    
    ' Se ocorreu algum erro na DLL
    If Not bRetorno Then
        
        GoTo Finalizar
    
    End If
    
    objFormAnotacoes.txtTexto.Text = objFase.Anotacao
   
    objFormAnotacoes.Show vbModal
     
    Screen.MousePointer = vbDefault
    
    objFase.Anotacao = objFormAnotacoes.sTexto
        
Finalizar:

    Set objFormAnotacoes = Nothing
    
    Exit Sub
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: TelaAnotacoes " & _
                            "da classe clsgenOC_SelecionarOTS"
    
    GoTo Finalizar

End Sub

Public Sub ObtemAnotacoesGen()

    'Desevolvido para melhor controle de anotações já digitadas e cliques em outras linhas
    'do grid. Necessidade identificada durante UAT do IP VPN

On Error GoTo ErrorHandler:

    Dim bRetorno                As Boolean
    Dim objDetalheAnotacao      As New clsGenDetalheAnotacao
    Dim lRow                    As Long
    Dim lFase                   As Long
    Dim lSequenciaAcao          As Long
    Dim lProdutoLinha           As Long
    
    If objFormSelecionarOTS.spdCircuitos.ActiveRow <= 0 Then
        
        Exit Sub
        
    End If
        
    objFormSelecionarOTS.spdCircuitos.Row = objFormSelecionarOTS.spdCircuitos.ActiveRow
    objFormSelecionarOTS.spdCircuitos.Col = 26
    
    lProdutoLinha = CLng(objFormSelecionarOTS.spdCircuitos.Text)
    
    Screen.MousePointer = vbArrowHourglass
    
    Set objMensagem = New clsgenMensagem
    
    objDetalheAnotacao.InformaBanco frmgenMDI_SGAPlus.objBanco
    objDetalheAnotacao.InformaMensagem objMensagem
    
    'Coluna 8 - Armazena o ID da FASE
    objFormSelecionarOTS.spdCircuitos.Row = objFormSelecionarOTS.spdCircuitos.ActiveRow
    objFormSelecionarOTS.spdCircuitos.Col = 8
    
    lFase = CLng(objFormSelecionarOTS.spdCircuitos.Text)
        
    If objDetalheAnotacao.RecuperarDados(lFase, "A") Then
        
        Anotacao = objDetalheAnotacao.Anotacao
    
    End If
    
    Screen.MousePointer = vbDefault
    
Finalizar:

    Exit Sub
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: ObtemAnotacoesGen " & _
                            "da classe clsgenOC_SelecionarOTS"
    
    GoTo Finalizar

End Sub

Public Sub TelaAnotacoesGen(ByVal pbChamouDaSelecionarOTS As Boolean)

On Error GoTo ErrorHandler:

    Dim bRetorno                As Boolean
    Dim objAnotacao             As New clsGenAnotacao
    Dim objDetalheAnotacao      As New clsGenDetalheAnotacao
    Dim lRow                    As Long
    Dim lFase                   As Long
    Dim objFormAnotacoes        As frmgenAnotacoesAcumuladas
    Dim lSequenciaAcao          As Long
    Dim lProdutoLinha           As Long
    
    If objFormSelecionarOTS.spdCircuitos.ActiveRow <= 0 Then
        
        Exit Sub
        
    End If
        
    objFormSelecionarOTS.spdCircuitos.Row = objFormSelecionarOTS.spdCircuitos.ActiveRow
    objFormSelecionarOTS.spdCircuitos.Col = 26
    
    lProdutoLinha = CLng(objFormSelecionarOTS.spdCircuitos.Text)
    
    Screen.MousePointer = vbArrowHourglass
    
    Set objMensagem = New clsgenMensagem
    
    Set objFormAnotacoes = New frmgenAnotacoesAcumuladas
    
    objFormAnotacoes.Top = 1290
    
    'Limpa Anotação anterior
    objFormAnotacoes.txtTexto.Text = ""
    
    If lProdutoLinha = GbIDProdutoWID Then
    
        ' Dedicated
        objFormAnotacoes.Caption = "Worldcom Internet Dedicated"
        
    ElseIf lProdutoLinha = GbIDProdutoDiginet Then
    
        ' Diginet
        objFormAnotacoes.Caption = "Diginet"
    
    ElseIf lProdutoLinha = GbIDProdutoEmbratelIPL Then
    
        ' Embratel IPL
        objFormAnotacoes.Caption = "Embratel  I P L"
    
    ElseIf lProdutoLinha = GbIDProdutoBLF Then
    
        ' BLF
        objFormAnotacoes.Caption = "Business Link Flex"

    ElseIf (lProdutoLinha = GbIDProdutoBIV) Then
    
        ' IVN
        objFormAnotacoes.Caption = "IP VPN"

    ElseIf (lProdutoLinha = GbIDProdutoJAH) Then
    
        ' JAH
        objFormAnotacoes.Caption = "JAH!"

    End If
    
    objDetalheAnotacao.InformaBanco frmgenMDI_SGAPlus.objBanco
    objDetalheAnotacao.InformaMensagem objMensagem
    
    'Coluna 8 - Armazena o ID da FASE
    objFormSelecionarOTS.spdCircuitos.Row = objFormSelecionarOTS.spdCircuitos.ActiveRow
    objFormSelecionarOTS.spdCircuitos.Col = 8
    
    lFase = CLng(objFormSelecionarOTS.spdCircuitos.Text)
        
    If objDetalheAnotacao.RecuperarDados(lFase, "A") Then
        
        objFormAnotacoes.txtTexto.Text = objDetalheAnotacao.Anotacao
    
    End If
    
    If Not pbChamouDaSelecionarOTS Then
       
       objFormAnotacoes.txtTexto.Text = Anotacao
    
    End If
    
    ' Se a chamada for da tela se seleção de OTS, colocar o botão de gravação visível
    objFormSelecionarOTS.spdCircuitos.Row = objFormSelecionarOTS.spdCircuitos.ActiveRow
    objFormSelecionarOTS.spdCircuitos.Col = 3
    
    If pbChamouDaSelecionarOTS Then
    
        objFormAnotacoes.cmdGravar.Enabled = True
        objFormAnotacoes.cmdGravar.Visible = True
        
        objFormAnotacoes.lblCabecalho.Caption = "Anotação OTS: " & objFormSelecionarOTS.spdCircuitos.Text
        
    Else
        
        objFormAnotacoes.cmdGravar.Enabled = False
        objFormAnotacoes.cmdGravar.Visible = False
        
        objFormAnotacoes.lblCabecalho.Caption = "Anotação OTS: " & objFormSelecionarOTS.spdCircuitos.Text
    
    End If
    
    ' Informar o objeto controlador
    objFormAnotacoes.informa_OC Me
    
    objAnotacao.InformaBanco frmgenMDI_SGAPlus.objBanco
    objAnotacao.InformaMensagem objMensagem
    
    ' Executa o método para recuperar os de anotação da fase
    
    'Coluna 23 - Armazena a Sequencia_Acao
        objFormSelecionarOTS.spdCircuitos.Row = objFormSelecionarOTS.spdCircuitos.ActiveRow
    objFormSelecionarOTS.spdCircuitos.Col = 23
    
    lSequenciaAcao = CLng(objFormSelecionarOTS.spdCircuitos.Text)
    
    'Coluna 9 - Armazena o ID do Servico
    objFormSelecionarOTS.spdCircuitos.Row = objFormSelecionarOTS.spdCircuitos.ActiveRow
    objFormSelecionarOTS.spdCircuitos.Col = 9
    
    bRetorno = objAnotacao.RecuperarAnotacaoServico(CLng(objFormSelecionarOTS.spdCircuitos.Text), CStr(lSequenciaAcao))
    
    ' Se ocorreu algum erro na DLL
    
    objFormAnotacoes.SpdListaAnotacao.MaxRows = 0
    
    If bRetorno Then
    
        lRow = 1
        
        For Each objDetalheAnotacao In objAnotacao.ObjColecaoAnotacao
        
            objFormAnotacoes.SpdListaAnotacao.MaxRows = lRow
            objFormAnotacoes.SpdListaAnotacao.Row = lRow
            
            objFormAnotacoes.SpdListaAnotacao.Col = 1
            objFormAnotacoes.SpdListaAnotacao.Text = objDetalheAnotacao.NomeFase
        
            objFormAnotacoes.SpdListaAnotacao.Col = 2
            objFormAnotacoes.SpdListaAnotacao.Text = objDetalheAnotacao.Anotacao
            
            objFormAnotacoes.ListaAnotacao.AddItem objDetalheAnotacao.NomeFase
            
            lRow = lRow + 1
            
        Next
    
        objFormAnotacoes.ListaAnotacao.ListIndex = 0
        
    End If
    
    objFormAnotacoes.Show vbModal
     
    Screen.MousePointer = vbDefault
    
Finalizar:

    Set objAnotacao = Nothing
    
    Exit Sub
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: TelaAnotacoes " & _
                            "da classe clsgenOC_SelecionarOTS"
    
    GoTo Finalizar

End Sub

Public Sub TelaAnotacoesMonitoracao(plProdutoID As Long, _
                                    plFaseID As Long, _
                                    psIdentificacaoOTS As String, _
                                    plSequenciaAcao As Long, _
                                    plServicoId As Long, _
                              ByVal pbHabilitaGravacao As Boolean)

On Error GoTo ErrorHandler:

    Dim bRetorno                As Boolean
    Dim objAnotacao             As New clsGenAnotacao
    Dim objDetalheAnotacao      As New clsGenDetalheAnotacao
    Dim lRow                    As Long
    Dim objFormAnotacoes        As frmgenAnotacoesAcumuladas
    Dim lSequenciaAcao          As Long
    
    Screen.MousePointer = vbArrowHourglass
    
    Set objMensagem = New clsgenMensagem
    
    Set objFormAnotacoes = New frmgenAnotacoesAcumuladas
    
    objFormAnotacoes.Top = 1290
    
    If plProdutoID = GbIDProdutoWID Then
    
        ' Dedicated
        objFormAnotacoes.Caption = "Worldcom Internet Dedicated"
        
    ElseIf plProdutoID = GbIDProdutoDiginet Then
    
        ' Diginet
        objFormAnotacoes.Caption = "Diginet"
    
    ElseIf plProdutoID = GbIDProdutoEmbratelIPL Then
    
        ' Embratel IPL
        objFormAnotacoes.Caption = "Embratel  I P L"
    
    ElseIf plProdutoID = GbIDProdutoBLF Then
    
        ' BLF
        objFormAnotacoes.Caption = "Business Link Flex"

    End If
    
    objDetalheAnotacao.InformaBanco frmgenMDI_SGAPlus.objBanco
    objDetalheAnotacao.InformaMensagem objMensagem
        
    If objDetalheAnotacao.RecuperarDados(plFaseID, "A") Then
        
        If Trim(Anotacao) = "" Then
        
            objFormAnotacoes.txtTexto.Text = objDetalheAnotacao.Anotacao
            Anotacao = objDetalheAnotacao.Anotacao
            
        End If
        
    End If
    
    If Not pbHabilitaGravacao Then
       
       objFormAnotacoes.txtTexto.Text = Anotacao
    
    End If
    
    If pbHabilitaGravacao Then
    
        objFormAnotacoes.cmdGravar.Enabled = True
        objFormAnotacoes.cmdGravar.Visible = True
        
        objFormAnotacoes.lblCabecalho.Caption = "Anotação OTS: " & psIdentificacaoOTS
        
    Else
        
        objFormAnotacoes.cmdGravar.Enabled = False
        objFormAnotacoes.cmdGravar.Visible = False
        
        objFormAnotacoes.lblCabecalho.Caption = "Anotação OTS: " & psIdentificacaoOTS
    
    End If
    
    ' Informar o objeto controlador
    objFormAnotacoes.informa_OC Me
    
    objAnotacao.InformaBanco frmgenMDI_SGAPlus.objBanco
    objAnotacao.InformaMensagem objMensagem
    
    ' Executa o método para recuperar os de anotação da fase
    
    bRetorno = objAnotacao.RecuperarAnotacaoServico(plServicoId, CStr(plSequenciaAcao))
    
    objFormAnotacoes.SpdListaAnotacao.MaxRows = 0
    
    If bRetorno Then
    
        lRow = 1
        
        For Each objDetalheAnotacao In objAnotacao.ObjColecaoAnotacao
        
            objFormAnotacoes.SpdListaAnotacao.MaxRows = lRow
            objFormAnotacoes.SpdListaAnotacao.Row = lRow
            
            objFormAnotacoes.SpdListaAnotacao.Col = 1
            objFormAnotacoes.SpdListaAnotacao.Text = objDetalheAnotacao.NomeFase
        
            objFormAnotacoes.SpdListaAnotacao.Col = 2
            objFormAnotacoes.SpdListaAnotacao.Text = objDetalheAnotacao.Anotacao
            
            objFormAnotacoes.ListaAnotacao.AddItem objDetalheAnotacao.NomeFase
            
            lRow = lRow + 1
            
        Next
    
        objFormAnotacoes.ListaAnotacao.ListIndex = 0
        
    End If
    
    objFormAnotacoes.Show vbModal
     
    Screen.MousePointer = vbDefault
    
Finalizar:

    Set objAnotacao = Nothing
    
    Exit Sub
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: TelaAnotacoesMonitoracao " & _
                            "da classe clsgenOC_SelecionarOTS"
    
    GoTo Finalizar

End Sub

Public Function ExibirListaCentroFuncional(ByRef cbo As ComboBox, _
                                           Optional ByVal piCodigoFase As Integer, _
                                           Optional pbRetiraFiltroUsuario As Boolean, _
                                           Optional plIDCfExclusao As Long) As Boolean

On Error GoTo ErrorHandler:

    Dim bRetorno As Boolean
    Dim aListaCentroFuncional() As Variant
    Dim objMensagem As clsgenMensagem
    
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
    ReDim aListaCentroFuncional(0, 0)
    
    If pbRetiraFiltroUsuario Then
        
        If plIDCfExclusao <> 0 Then
            
           bRetorno = objCentroFuncional.MontarListaCentroFuncional(aListaCentroFuncional, _
                                                         piCodigoFase, 0, plIDCfExclusao)
        
        Else
            
            bRetorno = objCentroFuncional.MontarListaCentroFuncional(aListaCentroFuncional, _
                                                             piCodigoFase)
        
        End If
    
    Else
    
'        bRetorno = ObjCentroFuncional.MontarListaCentroFuncional(aListaCentroFuncional, _
                                                             piCodigoFase, _
                                                             GbID_Usuario)
    
        bRetorno = objCentroFuncional.MontarListaCFsemUsuario(aListaCentroFuncional, _
                                                             piCodigoFase)
    
    End If
    ' Se ocorreu algum erro na DLL exibi-lo
    If objMensagem.TipoMsg <> "" Then
        
        GoTo Finalizar
    
    End If
    
    If UBound(aListaCentroFuncional, 2) = 0 Then
    
        objMensagem.TipoMsg = "A"
        If pbRetiraFiltroUsuario Then
        
            objMensagem.Descricao = "Não existem Centros Funcionais cadastrados para o Usuário logado no sistema!"
            
        Else
        
            objMensagem.Descricao = "O Centro Funcional que Usuário pertence não tem permissão de acesso, para esta fase!"
            
        End If
        
        GoTo Finalizar

    End If
    
    PreencheCombo cbo, aListaCentroFuncional

    ExibirListaCentroFuncional = True
    
Finalizar:

    Screen.MousePointer = vbDefault
    
    If objMensagem.TipoMsg <> "" Then
    
        objMensagem.ExibeMensagem
        
    End If
    
    Set objMensagem = Nothing
    Set objCentroFuncional = Nothing
    
    Exit Function
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Houve um erro de execução no Método:" & _
                            "ExibirListaCentroFuncional da classe clsgenOC_SelecionarOTS"
    
    GoTo Finalizar

End Function

Public Function ExibirListaCliente() As Boolean

On Error GoTo ErrorHandler:

    Dim bRetorno As Boolean
    Dim iContador As Integer
    
    Screen.MousePointer = vbArrowHourglass
    
    ExibirListaCliente = False
    
    ' Intância a classe de mensagem
    Set objMensagem = New clsgenMensagem
    
    ' Limpa a combo
    objFormSelecionarOTS.cboCliente.Clear
    
    ' Instância a classe de cliente
    Set objCliente = New clsgenCliente

    ' Informar o nome do banco desejado
    objCliente.InformaBanco frmgenMDI_SGAPlus.objBanco
                
    ' Informar o nome do objeto de mensagem
    
    ' Executar o método de montagem da lista de cliente na classe de cliente
    bRetorno = objCliente.MontarListaCliente("C", objFormSelecionarOTS.txtCliente, aListaCliente)
    
    ' Se ocorreu algum erro na DLL exibi-lo
    If Not bRetorno Then
        
        GoTo Finalizar
    
    End If
    
    ' Preenche a combo de cliente
    For iContador = 1 To UBound(aListaCliente, 2)
        
        If Not IsEmpty(aListaCliente(0, iContador)) Then
            
            objFormSelecionarOTS.cboCliente.AddItem aListaCliente(1, iContador)
        
        Else
            
            Exit For
        
        End If
    
    Next
    
    ExibirListaCliente = True
    
    GoTo Finalizar
       
Finalizar:

    Screen.MousePointer = vbDefault
    
    If objMensagem.TipoMsg <> "" Then
    
        objMensagem.ExibeMensagem
        
    End If
    
    Set objCliente = Nothing
    Set objMensagem = Nothing
    
    Exit Function
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Houve um erro de execução no Método:" & _
                            "ExibirListaCliente da classe clsgenOC_SelecionarOTS"
    
    GoTo Finalizar

End Function

Public Function ExibeOE() As Boolean
    'Aquino
    'Exibe a tela Analisar OE para qualquer produto de Internet
    'Complementar futuramente para suporte à OE dos produtos de Dados
On Error GoTo ErrorHandler:

    ExibeOE = False

    Dim bRetorno            As Boolean
    Dim lIDServico          As Long
    
    ' Intância a classe de mensagem
    Set objMensagem = New clsgenMensagem
    
    objFormSelecionarOTS.spdCircuitos.Row = objFormSelecionarOTS.spdCircuitos.ActiveRow
    objFormSelecionarOTS.spdCircuitos.Col = 9
    lIDServico = CLng(0 & objFormSelecionarOTS.spdCircuitos.Text)
    
'''    'Valida o produto
'''    bRetorno = objServico.RecuperarDados(lIDServico)
'''
'''    If (Not bRetorno) Then
'''
'''        GoTo Finalizar
'''
'''    End If
'''
''''    If (objServico.ID_Produto = GbIDProdutoBLD) Then
''''
''''        objMensagem.TipoMsg = "A"
''''        objMensagem.Descricao = "Funcionalidade não disponível para este produto"
''''        GoTo Finalizar
''''
''''    End If
    
    bRetorno = CarregaFormConsultaAnalisarOEInternet(lIDServico)
    
    If (Not bRetorno) Then
        
        GoTo Finalizar
    
    End If
    
    ExibeOE = True
    
    GoTo Finalizar
       
Finalizar:

    Screen.MousePointer = vbDefault
    
    If objMensagem.TipoMsg <> "" Then
    
        objMensagem.ExibeMensagem
        
    End If
    
    Set objMensagem = Nothing
'''    Set objServico = Nothing

    Exit Function
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: ExibirListaOTS " & _
                            "da classe clsgenOC_SelecionarOTS"
    
    GoTo Finalizar
    
End Function
'
Public Function ExibirListaOTS() As Boolean

On Error GoTo ErrorHandler:

    Dim bRetorno            As Boolean
    Dim sDataInicial        As String
    Dim sDataFinal          As String
    Dim iCentroFuncional    As Long
    Dim sAcao               As String
    Dim sStatusEncaminhado  As String
    Dim sStatusARealizar    As String
    Dim sStatusRealizado    As String
    Dim sStatusOTS          As String
    Dim sStatusFaseRealocar As String
    Dim sContaCliente       As String
    Dim sRegioes            As String
    Dim sTipoAnotacao       As String
    Dim objDadosAcao        As clsGenAcao
    
    ExibirListaOTS = False
    
    If objFormSelecionarOTS Is Nothing Or GbControlaSaida Then
    
        Exit Function
    
    End If
    
    ' Intância a classe de mensagem
    Set objMensagem = New clsgenMensagem
    
    ' Instância a classe de serviço
    Set objServico = New clsGenServico

    ' Critica os campos obrigatórios
    If objFormSelecionarOTS.cboCentroFunc.Text = "" Then
    
        objMensagem.TipoMsg = "A"
        objMensagem.Descricao = "Centro Funcionado não Selecionado"
        objFormSelecionarOTS.cboCentroFunc.SetFocus
        
        GoTo Finalizar
        
    End If
    
    ' Atribui as variáveis para serem passadas como parâmetros
    sDataInicial = objFormSelecionarOTS.dtpDataIni.Value
    sDataFinal = objFormSelecionarOTS.dtpDataFim.Value
    
    ' Define melhor a AÇÃO para ser passada como parâmetro
    sAcao = Trim(objFormSelecionarOTS.cboAcao.Text)
    
    Set objDadosAcao = New clsGenAcao
    
    objDadosAcao.InformaBanco frmgenMDI_SGAPlus.objBanco
    objDadosAcao.InformaMensagem objMensagem
    
    If objFormSelecionarOTS.cboAcao.ListIndex >= 0 Then
    
       bRetorno = objDadosAcao.RecuperarDados( _
                                         CLng(objFormSelecionarOTS.cboAcao.ItemData( _
                                              objFormSelecionarOTS.cboAcao.ListIndex)))
       If (bRetorno) Then

           sAcao = objDadosAcao.Sigla

        Else

            objMensagem.TipoMsg = ""
            objMensagem.Descricao = ""
           
       End If
    
    End If
    
    
'    If sAcao = "Ativação" Then
'
'        sAcao = "ATV"
'
'    ElseIf sAcao = "Alteração" Then
'
'        sAcao = "ALT"
'
'    ElseIf sAcao = "Cancelamento" Then
'
'        sAcao = "CAN"
'
'    ElseIf sAcao = "Desativação" Then
'
'        sAcao = "DES"
'
'    ElseIf sAcao = "Correção (Ativação)" Then
'
'        sAcao = "ATG"
'
'    ElseIf sAcao = "Correção (Alteração)" Then
'
'        sAcao = "ALG"
'
'    End If
    
    sStatusARealizar = "0"
    sStatusRealizado = "0"
    
    ' Define o status da fase
    
    If objFormSelecionarOTS.ChkFaseEncaminhada.Value = 1 Then
    
        sStatusEncaminhado = "1"
        
    End If
    
    If objFormSelecionarOTS.chkFaseEmAndamento.Value = 1 Then
    
        sStatusARealizar = "1"
        
    End If
        
    If objFormSelecionarOTS.chkFaseRealizada.Value = 1 Then
    
        sStatusRealizado = "1"
        
    End If
    
    If objFormSelecionarOTS.ChkFaseRealocar.Value = 1 Then
    
        sStatusFaseRealocar = "1"
        
        sTipoAnotacao = "R"
        
    Else
        
        sStatusFaseRealocar = "0"
        
        sTipoAnotacao = "N"
    
    End If

    ' Define o status da OTS
    If objFormSelecionarOTS.chkOTSRealizada.Value = 1 Then
    
        sStatusOTS = "1"
        
    Else
        
        sStatusOTS = "0"
        
    End If
        
    If objFormSelecionarOTS.cboCentroFunc.ListIndex > -1 Then
    
        iCentroFuncional = objFormSelecionarOTS.cboCentroFunc.ItemData(objFormSelecionarOTS.cboCentroFunc.ListIndex)
    
    Else
    
        iCentroFuncional = 0
        
    End If
        
    If objFormSelecionarOTS.cboCliente.ListIndex > -1 Then
    
        sContaCliente = aListaCliente(0, objFormSelecionarOTS.cboCliente.ListIndex + 1)
    
    Else
    
        sContaCliente = ""
        
    End If
    
    If objFormSelecionarOTS.CboServico.ListIndex > 0 Then
    
        Produto_ID = objFormSelecionarOTS.CboServico.ItemData( _
                     objFormSelecionarOTS.CboServico.ListIndex)
    
    Else
    
        Produto_ID = 0
        
    End If
    
    objFormSelecionarOTS.spdCircuitos.MaxRows = 0
    
    objFormSelecionarOTS.cmdAnotacao.Enabled = False
    objFormSelecionarOTS.cmdImpOTS.Enabled = False
    objFormSelecionarOTS.CmdRealocar.Enabled = False
    objFormSelecionarOTS.CmdExcel.Enabled = False
    
    
    If Familia_Produto = "I" Then
       
       objFormSelecionarOTS.CmdOE.Enabled = False
    
    End If
    
    ' Informar o nome do banco desejado
    objServico.InformaBanco frmgenMDI_SGAPlus.objBanco
            
    ' Informar o nome do objeto de mensagem
    objServico.InformaMensagem objMensagem
    
    ' Executar o método de montagem da lista na classe de serviço
    
    'Caso seja a fase de designacao do produto diginet, a lista virá com a Operadora
    'no lugar da coluna Designacao
    
    If (Familia_Produto = "D") And CodigoFase = GbCodigoFaseDesignacao Then

        sRegioes = ""
        
        If objFormSelecionarOTS.OptRegiao(0).Value = 1 Then
           
           sRegioes = GbIDRegiaoAmericaNorte
        
        End If

        If objFormSelecionarOTS.OptRegiao(1).Value = 1 Then
        
           If Len(sRegioes) > 0 Then
              
              sRegioes = sRegioes & ", " & GbIDRegiaoAmericaSul
           
           Else
              
              sRegioes = sRegioes & GbIDRegiaoAmericaSul
           
           End If
        
        End If

        If objFormSelecionarOTS.OptRegiao(2).Value = 1 Then
           
           If Len(sRegioes) > 0 Then
              
              sRegioes = sRegioes & ", " & GbIDRegiaoEuropa
           
           Else
              
              sRegioes = sRegioes & GbIDRegiaoEuropa
           
           End If
        
        End If

        If objFormSelecionarOTS.OptRegiao(3).Value = 1 Then
           
           If Len(sRegioes) > 0 Then
              
              sRegioes = sRegioes & ", " & GbIDRegiaoOutros
           
           Else
              
              sRegioes = sRegioes & GbIDRegiaoOutros
           
           End If
        
        End If

        bRetorno = objServico.MontarListaOTSporRegiao(iSequenciaFase, _
                                             sDataInicial, _
                                             sDataFinal, _
                                             iCentroFuncional, _
                                             sAcao, _
                                             sStatusARealizar, _
                                             sStatusRealizado, _
                                             sStatusOTS, _
                                             sContaCliente, _
                                             Produto_ID, _
                                             sRegioes, _
                                             sStatusFaseRealocar, _
                                             sTipoAnotacao, Familia_Produto)
                                             
'    ElseIf CodigoFase = GbCodigoFaseConfiguracaoConftel Then
        
'        bRetorno = objServico.MontarListaOTS_Netphone(iSequenciaFase, _
                                             sDataInicial, _
                                             sDataFinal, _
                                             iCentroFuncional, _
                                             sAcao, _
                                             sStatusARealizar, _
                                             sStatusRealizado, _
                                             sStatusOTS, _
                                             sContaCliente, _
                                             Produto_ID, _
                                             Familia_Produto, _
                                             sStatusEncaminhado)
    
'    ElseIf (CodigoFase = GbCodigoFaseConfiguracaoNGN) Or (CodigoFase = GbCodigoFaseConfiguracaoConftel) Then
        
'        bRetorno = objServico.MontarListaOTS_Cli_ATEN(iSequenciaFase, _
                                             sDataInicial, _
                                             sDataFinal, _
                                             iCentroFuncional, _
                                             sAcao, _
                                             sStatusARealizar, _
                                             sStatusRealizado, _
                                             sStatusOTS, _
                                             sContaCliente, _
                                             Produto_ID, _
                                             Familia_Produto, _
                                             sStatusEncaminhado)
    Else
        
        bRetorno = objServico.MontarListaOTSTodos(iSequenciaFase, _
                                             sDataInicial, _
                                             sDataFinal, _
                                             iCentroFuncional, _
                                             sAcao, _
                                             sStatusARealizar, _
                                             sStatusRealizado, _
                                             sStatusOTS, _
                                             sContaCliente, _
                                             Produto_ID, _
                                             sStatusFaseRealocar, _
                                             sTipoAnotacao, Familia_Produto, _
                                             sStatusEncaminhado)

'        bRetorno = objServico.MontarListaOTS(iSequenciaFase, _
                                             sDataInicial, _
                                             sDataFinal, _
                                             iCentroFuncional, _
                                             sAcao, _
                                             sStatusARealizar, _
                                             sStatusRealizado, _
                                             sStatusOTS, _
                                             sContaCliente, _
                                             Produto_ID, _
                                             sStatusFaseRealocar, _
                                             sTipoAnotacao, Familia_Produto, _
                                             sStatusEncaminhado)
    
    End If
        
   ' Se ocorreu algum erro na DLL exibi-lo
    If Not bRetorno Then
        
        GoTo Finalizar
    
    End If
    
    ' Preenche a grid
    bRetorno = PreencheGrid()
    
    If Not bRetorno Then
        
        GoTo Finalizar
    
    End If
    
    objFormSelecionarOTS.cmdAnotacao.Enabled = True
    objFormSelecionarOTS.cmdImpOTS.Enabled = True
    objFormSelecionarOTS.CmdRealocar.Enabled = True
    
    If Familia_Produto = "I" Then
        objFormSelecionarOTS.CmdOE.Enabled = True
    End If
    
    ExibirListaOTS = True
    
    GoTo Finalizar
       
Finalizar:

    Screen.MousePointer = vbDefault
    
    If objMensagem.TipoMsg <> "" Then
    
        objMensagem.ExibeMensagem
        
    End If
    
    Set objDadosAcao = Nothing
    'Set objServico = Nothing
    Set objMensagem = Nothing
    
    Exit Function
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: ExibirListaOTS " & _
                            "da classe clsgenOC_SelecionarOTS"
    
    GoTo Finalizar

End Function
Public Function PreencheGrid() As Boolean

    Dim iLinha          As Integer
    Dim sStatus         As String
    Dim sRealocacao     As String
    
    '================================================
    '= Autor: Carlos Soares - 26/06/2003                                                    =
    '= Ação: Tarefa 04 - Referencia da classe clsblf_pendencia                      =
    '================================================
    Dim objDadosPendencia           As clsblf_Pendencia
    Dim rEspOTS                     As String
    Dim aListaTipoPendencia()       As Variant

    Dim objdadosServico             As New clsGenServico
    Dim bRetorno                    As Boolean
    'IP VPN
    Dim objDadosFinalizacao         As clsblf_AtivacaoFinalizacao
    Dim objDadosPreteste            As clsBlf_DadosPreativacao
    
    '============================
    '================================================
    '================================================
    
    'RN_SGAD_21148YBJ
    Dim objFase                             As clsgenFase
    Dim lIdCentroFuncionalFase              As Long
    Dim sDescricaoCentroFuncionalFase       As String
    
    
    'Nova coluna. IP VPN
    If (Familia_Produto = "I") And _
        (CodigoFase = GbCodigoFasePreAtivacao Or _
            CodigoFase = GbCodigoFaseAtivacao) Then
        
        'Classe de dados de Pendência
        Set objDadosPendencia = New clsblf_Pendencia
            
        ' Informa o banco e mensagem para a classe de fase
        objDadosPendencia.InformaBanco frmgenMDI_SGAPlus.objBanco
        objDadosPendencia.InformaMensagem objMensagem
        
        
        'Classe de dados da Fase Pré-Ativação
        Set objDadosPreteste = New clsBlf_DadosPreativacao
                
        ' Informa o banco e mensagem  para a classe de fase
        objDadosPreteste.InformaBanco frmgenMDI_SGAPlus.objBanco
        objDadosPreteste.InformaMensagem objMensagem
        
        'Classe de dados da Fase Ativação/Finalização/Retirada de Equipamentos
        Set objDadosFinalizacao = New clsblf_AtivacaoFinalizacao
                
        ' Informa o banco e mensagem para a classe de fase
        objDadosFinalizacao.InformaBanco frmgenMDI_SGAPlus.objBanco
        objDadosFinalizacao.InformaMensagem objMensagem
        
    End If
    
    
    ' Instância a classe de serviço
    Set objFase = New clsgenFase

    ' Informa o banco e mensagem para a classe de fase
    objFase.InformaBanco frmgenMDI_SGAPlus.objBanco
    objFase.InformaMensagem objMensagem
    
    iLinha = 0
    
    PreencheGrid = False
    
    'A o objeto de coleção da lista de OE's
    For Each objListaOTS In objServico.ListaOTS

        iLinha = iLinha + 1
        objFormSelecionarOTS.spdCircuitos.MaxRows = iLinha

        With objFormSelecionarOTS.spdCircuitos

            ' Atribui o valor da linha e define um range para cores
            .BlockMode = True
            .Row = iLinha
            .Row2 = iLinha
            .Col = 1
            .Col2 = .MaxCols
            
            'Verifica o status da fase, para definir a cor
            
            If CodigoFase = GbCodigoFaseAlocacao Then
            
                sRealocacao = objListaOTS.Correcao
                
            Else
            
                sRealocacao = ""
                
            End If
                
            If sRealocacao = "" Or sRealocacao = "N" Then
            
                If objListaOTS.StatusOTS = "FF" Then
                
                    .ForeColor = &HC000&  ' Verde
                    
                    sStatus = "Finalizado"
                    'CASO seja uma circuito temporario, exibi-lo na lista em destaque
                    If objListaOTS.IndicaCCTOTemporario = "S" Then
                    
                        sStatus = "Finalizado (URGENTE)"
                    
                    End If
                    
                ' Em Andamento Realizado
                ElseIf (objListaOTS.DataFimReal > "00:00:00") Then
                
                    .ForeColor = &HC00000  ' Azul
                    sStatus = "Em Andamento (Realizado)"
                    
                    'CASO seja uma circuito temporario, exibi-lo na lista em destaque
                    If objListaOTS.IndicaCCTOTemporario = "S" Then
                    
                        sStatus = "Realizado (URGENTE)"
                    
                    End If
                
                ' Atrasado
                ElseIf (Date > objListaOTS.DataFimPrevisto) Then
                
                    .ForeColor = &HFF&  ' Vermelho
                    
                    'Encaminhado atrasado
                    If objListaOTS.DataInicioReal = "00:00:00" Then
                       
                       sStatus = "Encaminhado (Atrasado)"
                    
                    Else
                    
                       sStatus = "Em andamento (Atrasado)"
                    
                    End If
                    
                    'CASO seja uma circuito temporario, exibi-lo na lista em destaque
                    If objListaOTS.IndicaCCTOTemporario = "S" Then
                    
                        sStatus = "Em Andamento (URGENTE)"
                    
                    End If
                
                ' Encaminhado
                
                ElseIf (objListaOTS.DataInicioReal = "00:00:00") Then
                      
                    .ForeColor = &H80000008  ' Preto
                    sStatus = "Encaminhado (A Realizar)"
                      
                ' Em Andamento à Realizar
                ElseIf (Date <= objListaOTS.DataFimPrevisto) Then

                    .ForeColor = &H80C0FF    ' Laranja
                    sStatus = "Em Andamento (A Realizar)"
                    
                    'CASO seja uma circuito temporario, exibi-lo na lista em destaque
                    If objListaOTS.IndicaCCTOTemporario = "S" Then
                    
                        sStatus = "Em andamento (URGENTE)"
                    
                    End If
                End If
            
            Else
            
                .ForeColor = &H40C0&      ' Tijolo
                sStatus = "Realocar"
            
            End If
                
            .BlockMode = False
            
            .Col = 1
            .Text = objListaOTS.NomeProduto

            ' Ação
            .Col = 2
            If objListaOTS.Acao = "ATV" Then
        
                .Text = "Ativação"
                
            ElseIf objListaOTS.Acao = "ALT" Then
            
                .Text = "Alteração"
                
            ElseIf objListaOTS.Acao = "DES" Then
            
                .Text = "Desativação"
                
            ElseIf objListaOTS.Acao = "CAN" Then
            
                .Text = "Cancelamento"
                
                If objListaOTS.Origem = "SGA" Then
                
                    .Text = "Cancelamento Operacional"
                
                End If
                
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


            If objListaOTS.ID_Produto = GbIDProdutoDiginet Then
        
                ' Número da OTS
                .Col = 3
                .Text = objListaOTS.CentroFuncional & _
                        " DNI " & _
                        CStr(objListaOTS.NumeroOTS) & _
                        "/" & _
                        objListaOTS.AnoOTS
            
            ElseIf objListaOTS.ID_Produto = GbIDProdutoEmbratelIPL Then
            
                ' Número da OTS
                .Col = 3
                .Text = objListaOTS.CentroFuncional & _
                        " IPL " & _
                        CStr(objListaOTS.NumeroOTS) & _
                        "/" & _
                        objListaOTS.AnoOTS
            Else
            
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
            End If
            
            'Se Produto Diginet na fase de Designacao será mostrado o nome da Operadora
            'no lugar da coluna de Designacao

            If ((objListaOTS.ID_Produto = GbIDProdutoDiginet) Or _
                (objListaOTS.ID_Produto = GbIDProdutoEmbratelIPL)) And _
               (CodigoFase = GbCodigoFaseDesignacao) Then
                ' Nome do Cliente
                .Col = 4
                .Text = objListaOTS.NomeCliente
                .ColWidth(4) = 22

                ' Nome da Operadora
                .Col = 5
                .Text = objListaOTS.NomeOperadora

            Else
                ' Designação do Circuito
                .Col = 4
                ' Dedicated - VPN RA - VPN Direct
                If (objListaOTS.ID_Produto = GbIDProdutoWID) Or _
                   (objListaOTS.ID_Produto = GbIDProdutoVPN_RA) Or _
                   (objListaOTS.ID_Produto = GbIDProdutoVPN_DIRECT) Then
                   
                   .Text = objListaOTS.CNL_Ponta_A & _
                           " WD " & _
                           Format(objListaOTS.NumeroDesignacao, "00000")
                
                ElseIf (objListaOTS.ID_Produto = GbIDProdutoDiginet) Then
                       
                       If objListaOTS.CNL_Ponta_A <> "" Or _
                          objListaOTS.CNL_Ponta_B <> "" Then
                          
                          .Text = objListaOTS.CNL_Ponta_A & " " & objListaOTS.CNL_Ponta_B & _
                               " " & objListaOTS.Codigo_NP & " " & _
                               Format(objListaOTS.NumeroDesignacao, "0000")
                           
                       End If
                
                ElseIf (objListaOTS.ID_Produto = GbIDProdutoEmbratelIPL) Then
                       
                       If objListaOTS.CNL_Ponta_A <> "" Or _
                          objListaOTS.CNL_Ponta_B <> "" Then
                          
                          .Text = objListaOTS.CNL_Ponta_A & " " & objListaOTS.CNL_Ponta_B & _
                               " " & objListaOTS.Codigo_NP & " " & _
                               Format(objListaOTS.NumeroDesignacao, "0000000")
                           
                       End If
                    
                ElseIf objListaOTS.ID_Produto = GbIDProdutoBLF Then
                   
                   .Text = objListaOTS.CNL_Ponta_A & _
                           "/BF/" & _
                           Format(objListaOTS.NumeroDesignacao, "00000")
                
                ElseIf objListaOTS.ID_Produto = GbIDProdutoBIV Then
                   'IP VPN
                   'Itens de Topologia e Extranet não têm Designação
                   If objListaOTS.CNL_Ponta_A <> "" Then
                   
                    .Text = objListaOTS.CNL_Ponta_A & _
                            "/VN/" & _
                            Format(objListaOTS.NumeroDesignacao, "00000")
                    Else
                        If (objListaOTS.TipoItem = GbItem_TOPOLOGIA) Then
                        
                            .Text = "ITEM DE TOPOLOGIA"
                            
                        ElseIf (objListaOTS.TipoItem = GbItem_EXTRNET) Then
                        
                            .Text = "ITEM DE EXTRANET"
                            
                        End If
                    End If
                    
                ElseIf ((objListaOTS.ID_Produto = GbIDProdutoJAH) Or objListaOTS.ID_Produto = GbIDProdutoVIP) Then
                   
                   If objListaOTS.CNL_Ponta_A <> "" Then
                   
                    '=============================================================
                    'Autor: Carlos Soares                                        =
                    'VipLineIP                                                   =
                    '=============================================================
                    If objListaOTS.ID_Produto = GbIDProdutoVIP Then
                        .ColWidth(3) = 21
                    End If
                    '=============================================================
                    '=============================================================
                   
                   
                    .Text = objListaOTS.CNL_Ponta_A & _
                            "/" & objListaOTS.Produto & "/" & _
                            Format(objListaOTS.NumeroDesignacao, "00000")
                                
                    '#MTLP - 17/11/04 - Alteração formato Designação JAh!
                    .Text = Replace(.Text, "/", "-")
                    '#MTLP - Fim
                    
                    End If
                    
                ElseIf (objListaOTS.ID_Produto = GbIDProdutoBLD Or _
                            objListaOTS.ID_Produto = GbIDProdutoVOIP) Then
                   
                   .Text = objListaOTS.CNL_Ponta_A & _
                           "/IP/" & _
                           Format(objListaOTS.NumeroDesignacao, "00000")
                
                ElseIf (objListaOTS.ID_Produto = GbIDProdutoNetPhone) Then
                   .Text = "(" & objListaOTS.Cd_Area_Telef & ") " & objListaOTS.Cd_Telefone
                End If
                
                ' Nome do Cliente
                .Col = 5
                .Text = objListaOTS.NomeCliente
            
            End If
            
            ' Data Prevista
            .Col = 6
            .Text = Format(objListaOTS.DataFimPrevisto, "dd/mm/yyyy")
        
            ' Status
            .Col = 7
            .Text = sStatus
            .ColHidden = False
            
            .Col = 8
            .Text = objListaOTS.ID_FASE
            
            .Col = 9
            .Text = objListaOTS.ID_SERVICO
            
            .Col = 10
            .Text = objListaOTS.ID_SERVICO_REDUNDANTE
            
            .Col = 11
            .Text = objListaOTS.ID_SERVICO_ANTERIOR
            
            .Col = 12
            .Text = objListaOTS.DataInicioPrevisto
            
            .Col = 13
            .Text = objListaOTS.Origem
            
            .Col = 14
            .Text = objListaOTS.DataFimPrevisto
            
            .Col = 15
            
            If (objListaOTS.DataInicioReal = "0:00:00" Or _
                    objListaOTS.DataInicioReal = "00:00:00") Then
                
                If (Familia_Produto = "I" And _
                     objListaOTS.DataFimReal = 0) And _
                    (CodigoFase = GbCodigoFasePreAtivacao Or _
                        CodigoFase = GbCodigoFaseAtivacao) Then
                    
                    rEspOTS = "Não iniciada na fase"
                
                End If
                '=======================================
                '=======================================
            
               .Text = ""
            
            Else
            
               .Text = objListaOTS.DataInicioReal
            
            End If
            
            .Col = 16
            
            If objListaOTS.DataFimReal = "0:00:00" Or objListaOTS.DataFimReal = "00:00:00" Then
            
                .Text = ""
            
            Else
            
                .Text = objListaOTS.DataFimReal
            
            End If
            
            .Col = 17
            .Text = objListaOTS.NumOE
            
            .Col = 18
            .Text = objListaOTS.ContratoID
            
            .Col = 19
            .Text = objListaOTS.AnoOE
            
            .Col = 20
            .Text = objListaOTS.ItemOE
            
            .Col = 21
            .Text = objListaOTS.ID_CIRCUITO
            
            .Col = 22
            .Text = objListaOTS.TipoCircuito
            
            .Col = 23
            .Text = objListaOTS.Sequencia_Acao
            
            .Col = 24
            .Text = objListaOTS.IndicaCCTOTemporario
            
            .Col = 25
            .Text = objListaOTS.NrLocalInstalacao
            
            .Col = 26
            .Text = objListaOTS.ID_Produto
            
            .Col = 27
            .Text = objListaOTS.ID_ACAO
            
            .Col = 28
            .Text = objListaOTS.Acao
            
            .Col = 29
            .Text = objListaOTS.Produto
            
            .Col = 30
            .Text = objListaOTS.UBR
            
            .Col = 31
            .Text = objListaOTS.OTS
            
            .Col = 32
            .Col = 33
            .Text = objListaOTS.CNL_Ponta_A
            
            .Col = 34
            .Text = objListaOTS.CNL_Ponta_B
            
            .Col = 35
            .Text = objListaOTS.Codigo_NP
            
            .Col = 36
            .Text = objListaOTS.NumeroDesignacao
            
            .Col = 37
            .Text = objListaOTS.ContaCliente
            
            .Col = 38
            .Text = objListaOTS.NomeCliente
            
            .Col = 39
            .Text = objListaOTS.PermiteExecucao
            
            .Col = 40
            .Text = objListaOTS.NumContrato
            
            .Col = 41
            .Text = objListaOTS.AnoContrato
            
            .Col = 42
            .Text = objListaOTS.VersaoContrato
            
            .Col = 43
            .Text = objListaOTS.DataOTS

            .Col = 44
            .Text = objListaOTS.TipoAcesso

            .Col = 45
            .Text = objListaOTS.Correcao
        
            'IP VPN & JAH
            .Col = 47
            .Text = objListaOTS.TipoItem
        
            .Col = 48
            .Text = objListaOTS.DataPrevATVContrato

            .Col = 50
            .Text = ("" & objListaOTS.NomeGuerra)

'''
            '======================================================
            '= Autor: Carlos Soares - 26/06/2003                                                =
            '= Ação: Tarefa 04 - Inclusão da carga da coluna do grid que será apresentada       =
            '=          nas fases de pré-ativação ou na fase de ativação/finalização            =
            '======================================================
            If (Familia_Produto = "I" And _
                     objListaOTS.DataFimReal = 0) And _
                (CodigoFase = GbCodigoFasePreAtivacao Or _
                    CodigoFase = GbCodigoFaseAtivacao) Then
            
                'Verifica se a OTS tem pendência em aberto. Retorna True caso tenha
                If objDadosPendencia.VerificaPendencia(objListaOTS.ID_SERVICO, _
                        "A") = True Then
                        
                    rEspOTS = "Circuito com pendência"
                                
                End If
                
                If objMensagem.TipoMsg <> "" Then
                    objMensagem.ExibeMensagem
                End If

            If (objListaOTS.DataFimReal = 0 And _
                    rEspOTS = "") And _
                (CodigoFase = GbCodigoFasePreAtivacao) Then
                
                    'Se for uma pre_ativacao
                     If objDadosPreteste.RecuperarDados(objListaOTS.ID_FASE) Then
            
                        If (objDadosPreteste.DataProgramacaoPreTeste <> "" And _
                                objDadosPreteste.HoraProgramacaoPreTeste <> "") Then
                            
                            rEspOTS = "Possui Programação de pré-teste"
                            
                        End If
                     End If
                
                ElseIf (CodigoFase = GbCodigoFaseAtivacao And _
                            rEspOTS = "" And _
                            objListaOTS.DataFimReal = 0) Then
                    
                    objDadosFinalizacao.RecuperarDados (objListaOTS.ID_FASE)
                        
                    If (objListaOTS.Acao = "DES" And _
                        objListaOTS.Acao = "CAN") Then
                            
                        If (objDadosFinalizacao.DataProgramadaRetirada <> "" And _
                                objDadosFinalizacao.HoraProgramadaRetirada <> "") Then
                                
                            rEspOTS = "Possui Programação de Retirada dos Equipamentos"
                            
                        End If
    
                    Else
                        If (objDadosFinalizacao.DataReprogramada <> "" And _
                                objDadosFinalizacao.HoraReprogramada <> "") Then
                        
                            rEspOTS = "Possui Programação de Finalização"
                            
                        End If
                    End If
                End If
            '=========================================
            'Carlos Soares 25/02/2006                =
            '=========================================
            .Col = 49
            .ColWidth(49) = 30
            
            If objDadosFinalizacao.FlagExecucaoSart = "S" Then
                rEspOTS = IIf(Trim(rEspOTS) = "", "Repassado ao SART", rEspOTS & " - Repassado ao SART")
            End If
            '=========================================
            '=========================================
            .Text = rEspOTS
            
            rEspOTS = ""
    
            End If
            '======================================================

            'RN_SGAD_21148YBJ
            If (GbFlagRN_SGAD_21148YBJ) Then
            
                bRetorno = objFase.BuscarCentroFuncionalFase(objListaOTS.ID_SERVICO, _
                                objListaOTS.ID_FASE, _
                                lIdCentroFuncionalFase, _
                                sDescricaoCentroFuncionalFase)
                

                .Col = 51
                .Text = CLng(lIdCentroFuncionalFase)
                
                .Col = 52
                .Text = (sDescricaoCentroFuncionalFase)
                
                'MTLP: inclusão do ID do Centro Funcional da Fase
                .Col = 53
                .Text = ("" & objListaOTS.CentroFuncional_Fase_ID)

                .Col = 54
                .Text = "" '("" & objListaOTS.Cd_Area_Telef)

                .Col = 55
                .Text = "" '("" & objListaOTS.Cd_Telefone)

                .Col = 56
                .Text = ("" & objListaOTS.NumeroItemATEN)

'                .Col = 57
'                .Text = ("" & objListaOTS.AnoItemATEN)

            End If
            
        End With
    Next
    
    objFormSelecionarOTS.spdCircuitos.Row = 0
    objFormSelecionarOTS.spdCircuitos.Col = 7
    objFormSelecionarOTS.spdCircuitos.Text = "Status da Fase"
    objFormSelecionarOTS.spdCircuitos.ColWidth(7) = 22.88
    
    objFormSelecionarOTS.CmdExcel.Enabled = True
    
    
    Set objDadosPendencia = Nothing
    Set objDadosPreteste = Nothing
    Set objDadosFinalizacao = Nothing
    
    'RN_SGAD_21148YBJ
    Set objFase = Nothing
    
    PreencheGrid = True
    
    Exit Function
    

End Function
Public Function AtribuiPropriedadesFase() As Boolean

On Error GoTo ErrorHandler:
    
    AtribuiPropriedadesFase = False
    
    ' Intância a classe de mensagem
    Set objMensagem = New clsgenMensagem
    
    objFormSelecionarOTS.spdCircuitos.Row = objFormSelecionarOTS.spdCircuitos.ActiveRow
    objFormSelecionarOTS.spdCircuitos.Col = 8
    
    objFase.ID_FASE = CLng("0" & objFormSelecionarOTS.spdCircuitos.Text)
    
    lFase_Id = objFase.ID_FASE

    objFormSelecionarOTS.spdCircuitos.Col = 9
    objFase.ID_SERVICO = CLng("0" & objFormSelecionarOTS.spdCircuitos.Text)
    lServico_ID = objFase.ID_SERVICO

    objFormSelecionarOTS.spdCircuitos.Col = 11
    objFase.ID_SERVICO_ANTERIOR = CLng("0" & objFormSelecionarOTS.spdCircuitos.Text)
    lServico_ID_Anterior = objFase.ID_SERVICO_ANTERIOR

    objFormSelecionarOTS.spdCircuitos.Col = 10
    objFase.ID_SERVICO_REDUNDANTE = CLng("0" & objFormSelecionarOTS.spdCircuitos.Text)

    objFormSelecionarOTS.spdCircuitos.Col = 12
    objFase.Data_Inicio_Previsto = objFormSelecionarOTS.spdCircuitos.Text
    sDataInicioPrevisto = Format(objFase.Data_Inicio_Previsto, "dd/mm/yyyy")

    objFormSelecionarOTS.spdCircuitos.Col = 13
    objFase.Origem = objFormSelecionarOTS.spdCircuitos.Text
    
    objFormSelecionarOTS.spdCircuitos.Col = 14
    objFase.Data_Fim_Previsto = objFormSelecionarOTS.spdCircuitos.Text

    objFormSelecionarOTS.spdCircuitos.Col = 15
    
    If objFormSelecionarOTS.spdCircuitos.Text = "" Then
    
        objFase.Data_Inicio_Real = "00:00:00"
         
    Else
    
        objFase.Data_Inicio_Real = objFormSelecionarOTS.spdCircuitos.Text
    
    End If

    objFormSelecionarOTS.spdCircuitos.Col = 16

    If objFormSelecionarOTS.spdCircuitos.Text = "" Then

       sDataFimReal = ""
       objFase.Data_Fim_Real = "00:00:00"
    Else
       
       objFase.Data_Fim_Real = Format(objFormSelecionarOTS.spdCircuitos.Text, "dd/mm/yyyy")
       sDataFimReal = Format(objFormSelecionarOTS.spdCircuitos.Text, "dd/mm/yyyy")

    End If

    objFormSelecionarOTS.spdCircuitos.Col = 17
    objFase.NumOE = CLng("0" & objFormSelecionarOTS.spdCircuitos.Text)

    lNumeroOE = objFase.NumOE

    objFormSelecionarOTS.spdCircuitos.Col = 18
    lContratoID = CLng("0" & objFormSelecionarOTS.spdCircuitos.Text)
    
    objFormSelecionarOTS.spdCircuitos.Col = 19
    objFase.AnoOE = CInt("0" & objFormSelecionarOTS.spdCircuitos.Text)
    lAnoOE = objFase.AnoOE
         
    objFormSelecionarOTS.spdCircuitos.Col = 20
    objFase.ItemOE = CLng("0" & objFormSelecionarOTS.spdCircuitos.Text)
    lNumeroITemOE = objFase.ItemOE

    objFormSelecionarOTS.spdCircuitos.Col = 21
    objFase.ID_CIRCUITO = CLng("0" & objFormSelecionarOTS.spdCircuitos.Text)
    lCircuitoID = objFase.ID_CIRCUITO

    objFormSelecionarOTS.spdCircuitos.Col = 22
    objFase.TipoCircuito = objFormSelecionarOTS.spdCircuitos.Text

    objFormSelecionarOTS.spdCircuitos.Col = 23
    lSequencia_Acao = CLng("0" & objFormSelecionarOTS.spdCircuitos.Text)

    objFormSelecionarOTS.spdCircuitos.Col = 24
    sIndicaCCTOTemp = objFormSelecionarOTS.spdCircuitos.Text
         
'    lNrLocalInsta = objServico.ListaOTS(objFormSelecionarOTS.spdCircuitos.ActiveRow).NrLocalInstalacao
    objFormSelecionarOTS.spdCircuitos.Col = 25
    lNrLocalInsta = CLng("0" & objFormSelecionarOTS.spdCircuitos.Text)

    objFase.SequenciaFase = iSequenciaFase

    objFormSelecionarOTS.spdCircuitos.Col = 26
    objFase.ID_Produto = CLng("0" & objFormSelecionarOTS.spdCircuitos.Text)
    Produto_ID = objFase.ID_Produto

    objFormSelecionarOTS.spdCircuitos.Col = 27
    objFase.ID_ACAO = CLng("0" & objFormSelecionarOTS.spdCircuitos.Text)
    
    lAcao_ID = objFase.ID_ACAO

    objFormSelecionarOTS.spdCircuitos.Col = 28
    objFase.Acao = objFormSelecionarOTS.spdCircuitos.Text
    sSigla_Acao = objFase.Acao

    objFormSelecionarOTS.spdCircuitos.Col = 29
    objFase.Produto = objFormSelecionarOTS.spdCircuitos.Text

    objFormSelecionarOTS.spdCircuitos.Col = 30
    objFase.UBR = objFormSelecionarOTS.spdCircuitos.Text

    objFormSelecionarOTS.spdCircuitos.Col = 31
    objFase.OTS = objFormSelecionarOTS.spdCircuitos.Text

    objFormSelecionarOTS.spdCircuitos.Col = 33
    objFase.Designacao = objFormSelecionarOTS.spdCircuitos.Text

    objFormSelecionarOTS.spdCircuitos.Col = 37
    objFase.ContaCliente = objFormSelecionarOTS.spdCircuitos.Text
    
    objFormSelecionarOTS.spdCircuitos.Col = 38
    objFase.NomeCliente = objFormSelecionarOTS.spdCircuitos.Text

    'IP VPN
    objFormSelecionarOTS.spdCircuitos.Col = 47
    objFase.TipoItem = objFormSelecionarOTS.spdCircuitos.Text

    
    objFormSelecionarOTS.spdCircuitos.Col = 48
    objFase.DataPrevATVContrato = objFormSelecionarOTS.spdCircuitos.Text
    
    AtribuiPropriedadesFase = True
    
    GoTo Finalizar
        
Finalizar:

    Screen.MousePointer = vbDefault
    On Error GoTo 0
    
    If objMensagem.TipoMsg <> "" Then
    
        objMensagem.ExibeMensagem
        
    End If
    
    Exit Function
        
ErrorHandler:

    If Err.Number <> 0 Then
       objMensagem.TipoMsg = "E"
       objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                   "Houve um erro de execução no Método: " & vbCrLf & vbCrLf & "AtribuiPropriedadesFase " & _
                   "da classe clsGenOC_SelecionarOTS"
    End If
       
    objMensagem.ExibeMensagem
       
    GoTo Finalizar

End Function

Public Function GravarAnotacao() As Boolean

On Error GoTo ErrorHandler:

    Dim bRetorno As Boolean
    
    GravarAnotacao = False
    
    Screen.MousePointer = vbArrowHourglass

    objFase.Anotacao = Anotacao
    
    ' Inicia a transação
    bRetorno = frmgenMDI_SGAPlus.objBanco.Iniciar_Transacao()
    
        ' Se ocorreu algum erro na DLL
        If Not bRetorno Then
        
            GoTo Finalizar
            
        End If
    
        ' Informa o banco para a classe de fase
        objFase.InformaBanco frmgenMDI_SGAPlus.objBanco
        
        ' Informa a mensagem para a classe de fase
        objFase.InformaMensagem objMensagem
        
        ' Grava a anotação da fase
        
        'Coluna 8 ===> ID FASE
        objFormSelecionarOTS.spdCircuitos.Row = objFormSelecionarOTS.spdCircuitos.ActiveRow
        objFormSelecionarOTS.spdCircuitos.Col = 8
        
        bRetorno = objFase.GravarAnotacao(objFormSelecionarOTS.spdCircuitos.Text, _
                                          objFase.Anotacao)
        
        ' Ocorreu um erro ao executar o método
        If Not bRetorno Then
    
            frmgenMDI_SGAPlus.objBanco.Cancelar_Transacao
    
            GoTo Finalizar

        End If
        
        ' Atualiza a mensagem de retorno como gravação realizada com sucesso
        objMensagem.TipoMsg = "A"
        objMensagem.Descricao = "Gravação realizada com sucesso"
    
    
    ' Finaliza a transação
    frmgenMDI_SGAPlus.objBanco.Finalizar_Transacao
    
    GravarAnotacao = True
    
Finalizar:

    Screen.MousePointer = vbDefault
    
    If objMensagem.TipoMsg <> "" Then
    
        objMensagem.ExibeMensagem
        
    End If
    
    Set objMensagem = Nothing
    
    On Error GoTo 0
    
    Exit Function
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: GravarAnotacao " & _
                            "da classe clsgenOC_SelecionarOTS"
    
    GoTo Finalizar

End Function

Public Function GravarAnotacaoGen(pbControlaTransacao As Boolean, _
                                  Optional psTipoAnotacao As String) As Boolean

On Error GoTo ErrorHandler:

    Dim bRetorno                   As Boolean
    Dim objDetalheAnotacao         As New clsGenDetalheAnotacao
    Dim objDadosFase               As clsgenFase
    Dim objdadosServico            As clsGenServico
    Dim lIDFase                    As Long
    Dim lIDServico                 As Long
    Dim FlagIniciouTransacao       As Boolean
    
    GravarAnotacaoGen = False
    
    FlagIniciouTransacao = False
    
    If objMensagem Is Nothing Then
        
        Set objMensagem = New clsgenMensagem
        
    End If
    
    Screen.MousePointer = vbArrowHourglass

    objDetalheAnotacao.InformaBanco frmgenMDI_SGAPlus.objBanco
    objDetalheAnotacao.InformaMensagem objMensagem
    
    If pbControlaTransacao Then
        
        ' Inicia a transação
        bRetorno = frmgenMDI_SGAPlus.objBanco.Iniciar_Transacao()
    
        ' Se ocorreu algum erro na DLL
        If Not bRetorno Then
        
            GoTo Finalizar
            
        End If
    
        FlagIniciouTransacao = True
    
    End If
    
    If psTipoAnotacao = "" Then
    
        psTipoAnotacao = "A"
        
    End If
    
    'Coluna 9 ===> ID SERVICO
    objFormSelecionarOTS.spdCircuitos.Row = objFormSelecionarOTS.spdCircuitos.ActiveRow
    objFormSelecionarOTS.spdCircuitos.Col = 9
    
    lIDServico = CLng(objFormSelecionarOTS.spdCircuitos.Text)
    
    'Coluna 8 ===> ID FASE
    objFormSelecionarOTS.spdCircuitos.Row = objFormSelecionarOTS.spdCircuitos.ActiveRow
    objFormSelecionarOTS.spdCircuitos.Col = 8
    
    lIDFase = CLng(objFormSelecionarOTS.spdCircuitos.Text)
    
    objDetalheAnotacao.ID = CLng(objFormSelecionarOTS.spdCircuitos.Text)

    objDetalheAnotacao.Anotacao = TrataPlick(Anotacao)
    
    If Not objDetalheAnotacao.FaseExistente(objDetalheAnotacao.ID, psTipoAnotacao) Then
    
        bRetorno = objDetalheAnotacao.IncluirDados(psTipoAnotacao)
        
    Else
    
        bRetorno = objDetalheAnotacao.AtualizarDados(psTipoAnotacao)
    
    End If
        
    If objMensagem.TipoMsg <> "" Then
    
        GoTo ErrorHandler
        
    End If
    
    'Caso o Tipo de Anotacao de Realocacao, o sistema deverá atualizar o Atributo Realocacao
    'da tabela de Fase
    
    If psTipoAnotacao = "R" Or psTipoAnotacao = "C" Then
    
        Set objDadosFase = New clsgenFase
        
        objDadosFase.InformaBanco frmgenMDI_SGAPlus.objBanco
        objDadosFase.InformaMensagem objMensagem
        
        bRetorno = objDadosFase.GravarCorrecao(lIDFase, psTipoAnotacao)
            
        If objMensagem.TipoMsg <> "" Then
        
            GoTo ErrorHandler
            
        End If
            
        bRetorno = objDadosFase.RetirarPeriodoPrevisto(lIDFase)
            
        If objMensagem.TipoMsg <> "" Then
        
            GoTo ErrorHandler
            
        End If
        
        Set objdadosServico = New clsGenServico
        
        objdadosServico.InformaBanco frmgenMDI_SGAPlus.objBanco
        objdadosServico.InformaMensagem objMensagem
        
        bRetorno = objdadosServico.GravarCorrecao(lIDServico, psTipoAnotacao)
            
    End If
    
    If pbControlaTransacao Then
    
        ' Ocorreu um erro ao executar o método
        If Not bRetorno Then
    
            GoTo ErrorHandler

        End If
    
    
        ' Atualiza a mensagem de retorno como gravação realizada com sucesso
        objMensagem.TipoMsg = "A"
        '= Carlos Soares 25/06/2003 - Deslocado para baixo =============================
        'objMensagem.Descricao = "Gravação realizada com sucesso"
        '===============================================================================
        
        '===============================================================================
        '= Autor: Carlos Soares - 25/06/2003                                           =
        '= Ação: Verifica se a atualização está nula, e se estiver exclui da tabela    =
        '= anotacao                                                                    =
        '===============================================================================
        If Trim(Anotacao) = "" Then
            objMensagem.Descricao = "Atualização realizada com sucesso"
            
        Else
            objMensagem.Descricao = "Gravação realizada com sucesso"
            
        End If
        '===============================================================================
        '===============================================================================
    
        ' Finaliza a transação
        frmgenMDI_SGAPlus.objBanco.Finalizar_Transacao
    
    End If
    
    GravarAnotacaoGen = True
    
Finalizar:

    Screen.MousePointer = vbDefault
    Set objDetalheAnotacao = Nothing
    
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

    If objMensagem.TipoMsg = "" Then
        
        objMensagem.TipoMsg = "E"
        
        objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                                "Houve um erro de execução no método: GravarAnotacao " & _
                                "da classe clsgenOC_SelecionarOTS"
        
    End If
    
    objMensagem.ExibeMensagem
    
    GoTo Finalizar

End Function

Public Function GravarAnotacaoMonitoracao(plServicoId As Long, _
                                          plFaseID As Long, _
                                          pbControlaTransacao As Boolean, _
                                  Optional psTipoAnotacao As String) As Boolean

On Error GoTo ErrorHandler:

    Dim bRetorno                   As Boolean
    Dim objDetalheAnotacao         As New clsGenDetalheAnotacao
    Dim objDadosFase               As clsgenFase
    Dim objdadosServico            As clsGenServico
    Dim FlagIniciouTransacao       As Boolean
    
    GravarAnotacaoMonitoracao = False
    
    FlagIniciouTransacao = False
    
    If objMensagem Is Nothing Then
        
        Set objMensagem = New clsgenMensagem
        
    End If
    
    Screen.MousePointer = vbArrowHourglass

    objDetalheAnotacao.InformaBanco frmgenMDI_SGAPlus.objBanco
    objDetalheAnotacao.InformaMensagem objMensagem
    
    If pbControlaTransacao Then
        
        ' Inicia a transação
        bRetorno = frmgenMDI_SGAPlus.objBanco.Iniciar_Transacao()
    
        ' Se ocorreu algum erro na DLL
        If Not bRetorno Then
        
            GoTo Finalizar
            
        End If
    
        FlagIniciouTransacao = True
    
    End If
    
    If psTipoAnotacao = "" Then
    
        psTipoAnotacao = "A"
        
    End If
    
    objDetalheAnotacao.ID = plFaseID

    objDetalheAnotacao.Anotacao = TrataPlick(Anotacao)
    
    If Not objDetalheAnotacao.FaseExistente(objDetalheAnotacao.ID, psTipoAnotacao) Then
    
        bRetorno = objDetalheAnotacao.IncluirDados(psTipoAnotacao)
        
    Else
    
        bRetorno = objDetalheAnotacao.AtualizarDados(psTipoAnotacao)
    
    End If
        
    If objMensagem.TipoMsg <> "" Then
    
        GoTo ErrorHandler
        
    End If
    
    'Caso o Tipo de Anotacao de Realocacao, o sistema deverá atualizar o Atributo Realocacao
    'da tabela de Fase
    
    If psTipoAnotacao = "R" Or psTipoAnotacao = "C" Then
    
        Set objDadosFase = New clsgenFase
        
        objDadosFase.InformaBanco frmgenMDI_SGAPlus.objBanco
        objDadosFase.InformaMensagem objMensagem
        
        bRetorno = objDadosFase.GravarCorrecao(plFaseID, psTipoAnotacao)
            
        If objMensagem.TipoMsg <> "" Then
        
            GoTo ErrorHandler
            
        End If
            
        bRetorno = objDadosFase.RetirarPeriodoPrevisto(plFaseID)
            
        If objMensagem.TipoMsg <> "" Then
        
            GoTo ErrorHandler
            
        End If
        
        Set objdadosServico = New clsGenServico
        
        objdadosServico.InformaBanco frmgenMDI_SGAPlus.objBanco
        objdadosServico.InformaMensagem objMensagem
        
        bRetorno = objdadosServico.GravarCorrecao(plServicoId, psTipoAnotacao)
            
    End If
    
    If pbControlaTransacao Then
    
        ' Ocorreu um erro ao executar o método
        If Not bRetorno Then
    
            GoTo ErrorHandler

        End If
    
    
        ' Atualiza a mensagem de retorno como gravação realizada com sucesso
        objMensagem.TipoMsg = "A"
        '= Carlos Soares 25/06/2003 - Deslocado para baixo =============================
        'objMensagem.Descricao = "Gravação realizada com sucesso"
        '===============================================================================
        
        '===============================================================================
        '= Autor: Carlos Soares - 25/06/2003                                           =
        '= Ação: Verifica se a atualização está nula, e se estiver exclui da tabela    =
        '= anotacao                                                                    =
        '===============================================================================
        If Trim(Anotacao) = "" Then
            objMensagem.Descricao = "Atualização realizada com sucesso"
            
        Else
            objMensagem.Descricao = "Gravação realizada com sucesso"
            
        End If
        '===============================================================================
        '===============================================================================
    
        ' Finaliza a transação
        frmgenMDI_SGAPlus.objBanco.Finalizar_Transacao
    
    End If
    
    GravarAnotacaoMonitoracao = True
    
Finalizar:

    Screen.MousePointer = vbDefault
    Set objDetalheAnotacao = Nothing
    
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

    If objMensagem.TipoMsg = "" Then
        
        objMensagem.TipoMsg = "E"
        
        objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                                "Houve um erro de execução no método: GravarAnotacaoMonitoracao " & _
                                "da classe clsgenOC_SelecionarOTS"
        
    End If
    
    objMensagem.ExibeMensagem
    
    GoTo Finalizar

End Function

Public Function GravarJustificativaCorrecao(Optional psTipoCorrecao As String) As Boolean

On Error GoTo ErrorHandler:

    Dim bRetorno                   As Boolean
    Dim objDadosCtrlCorrecao       As New clsGenDadosControleCorrecao
    Dim objDadosFase               As clsgenFase
    Dim objdadosServico            As clsGenServico
    Dim objDadosIntegracaoBlf      As clsBlf_IntegracaoWebIP
    Dim lIDFase                    As Long
    Dim lIDServico                 As Long
    Dim FlagManterPeriodoSeguinte  As Boolean
    Dim sDataIniPrev               As String
    Dim sDataFimPrev               As String
    Dim FlagIniciouTransacao       As Boolean
    
    GravarJustificativaCorrecao = False
    
    FlagIniciouTransacao = False
    
    Screen.MousePointer = vbArrowHourglass

    If objMensagem Is Nothing Then
        
        Set objMensagem = New clsgenMensagem
        
    End If
    
    objDadosCtrlCorrecao.InformaBanco frmgenMDI_SGAPlus.objBanco
    objDadosCtrlCorrecao.InformaMensagem objMensagem
    
    ' Inicia a transação
    bRetorno = frmgenMDI_SGAPlus.objBanco.Iniciar_Transacao()

    ' Se ocorreu algum erro na DLL
    If Not bRetorno Then
        
        GoTo Finalizar
        
    End If
    
    FlagIniciouTransacao = True
    
    'Coluna 12 ===> Data inicio Previsto
    objFormSelecionarOTS.spdCircuitos.Row = objFormSelecionarOTS.spdCircuitos.ActiveRow
    objFormSelecionarOTS.spdCircuitos.Col = 12
    
    sDataIniPrev = objFormSelecionarOTS.spdCircuitos.Text
    
    'Coluna 14 ===> Data Fim Previsto
    objFormSelecionarOTS.spdCircuitos.Row = objFormSelecionarOTS.spdCircuitos.ActiveRow
    objFormSelecionarOTS.spdCircuitos.Col = 14
    
    sDataFimPrev = objFormSelecionarOTS.spdCircuitos.Text
    
    'Coluna 9 ===> ID SERVICO
    objFormSelecionarOTS.spdCircuitos.Row = objFormSelecionarOTS.spdCircuitos.ActiveRow
    objFormSelecionarOTS.spdCircuitos.Col = 9
    
    lIDServico = CLng(objFormSelecionarOTS.spdCircuitos.Text)
    
    'Dados do Serviço
    Set objServico = New clsGenServico

    objServico.InformaBanco frmgenMDI_SGAPlus.objBanco
    objServico.InformaMensagem objMensagem
    
    bRetorno = objServico.RecuperarDados(lIDServico)
    
    If Not bRetorno Then
        GoTo Finalizar
    End If
    
    
    'Coluna 8 ===> ID FASE
    objFormSelecionarOTS.spdCircuitos.Row = objFormSelecionarOTS.spdCircuitos.ActiveRow
    objFormSelecionarOTS.spdCircuitos.Col = 8
    
    lIDFase = CLng(objFormSelecionarOTS.spdCircuitos.Text)
    
    objDadosCtrlCorrecao.FaseID = CLng(objFormSelecionarOTS.spdCircuitos.Text)

    objDadosCtrlCorrecao.Justificativa = TrataPlick(Anotacao)
    objDadosCtrlCorrecao.DataSolicitacao = Format(Date, "DD/MM/YYYY")
    
    objDadosCtrlCorrecao.DataInicioPrevAnt = "" & sDataIniPrev
    objDadosCtrlCorrecao.DataFimPrevAnt = "" & sDataFimPrev
    objDadosCtrlCorrecao.TipoCorrecao = psTipoCorrecao
    objDadosCtrlCorrecao.UsuarioIDSolic = GbID_Usuario
    
    If Not objDadosCtrlCorrecao.FaseExistente(lIDFase, psTipoCorrecao) Then
    
        bRetorno = objDadosCtrlCorrecao.IncluirDados(psTipoCorrecao)
        
    Else
    
        bRetorno = objDadosCtrlCorrecao.AtualizarDados(lIDFase, psTipoCorrecao)
    
    End If
        
    If Not bRetorno Then
    
        GoTo ErrorHandler
        
    End If
    
    Set objDadosFase = New clsgenFase
    
    objDadosFase.InformaBanco frmgenMDI_SGAPlus.objBanco
    objDadosFase.InformaMensagem objMensagem
    
    bRetorno = objDadosFase.GravarCorrecao(lIDFase, psTipoCorrecao)
        
    If objMensagem.TipoMsg <> "" Then
    
        GoTo ErrorHandler
        
    End If
        
    FlagManterPeriodoSeguinte = True
    
    bRetorno = objDadosFase.RetirarPeriodoPrevisto(lIDFase)
        
    If objMensagem.TipoMsg <> "" Then
    
        GoTo ErrorHandler
        
    End If
    
    Set objdadosServico = New clsGenServico
    
    objdadosServico.InformaBanco frmgenMDI_SGAPlus.objBanco
    objdadosServico.InformaMensagem objMensagem
    
    bRetorno = objdadosServico.GravarCorrecao(lIDServico, psTipoCorrecao)
    
    If Familia_Produto = "I" Then
       
       Set objDadosIntegracaoBlf = New clsBlf_IntegracaoWebIP
    
       objDadosIntegracaoBlf.InformaBanco frmgenMDI_SGAPlus.objBanco
       objDadosIntegracaoBlf.InformaMensagem objMensagem
       
       If (Produto_ID = GbIDProdutoBIV And _
            objServico.TipoItem = GbItem_SITE) Or _
            (Produto_ID <> GbIDProdutoBIV) Then
       
            bRetorno = objDadosIntegracaoBlf.RetirarOTSListaWeb(lIDServico)
       
       Else
            'Confirmando com o SPRADO se há Correção para itens de Topologia e Extranet
            objMensagem.TipoMsg = "A"
            objMensagem.Descricao = "Correção não habilitada para este Tipo de Item"
            GoTo ErrorHandler
            'bRetorno = False
       End If

    End If
    
    ' Ocorreu um erro ao executar o método
    If Not bRetorno Then
        frmgenMDI_SGAPlus.objBanco.Cancelar_Transacao
        GoTo Finalizar
    End If

    ' Atualiza a mensagem de retorno como gravação realizada com sucesso
    objMensagem.TipoMsg = "A"
    objMensagem.Descricao = "Gravação realizada com sucesso"

    ' Finaliza a transação
    frmgenMDI_SGAPlus.objBanco.Finalizar_Transacao
    
    GravarJustificativaCorrecao = True
    
Finalizar:

    Screen.MousePointer = vbDefault
    
    Set objDadosCtrlCorrecao = Nothing
    
    Set objMensagem = Nothing
    
    On Error GoTo 0
    
    Exit Function
        
ErrorHandler:

    If FlagIniciouTransacao Then
       frmgenMDI_SGAPlus.objBanco.Cancelar_Transacao
    End If
    
    If objMensagem.TipoMsg = "" Then
        
        objMensagem.TipoMsg = "E"
        objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                                "Houve um erro de execução no método: GravarAnotacao " & _
                                "da classe clsgenOC_SelecionarOTS"
    End If
    
    objMensagem.ExibeMensagem
    
    GoTo Finalizar

End Function
Public Sub MontarImpressaoOTS()

On Error GoTo ErrorHandler:

    Dim bRetorno                As Boolean
    Dim lIDServicoImp           As Long
    Dim lIDServicoAntImp        As Long
    Dim sAcaoImp                As String

    Screen.MousePointer = vbArrowHourglass

    objFormSelecionarOTS.spdCircuitos.Row = objFormSelecionarOTS.spdCircuitos.ActiveRow
    objFormSelecionarOTS.spdCircuitos.Col = 9
    lIDServicoImp = CLng(objFormSelecionarOTS.spdCircuitos.Text)
    
    objFormSelecionarOTS.spdCircuitos.Row = objFormSelecionarOTS.spdCircuitos.ActiveRow
    objFormSelecionarOTS.spdCircuitos.Col = 11
    lIDServicoAntImp = CLng(objFormSelecionarOTS.spdCircuitos.Text)
    
    objFormSelecionarOTS.spdCircuitos.Row = objFormSelecionarOTS.spdCircuitos.ActiveRow
    objFormSelecionarOTS.spdCircuitos.Col = 2
    sAcaoImp = objFormSelecionarOTS.spdCircuitos.Text
    
    objFormSelecionarOTS.spdCircuitos.Row = objFormSelecionarOTS.spdCircuitos.ActiveRow
    objFormSelecionarOTS.spdCircuitos.Col = 26
    Produto_ID = CLng(objFormSelecionarOTS.spdCircuitos.Text)
    
    
    ' Intância a classe de mensagem
    Set objMensagem = New clsgenMensagem
    
    Call Imprime_OTS(Produto_ID, _
                lIDServicoImp, _
                lIDServicoAntImp, _
                sAcaoImp)
                
                
    
''    If (Produto_ID = GbIDProdutoWID) Or _
''       (Produto_ID = GbIDProdutoVPN_RA) Or _
''       (Produto_ID = GbIDProdutoVPN_DIRECT) Or _
''       (Produto_ID = GbIDProdutoBLF) Then
''
''        bRetorno = OTS_ProdInternet(Produto_ID, _
''                                     lIDServicoImp, _
''                                    lIDServicoAntImp, _
''                                    sAcaoImp)
''
''    ElseIf (Produto_ID = GbIDProdutoDiginet Or Produto_ID = GbIDProdutoEmbratelIPL) Then
''
''        bRetorno = OTS_ProdDados(Produto_ID, _
''                                 lIDServicoImp, _
''                                 lIDServicoAntImp, _
''                                 sAcaoImp)
''
''    '================================================
''    '= Autor: Carlos Soares - 30/07/2003                                                    =
''    '= Ação: Implementacao de impressão OTS para IP VPN                          =
''    '================================================
''    ElseIf (Produto_ID = GbIDProdutoBIV) Then
''
''        bRetorno = OTS_ProdIPVPN(Produto_ID, _
''                                 lIDServicoImp, _
''                                 lIDServicoAntImp, _
''                                 sAcaoImp)
''
''    ElseIf (Produto_ID = GbIDProdutoJAH) Then
''
''        Call OTS_ProdJAH(Produto_ID, _
''                                 lIDServicoImp, _
''                                 lIDServicoAntImp, _
''                                 sAcaoImp)
''
''    '================================================
''    '================================================
''
''    End If
    
    ' Se ocorreu algum erro na DLL exibi-lo
    If Not bRetorno Then
        
        GoTo Finalizar
    
    End If

Finalizar:

    Screen.MousePointer = vbDefault
    
    If objMensagem.TipoMsg <> "" Then
    
        objMensagem.ExibeMensagem
        
    End If
    
    Set objMensagem = Nothing
    
    Exit Sub
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: MontarImpressaoOTS " & _
                            "da classe clsgenOC_SelecionarOTS"
    
    GoTo Finalizar

End Sub

Public Function RealocarCentroFuncional(ByVal pbChamouDaSelecionarOTS As Boolean, _
                                        Optional psTipoCorrecao As String) As Boolean

On Error GoTo ErrorHandler:

    Dim bRetorno                    As Boolean
    Dim objDadosCtrlCorrecao        As New clsGenDadosControleCorrecao
    Dim objFormAnotacoes            As frmGenObservacao
    Dim objDadosFase                As clsgenFase
    Dim objDadosUsuario             As clsgenUsuario
    Dim lSequenciaAcao              As Long
    Dim lIDFase                     As Long
    Dim lIDServico                  As Long
    Dim sTipoCorrecao               As String
    
    Set objMensagem = New clsgenMensagem
    'Verifica o Produto da OTS selecionada
    
    objFormSelecionarOTS.spdCircuitos.Col = 26
    lProduto_ID = Val(objFormSelecionarOTS.spdCircuitos.Text)
    
    If lProduto_ID = GbIDProdutoNetPhone Then
        
        objMensagem.TipoMsg = "A"
        objMensagem.Descricao = "Esta opção não está disponível para este serviço!"
        objMensagem.ExibeMensagem
        
        GoTo Finalizar
    
    End If
    
    'Coluna 28 - Armazena a Acao do servico selecionado
    objFormSelecionarOTS.spdCircuitos.Row = objFormSelecionarOTS.spdCircuitos.ActiveRow
    objFormSelecionarOTS.spdCircuitos.Col = 28
    
    If objFormSelecionarOTS.spdCircuitos.Text = "CAN" Or objFormSelecionarOTS.spdCircuitos.Text = "DES" Then
        
        objMensagem.TipoMsg = "A"
        objMensagem.Descricao = "Não é possível corrigir uma OTS de Cancelamento ou Desativação!"
        objMensagem.ExibeMensagem
        
        GoTo Finalizar
    
    End If
    
    Screen.MousePointer = vbArrowHourglass
    
    Set objFormAnotacoes = New frmGenObservacao

    objFormAnotacoes.Top = 1290

    If lProduto_ID = GbIDProdutoWID Then

        ' Dedicated
        objFormAnotacoes.Caption = "Worldcom Internet Dedicated"

    ElseIf lProduto_ID = GbIDProdutoVPN_RA Then

        ' VPN RA
        objFormAnotacoes.Caption = "VPN RA"

    ElseIf lProduto_ID = GbIDProdutoVPN_DIRECT Then

        ' VPN Direct
        objFormAnotacoes.Caption = "VPN Direct"


    ElseIf lProduto_ID = GbIDProdutoDiginet Then

        ' Diginet
        objFormAnotacoes.Caption = "Diginet"

    ElseIf lProduto_ID = GbIDProdutoEmbratelIPL Then

        ' IPL
        objFormAnotacoes.Caption = "Embratel IPL"

    ElseIf lProduto_ID = GbIDProdutoBLF Then

        ' BLF
        objFormAnotacoes.Caption = "Business Link Flex"
        
    ElseIf lProduto_ID = GbIDProdutoBIV Then
    
        ' IP VPN
        objFormAnotacoes.Caption = "IP VPN"
    
    ElseIf lProduto_ID = GbIDProdutoJAH Then

        ' JAH
        objFormAnotacoes.Caption = "JAH!"
        
    ElseIf lProduto_ID = GbIDProdutoBLD Then

        '#BLD
        objFormAnotacoes.Caption = "Business Link Direct"

    ElseIf (lProduto_ID = GbIDProdutoVOIP) Then

        '#VOIP
        objFormAnotacoes.Caption = "Internet Voice"

    End If

    'Coluna 15 - Armazena a Data de Inicio Real da fase
    objFormSelecionarOTS.spdCircuitos.Row = objFormSelecionarOTS.spdCircuitos.ActiveRow
    objFormSelecionarOTS.spdCircuitos.Col = 15
    
    If (objFormSelecionarOTS.spdCircuitos.Text <> "") And pbChamouDaSelecionarOTS Then
        
        objFormAnotacoes.OptRealocacao.Enabled = False
        objFormAnotacoes.OptCorrecao.Value = True
            
    End If

    objDadosCtrlCorrecao.InformaBanco frmgenMDI_SGAPlus.objBanco
    objDadosCtrlCorrecao.InformaMensagem objMensagem

    'Coluna 8 - Armazena o ID da FASE
    objFormSelecionarOTS.spdCircuitos.Row = objFormSelecionarOTS.spdCircuitos.ActiveRow
    objFormSelecionarOTS.spdCircuitos.Col = 8
    
    lIDFase = CLng(objFormSelecionarOTS.spdCircuitos.Text)

    'Coluna 9 - Armazena o ID do Servico
    objFormSelecionarOTS.spdCircuitos.Row = objFormSelecionarOTS.spdCircuitos.ActiveRow
    objFormSelecionarOTS.spdCircuitos.Col = 9
    
    lIDServico = CLng(objFormSelecionarOTS.spdCircuitos.Text)

    Set objDadosFase = New clsgenFase
    
    objDadosFase.InformaBanco frmgenMDI_SGAPlus.objBanco
    objDadosFase.InformaMensagem objMensagem
    
    lIDFase = objDadosFase.ObtemIDFaseCorrigir(lIDServico, sTipoCorrecao)
    
    'Obtem anotacao registrada para justificar a realocação do Centro Funcional da fase
    
    If objDadosCtrlCorrecao.RecuperarDados(lIDFase, sTipoCorrecao) Then

        objFormAnotacoes.txtTexto.Text = objDadosCtrlCorrecao.Justificativa
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
        
        If Not pbChamouDaSelecionarOTS Then
    
           objFormAnotacoes.txtTexto.Text = Anotacao
        
        Else
        
            objFormAnotacoes.LabelSolicitacao.Visible = False
            objFormAnotacoes.LblUsuarioSolicitacao.Visible = False
            objFormAnotacoes.LabelData.Visible = False
            objFormAnotacoes.LblDataSolicitacao.Visible = False
            
        End If
    
    End If
    
    objFormAnotacoes.OptDevolucao.Visible = False
    
    ' Informar o objeto controlador
    objFormAnotacoes.informa_OC Me

    ' Se a chamada for da tela se seleção de OTS, colocar o botão de gravação visível
    objFormSelecionarOTS.spdCircuitos.Row = objFormSelecionarOTS.spdCircuitos.ActiveRow
    objFormSelecionarOTS.spdCircuitos.Col = 3

    If pbChamouDaSelecionarOTS Then

        objFormAnotacoes.cmdGravar.Enabled = True
        objFormAnotacoes.cmdGravar.Visible = True

        objFormAnotacoes.lblCabecalho.Caption = "OTS: " & objFormSelecionarOTS.spdCircuitos.Text

    Else

        objFormAnotacoes.cmdGravar.Enabled = False
        objFormAnotacoes.cmdGravar.Visible = False
        objFormAnotacoes.txtTexto.Locked = True
        objFormAnotacoes.OptCorrecao.Enabled = False
        objFormAnotacoes.OptRealocacao.Enabled = False
        
        objFormAnotacoes.lblCabecalho.Caption = "OTS: " & objFormSelecionarOTS.spdCircuitos.Text
    
    End If

    If sTipoCorrecao = "R" Then
        
        objFormAnotacoes.OptRealocacao.Value = True
        
    ElseIf sTipoCorrecao = "C" Then
    
        objFormAnotacoes.OptCorrecao.Value = True
    
    End If
             
    If lIDFase = 0 Then
        
        lIDFase = objDadosFase.ObtemIDFaseCorrigir(lIDServico, "D")
    
    End If
    
    'Caso tenha obs de devolucao, recuperar e carrega na tela
    If objDadosCtrlCorrecao.RecuperarDados(lIDFase, "D") And pbChamouDaSelecionarOTS Then
       
       objFormAnotacoes.TxtObsCorrecao.Text = objDadosCtrlCorrecao.Justificativa
       objFormAnotacoes.framePrincipal.Height = 1245
       objFormAnotacoes.framePrincipal.Width = 7290
       objFormAnotacoes.framePrincipal.Top = 2475
       objFormAnotacoes.FrameObsCorrecao.Visible = True
       
    Else
    
       objFormAnotacoes.framePrincipal.Height = 2385
       objFormAnotacoes.framePrincipal.Width = 7275
       objFormAnotacoes.framePrincipal.Top = 1260
       objFormAnotacoes.txtTexto.Height = 2055
       objFormAnotacoes.txtTexto.Width = 7035
       
       objFormAnotacoes.FrameObsCorrecao.Visible = False

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

Public Property Get Familia_Produto() As String

    Familia_Produto = sFamilia_Produto

End Property

Public Property Let Familia_Produto(ByVal sNewValue As String)

    sFamilia_Produto = sNewValue
    
End Property

Private Sub Class_Terminate()

    Call LiberaMemoria
    
End Sub

Public Sub LiberaMemoria()

    Set objOcAlocarCF = Nothing
    Set objOC_AcessoEmbratel = Nothing
    Set objOC_AcessoCliente = Nothing
    Set objocConfiguracao = Nothing
    Set objocTratarCentroLocal = Nothing
    Set objocCoordCPE = Nothing
    Set objocPreAtivacao = Nothing
    Set objocAtivacao = Nothing
    Set objocTratarBGP = Nothing
    Set objMensagem = Nothing
    Set objCentroFuncional = Nothing
    Set objServico = Nothing
    Set objCliente = Nothing
    Set objListaOTS = Nothing
    Set objFase = Nothing
    Set objFormSelecionarOTS = Nothing
    Set objOC_Designacao = Nothing
    Set objOC_TimeSlotDD = Nothing
    
    'JAH
    Set objocConfiguracaoTel = Nothing

End Sub
