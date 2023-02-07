VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsgenOC_OTSCorrecao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Declara as classes utilizadas
Public objOCSelecionarOTS   As clsgenOC_SelecionarOTSCorrecao

Private objAnalisarOE       As clsblfOrderEntry
Private objServico          As clsGenServico
Private objCircuito         As clsgenCircuito
Private objCliente          As clsgenCliente
Private objAcao             As clsGenAcao
Public objMensagem          As clsgenMensagem
Private objCPE              As clsblf_CPE
Private objFase             As clsgenFase

' Declara os formulários usados nesta classe
Private objFormOTSCorrecao         As frmgenOTSCorrecao
Private objFormExibirDesignacaoOTS As Form

' Outras variáveis
Private sTemCPE             As String
Private sDesignacaoCircuito As String
Private sOTS                As String
Private bTemCPE             As Boolean
Private sSiglaProduto       As String
Private bRedundante         As Boolean

'IP VPN
Private lProduto_ID         As Long
Private lGrupo_ID           As Long
Private sCodigoProduto      As String
Private lFase_Id            As Long

'''
Public Property Get CodigoProduto() As String
    CodigoProduto = sCodigoProduto
End Property

Public Property Let CodigoProduto(ByVal vNewValue As String)
    sCodigoProduto = vNewValue
End Property
'''

Public Property Get Fase_ID() As Long
    Fase_ID = lFase_Id
End Property

Public Property Let Fase_ID(ByVal vNewValue As Long)
    lFase_Id = vNewValue
End Property

'''
Public Property Get Produto_ID() As Long
    Produto_ID = lProduto_ID
End Property

Public Property Let Produto_ID(ByVal vNewValue As Long)
    lProduto_ID = vNewValue
End Property

Public Property Get Grupo_ID() As Long
    Grupo_ID = lGrupo_ID
End Property

Public Property Let Grupo_ID(ByVal vNewValue As Long)
    lGrupo_ID = vNewValue
End Property

Public Function TelaOTSCorrecao(ByVal plID_SERVICO As Long) As Boolean

On Error GoTo ErrorHandler:
    Dim bRetorno               As Boolean
    Dim pbExibeExcluido        As Boolean
    Dim objDadosProduto        As New clsGenProduto
    Dim objDadosFase           As clsgenFase
    Dim lFaseIDAux             As Long
    
    '=================================================================
    'CALLBAT - 18/09/2014 - BSoD Unificado
    'Correção efetuada para atendimento da Solicitação da Analista de
    'Processo Aline Silva
    '=================================================================
    Dim lContingencia_Acesso   As Long
    Dim objContingencia_Acesso As New clsGenAcesso
    '=================================================================
    '=================================================================
    
    '================================================================
    'Callbat - 07/04/2015
    'Chamado CH-74999UGL
    '================================================================
    Dim objLogSGAPCLA           As clsBLD_LogSGAPCLA
    Set objLogSGAPCLA = New clsBLD_LogSGAPCLA
    
    objLogSGAPCLA.InformaBanco frmgenMDI_SGAPlus.objBanco
    objLogSGAPCLA.InformaMensagem objMensagem
    '================================================================
    '=================================================================
    '''===FABIMENDONCA - 04/2021 - SO-SOLEN-28213JNU - Acesso 4G - Rede Wireless
    Dim objRespostaWebIP       As New clsBlf_RespostaWebIP
    Dim objDadosTecnicos        As clsbldDadosTecnicos
    
    
    '================================================================
    
    
    
    Screen.MousePointer = vbArrowHourglass

    TelaOTSCorrecao = False

    ' Instância o formulário de correção de OTS
    Set objFormOTSCorrecao = New frmgenOTSCorrecao
    
    ' Informar o objeto controlador
    objFormOTSCorrecao.informa_OC Me
    
    '
    ' Instância as classes necessárias para geração da ots
    Set objMensagem = New clsgenMensagem
    Set objServico = New clsGenServico
    Set objCircuito = New clsgenCircuito
    Set objCliente = New clsgenCliente
    Set objAcao = New clsGenAcao
    Set objFase = New clsgenFase
    
    Set objDadosProduto = New clsGenProduto
    
    '=================================================================
    'CALLBAT - 18/09/2014 - BSoD Unificado
    'Correção efetuada para atendimento da Solicitação da Analista de
    'Processo Aline Silva
    '=================================================================
    Set objContingencia_Acesso = New clsGenAcesso
    Set objDadosFase = New clsgenFase
    '=================================================================
    '=================================================================
    
    
    '
    ' Informa a classe e banco para as classes
    '
    objServico.InformaBanco frmgenMDI_SGAPlus.objBanco
    objCircuito.InformaBanco frmgenMDI_SGAPlus.objBanco
    objCliente.InformaBanco frmgenMDI_SGAPlus.objBanco
    objAcao.InformaBanco frmgenMDI_SGAPlus.objBanco
    objFase.InformaBanco frmgenMDI_SGAPlus.objBanco
    objDadosProduto.InformaBanco frmgenMDI_SGAPlus.objBanco
    
    '=================================================================
    'CALLBAT - 18/09/2014 - BSoD Unificado
    'Correção efetuada para atendimento da Solicitação da Analista de
    'Processo Aline Silva
    '=================================================================
    objContingencia_Acesso.InformaBanco frmgenMDI_SGAPlus.objBanco
    objDadosFase.InformaBanco frmgenMDI_SGAPlus.objBanco
    '=================================================================
    '=================================================================
    
    
    '
    ' Informa a classe de mensagem para as classes
    '
    objServico.InformaMensagem objMensagem
    objCircuito.InformaMensagem objMensagem
    objCliente.InformaMensagem objMensagem
    objAcao.InformaMensagem objMensagem
    objFase.InformaMensagem objMensagem
    
    objDadosProduto.InformaMensagem objMensagem
    
    '=================================================================
    'CALLBAT - 18/09/2014 - BSoD Unificado
    'Correção efetuada para atendimento da Solicitação da Analista de
    'Processo Aline Silva
    '=================================================================
    objContingencia_Acesso.InformaMensagem objMensagem
    objDadosFase.InformaMensagem objMensagem
    '=================================================================
    '=================================================================
    
    
    '-- VAGNERR 19/09
    ' Recuperar os dados do servico
    bRetorno = objServico.RecuperarDados(plID_SERVICO)
    
    If Not bRetorno Then
    
        GoTo Finalizar
        
    End If
    '------

        '-- VAGNERR 19/09
    'Obtem dados do Produto
    'bRetorno = objDadosProduto.RecuperarDados(Produto_ID)
    bRetorno = objDadosProduto.RecuperarDados(objServico.ID_Produto)
        '------

    'Trata retorno do método
    If (Not bRetorno) Then
    
        GoTo Finalizar
        
    End If
    
    objFormOTSCorrecao.lblProduto.Caption = objDadosProduto.Descricao
    objFormOTSCorrecao.Caption = objDadosProduto.Sigla & " - Geração de OTS de correção"
    
    ' Recuperar os dados do servico
    bRetorno = objServico.RecuperarDados(plID_SERVICO)
    
    If Not bRetorno Then
    
        GoTo Finalizar
        
    End If
    
    ' Recuperar os dados do circuito
    bRetorno = objCircuito.RecuperarDados(plID_SERVICO)
    
    If Not bRetorno Then
    
        GoTo Finalizar
        
    End If
    
    ' Recuperar os dados do cliente
    bRetorno = objCliente.RecuperarDados(objCircuito.ContaCliente)
    
    If Not bRetorno Then
        GoTo Finalizar
    End If
    
    ' Recuperar os dados da ação
    bRetorno = objAcao.RecuperarDados(objServico.Id_Acao)
    
    If Not bRetorno Then
    
        GoTo Finalizar
        
    End If
    
    
    '=========================================================================
    'Callbat - 07/04/2015
    'Chamado CH-74999UGL
    '----------------------------------------------
    'Verifica se, já existe esse circuito na interface junto ao CLA,
    'caso não exista, essa correção, obrigatóriamente, deve ser encaminhada
    'para o CLA
    '=========================================================================
    
    'If objServico.Acao = 7 Or objServico.Acao = 14 Then
    
    'AGRC - 04/11/2016 - SO-MKTIP-18424CTY
    If objCircuito.EBkp3G = True Then
        objFormOTSCorrecao.OptMotivoCorrecao.Item(0).Enabled = False
        objFormOTSCorrecao.OptMotivoCorrecao.Item(1).Enabled = True
        objFormOTSCorrecao.OptMotivoCorrecao.Item(1).Value = True
    
    'TBASTOS - FLEX VPN não exibe mensagem, mas obriga a ir na fase Tratar Acesso.
    ElseIf objServico.ID_Produto = GbIDProdutoFLEXVPN Then
    
        objFormOTSCorrecao.OptMotivoCorrecao.Item(1).Enabled = False
        objFormOTSCorrecao.OptMotivoCorrecao.Item(0).Enabled = True
        objFormOTSCorrecao.OptMotivoCorrecao.Item(0).Value = True
    
    Else
        bRetorno = objLogSGAPCLA.VerificaExisteCCTOCla(objServico.ID_CIRCUITO)
        
        If bRetorno = False Then
            objMensagem.TipoMsg = "A"
            objMensagem.Descricao = "ATENÇÃO: Essa OTS de correção deve ser de Tipo de Acesso pois" & Chr(13) & _
                                    "                    não existe esse circuito na interface junto ao CLA." & Chr(13)
            
            objMensagem.ExibeMensagem
            
            objFormOTSCorrecao.OptMotivoCorrecao.Item(0).Enabled = True
            objFormOTSCorrecao.OptMotivoCorrecao.Item(1).Enabled = False
            objFormOTSCorrecao.OptMotivoCorrecao.Item(0).Value = True
            
        End If
    End If
    'End If
    '================================================================
    '================================================================
    
    
    '=================================================================
    'CALLBAT - 18/09/2014 - BSoD Unificado
    'Correção efetuada para atendimento da Solicitação da Analista de
    'Processo Aline Silva
    '=================================================================
    lFase_Id = objDadosFase.ObtemIDFase(objServico.ID_Servico, GbCodigoFaseAcesso)
    
    
    '=================================================================
    'CALLBAT - Chamado CH-31271JVU
    '=================================================================
    'bRetorno = objContingencia_Acesso.RecuperarDados("", objServico.ID_SERVICO, lFase_Id, objMensagem)
    'If Not bRetorno Then
    '    GoTo Finalizar
    'End If
    
    Call objContingencia_Acesso.RecuperarDados("", objServico.ID_Servico, lFase_Id, objMensagem)
    '=================================================================
    '=================================================================
    
    If objContingencia_Acesso.PropridedadeAcesso = "T" Then
    
        Dim objPropriedadeTerceiro As New clsGenPropriedadeTerceiros
        
        Set objPropriedadeTerceiro = New clsGenPropriedadeTerceiros
        objPropriedadeTerceiro.InformaBanco frmgenMDI_SGAPlus.objBanco
        objPropriedadeTerceiro.InformaMensagem objMensagem
        
        
        bRetorno = objPropriedadeTerceiro.RecuperarDados(objContingencia_Acesso.iD, objMensagem)
        
        If Not bRetorno Then
            GoTo Finalizar
        End If
        
        If objPropriedadeTerceiro.TipoAcesso = "N" Or objPropriedadeTerceiro.TipoAcesso = "V" Then
        
                '=================================================================
                'CALBAT - 19/09/2014
                '=================================================================
                objMensagem.TipoMsg = "A"
                objMensagem.Descricao = "ATENÇÃO: Estão bloqueadas novas alterações para o tipo de acesso " & IIf(objPropriedadeTerceiro.TipoAcesso = "N", "HFC BsoD Net", "HFC BsoD Via") & ". " & Chr(13) & _
                                        "                  Assim, é necessário realizar a migração do acesso para HFC BSoD." & Chr(13) & _
                                        "                  Em caso de dúvidas, contatar a equipe do SOE."
                
                objMensagem.ExibeMensagem
                '=================================================================
                '=================================================================
                
                objFormOTSCorrecao.OptMotivoCorrecao.Item(0).Enabled = True
                objFormOTSCorrecao.OptMotivoCorrecao.Item(1).Enabled = False
                
                objFormOTSCorrecao.OptMotivoCorrecao.Item(0).Value = True
        
        
        End If
        
        Set objPropriedadeTerceiro = Nothing
    
    
    End If
    
    
    'Callbat - Projeto Raio-X
    'objFormOTSCorrecao.tabDadosOE.TabVisible(3) = False
    
    '=================================================================
    '=================================================================
    
    ' Preenche os campos do formulário
    '
    ' Campos do cabecalho
    Call PreencherCabecalhoTelaOTSCorrecao
    
    ' Atribui valores especicíficos para cada produto
    
    Select Case objServico.ID_Produto
    ' Atribui valores especicíficos para cada produto

           Case GbIDProdutoWID, GbIDProdutoVPN_RA, GbIDProdutoVPN_DIRECT
           
                ' Campos do CPE
                bTemCPE = False
                bRetorno = PreencheCamposCPE_WID(objMensagem, _
                                                 objFormOTSCorrecao.UsrBlfCPE1, _
                                                 objServico.ID_Contrato, _
                                                 bTemCPE)
                If Not bRetorno Then
                    GoTo Finalizar
                End If
                
                If bTemCPE Then
                
                  'TBASTOS CH-16265SQF
                    'sTemCPE = "S"
                    sTemCPE = "N"
                
                End If
                
                ' Campos dados técnicos
'VCB 08/08/2019 --- Correção causa raíz WID
'                bRetorno = PreencheCamposDadosTecnicos_WID()
                bRetorno = PreencheCamposDadosTecnicos_WID(objServico.ID_Servico)
'-------
                If Not bRetorno Then
                    GoTo Finalizar
                End If
           
                ' Campos dos dados cadastrais
                bRetorno = PreencheCamposDadosCadastrais_WID(objMensagem, _
                                                            objFormOTSCorrecao.UsrBLFDadosCadastrais1, _
                                                            objServico.ID_Contrato, _
                                                            objServico.NR_LOCAL_INSTA, _
                                                            "G")
                
                If Not bRetorno Then
                
                    GoTo Finalizar
                
                End If
           
           Case GbIDProdutoBLF
                
                ' Campos do CPE
                sTemCPE = "N"
                ' Campos do CPE
                bRetorno = PreencheCamposCPE_BLF(objMensagem, _
                                                 objFormOTSCorrecao.UsrBlfCPE1, _
                                                 objServico.ID_Contrato, _
                                                 objServico.NR_LOCAL_INSTA)
                If Not bRetorno Then
                    GoTo Finalizar
                End If
                
                If objFormOTSCorrecao.UsrBlfCPE1.PropriedadeCPE = "EMBRATEL" Then
                
           'TBASTOS CH-16265SQF
                    'sTemCPE = "S"
                    sTemCPE = "N"
                
                End If
                
                ' Campos dados técnicos
                bRetorno = PreencheCamposDadosTecnicos_BLF()
                
                If Not bRetorno Then
                    GoTo Finalizar
                End If
            
                ' Campos dos dados cadastrais
                bRetorno = PreencheCamposDadosCadastrais_BLF(objMensagem, _
                                                            objFormOTSCorrecao.UsrBLFDadosCadastrais1, _
                                                            objServico.ID_Contrato, _
                                                            objServico.NR_LOCAL_INSTA, _
                                                             "T")
                
                If Not bRetorno Then
                
                    GoTo Finalizar
                
                End If
                
           Case GbIDProdutoBIV

                CodigoProduto = GbCodigo_Produto_IPVPN

                objFormOTSCorrecao.frameWD1.Visible = False
                objFormOTSCorrecao.FramBLF.Visible = False
                objFormOTSCorrecao.FramInfTecnicasSite.Visible = True
                objFormOTSCorrecao.Frame19.Caption = "Tipo Item / Circuito"
           
                ' Campos do CPE
                If (objServico.Possui_CPE = "N") Then

                    sTemCPE = "N"
                    objFormOTSCorrecao.LbCpe_Site.Caption = "N"
                    objFormOTSCorrecao.tabDadosOE.TabVisible(2) = False
                    
                Else
                
                    objFormOTSCorrecao.LbCpe_Site.Caption = "S"
                    
                    ' Campos do CPE
                    bRetorno = PreencheCamposCPE_BLF(objMensagem, _
                                                     objFormOTSCorrecao.UsrBlfCPE1, _
                                                     objServico.ID_Contrato, _
                                                     objServico.NR_LOCAL_INSTA, _
                                                     CodigoProduto)
                    If Not bRetorno Then
                        GoTo Finalizar
                    End If
                    
                    If objFormOTSCorrecao.UsrBlfCPE1.PropriedadeCPE = "EMBRATEL" Then
           'TBASTOS CH-16265SQF
                    'sTemCPE = "S"
                    sTemCPE = "N"
                    End If
                    
                End If
                
                ' Campos dos dados cadastrais
                bRetorno = PreencheCamposDadosCadastrais_BLF(objMensagem, _
                                                            objFormOTSCorrecao.UsrBLFDadosCadastrais1, _
                                                            objServico.ID_Contrato, _
                                                            objServico.NR_LOCAL_INSTA, _
                                                            "T")
                If Not bRetorno Then
                    GoTo Finalizar
                End If
    
                ' Campos dados técnicos
                'Trata se exibe ao usuário os sites enviados pelo SOE
                'para exclusão(Desativação/Cancelamento)
                
                If (objAcao.Sigla = "DES" Or _
                        objAcao.Descricao = "CAN") Then
                    
                    pbExibeExcluido = True
                    
                Else
                
                    pbExibeExcluido = False
                    
                End If
                
                bRetorno = PreencheCamposDadosTecnicos_IVN(objServico.TipoItem, _
                    pbExibeExcluido)
                
                If (Not bRetorno) Then
                
                    GoTo Finalizar
                    
                End If
           '
           Case GbIDProdutoJAH

                CodigoProduto = GbCodigo_Produto_JAH

                objFormOTSCorrecao.frameWD1.Visible = False
                objFormOTSCorrecao.FramBLF.Visible = False
                objFormOTSCorrecao.FramInfTecnicasSite.Visible = False
                objFormOTSCorrecao.Frame19.Caption = "Tipo Item / Circuito"
                objFormOTSCorrecao.FramJAH.Visible = True
           
                sTemCPE = "N"
                objFormOTSCorrecao.tabDadosOE.TabVisible(2) = False
                objFormOTSCorrecao.tabDadosOE.TabVisible(2) = False
                
                    
                ' Campos dos dados cadastrais
                bRetorno = PreencheCamposDadosCadastrais_BLF(objMensagem, _
                                                            objFormOTSCorrecao.UsrBLFDadosCadastrais1, _
                                                            objServico.ID_Contrato, _
                                                            objServico.NR_LOCAL_INSTA, _
                                                            "T", _
                                                            Produto_ID)
                If (Not bRetorno) Then
                
                    GoTo Finalizar
                    
                End If
    
                ' Campos dados técnicos
                bRetorno = PreencheCamposDadosTecnicos_JAH()
                
                If (Not bRetorno) Then
                
                    GoTo Finalizar
                    
                End If
           '
           Case GbIDProdutoBLD, GbIDProdutoVOIP, GbIDProdutoRUDMPLS, GbIDProdutoVBL, GbIDProdutoVipUnico
           
                CodigoProduto = "IP"

                'objFormOTSCorrecao.frameWD1.Visible = False
                'objFormOTSCorrecao.FramBLF.Visible = False
                'objFormOTSCorrecao.FramInfTecnicasSite.Visible = False
                'objFormOTSCorrecao.Frame19.Caption = "Tipo Item / Circuito"
                'objFormOTSCorrecao.FramJAH.Visible = True
           
                sTemCPE = "N"
                objFormOTSCorrecao.tabDadosOE.TabVisible(2) = False
                
                ' Campos dos dados cadastrais
                bRetorno = PreencheCamposDadosCadastrais_BLD(objMensagem, _
                                                            objFormOTSCorrecao.UsrBLFDadosCadastrais1, _
                                                            objServico.NumOE, _
                                                            objServico.AnoOE, _
                                                            objServico.NR_LOCAL_INSTA, _
                                                            "T")
                If (Not bRetorno) Then
                
                    GoTo Finalizar
                    
                End If

                ' Campos dados técnicos
                bRetorno = PreencheCamposDadosTecnicos_BLD()
                
                If (Not bRetorno) Then
                
                    GoTo Finalizar
                    
                End If
                
                '=================================================================
                '''===FABIMENDONCA - 04/2021 - SO-SOLEN-28213JNU - Acesso 4G - Rede Wireless
                Set objDadosTecnicos = New clsbldDadosTecnicos
                
                objDadosTecnicos.InformaBanco frmgenMDI_SGAPlus.objBanco
                objDadosTecnicos.InformaMensagem objMensagem
                
                bRetorno = objDadosTecnicos.RecuperarDados(objServico.NumOE, objServico.AnoOE)
                If (Not bRetorno) Then
                
                    GoTo Finalizar
                    
                End If
                If objDadosTecnicos.WIRELESS = "S" Then
                    objFormOTSCorrecao.OptMotivoCorrecao(0).Value = False
                    objFormOTSCorrecao.OptMotivoCorrecao(0).Enabled = False
                    
                    objFormOTSCorrecao.OptMotivoCorrecao(1).Value = True
                    objFormOTSCorrecao.OptMotivoCorrecao(1).Enabled = True
                End If
                '=================================================================
                Set objDadosTecnicos = Nothing
    End Select
    
    'Trata se o botão Gerar OTS deve ser desabilitado
    If (objServico.Correcao <> "C") Then

        objFormOTSCorrecao.cmdGerarOTS.Enabled = False
        
    End If
    
    ' A opção de Alteracao de Configuração só poderá ser habilitada para seleção se a fase de acesso da OTS anterior estiver finalizada
    ' Conforme solicitado, as duas opções estarão sempre habilitadas 18/07/2011 vcb
'    Set objDadosFase = New clsgenFase
'
'    objDadosFase.InformaBanco frmgenMDI_SGAPlus.objBanco
'    objDadosFase.InformaMensagem objMensagem
'
'    lFaseIDAux = objDadosFase.ObtemIDFase(objServico.ID_SERVICO, GbCodigoFaseAcesso)
'
'    If objDadosFase.ObtemDataFimReal(lFaseIDAux) = "" Then
'
'        objFormOTSCorrecao.OptMotivoCorrecao(0).Value = True
'        objFormOTSCorrecao.OptMotivoCorrecao(1).Enabled = False
'
'    End If
'
'
'    If (objMensagem.TipoMsg <> "") Then
'
'       GoTo ErrorHandler
'
'    End If
    
    
    Screen.MousePointer = vbDefault
    
    objFormOTSCorrecao.Show vbModal

    TelaOTSCorrecao = True
    
Finalizar:
    Screen.MousePointer = vbDefault
    
    If objMensagem.TipoMsg <> "" Then

        objMensagem.ExibeMensagem
        objMensagem.TipoMsg = ""

    End If
    
    
    Set objDadosFase = Nothing
    Set objContingencia_Acesso = Nothing
    
    On Error GoTo 0

    Exit Function

ErrorHandler:

    If objMensagem.TipoMsg = "" And Err.Number <> 0 Then

        objMensagem.TipoMsg = "E"
        objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                                "Houve um erro de execução no método: TelaOTSCorrecao " & _
                                "da classe clsgenOC_OTSCorrecao"
    

    End If
    
    GoTo Finalizar

End Function
Private Function TelaExibirDesignacaoOTS() As Boolean

On Error GoTo ErrorHandler:

    Dim bRetorno As Boolean
    
    TelaExibirDesignacaoOTS = False

    If objServico.ID_Produto = GbIDProdutoWID Then
    
        ' Instância o formulário de que exibe os números de designação e OTS
        Set objFormExibirDesignacaoOTS = New frmwidExibeDesignacaoOTS
    
    'IP VPN
    'ElseIf objServico.ID_Produto = GbIDProdutoBLF Or _
            objServico.ID_Produto = GbIDProdutoBIV Or _
            objServico.ID_Produto = GbIDProdutoJAH Or _
            objServico.ID_Produto = GbIDProdutoBLD Or _
            objServico.ID_Produto = GbIDProdutoVOIP Then
            
    ElseIf objServico.ID_Produto = GbIDProdutoBLF Or _
            objServico.ID_Produto = GbIDProdutoBIV Or _
            objServico.ID_Produto = GbIDProdutoJAH Or _
            objServico.ID_Grupo = GbIDGrupoBLD Or _
            objServico.ID_Grupo = GbIDGrupoVLI Or _
            objServico.ID_Produto = GbIDProdutoVOIP Then

    
        ' Instância o formulário de que exibe os números de designação e OTS
        Set objFormExibirDesignacaoOTS = New frmblfExibeDesignacaoOTS
        
    End If
    
    Screen.MousePointer = vbDefault

    ' Preenche os campos do formulário
    Call PreencheCamposFormExibirDesignacaoOTS

    objFormExibirDesignacaoOTS.Show vbModal
    
    objFormOTSCorrecao.cmdGerarOTS.Enabled = False

    TelaExibirDesignacaoOTS = True

Finalizar:

    Screen.MousePointer = vbDefault

    If objMensagem.TipoMsg <> "" Then

        objMensagem.ExibeMensagem
        objMensagem.TipoMsg = ""

    End If
    
    objMensagem.TipoMsg = ""
    
    On Error GoTo 0

    Exit Function

ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: TelaExibirDesignacaoOTS " & _
                            "da classe clsgenOC_OTSCorrecao"

    GoTo Finalizar

End Function

Private Sub Class_Terminate()

    Set objServico = Nothing
    Set objCircuito = Nothing
    Set objCliente = Nothing
    Set objAcao = Nothing
    Set objMensagem = Nothing
    Set objFormOTSCorrecao = Nothing
    Set objCPE = Nothing
    Set objFase = Nothing
    
End Sub

Private Function PreencheCamposFormExibirDesignacaoOTS() As Boolean
    Dim objCentroFuncional As clsgenCentroFuncional
    Dim bRetorno As Boolean
    
    'IP VPN
    
    Dim objDadosProduto As clsGenProduto
    
    Set objDadosProduto = New clsGenProduto
    
    objDadosProduto.InformaBanco frmgenMDI_SGAPlus.objBanco
    objDadosProduto.InformaMensagem objMensagem
'VAGNERR 17/09
    'Obtem dados do Produto
    'bRetorno = objDadosProduto.RecuperarDados(Produto_ID)
    bRetorno = objDadosProduto.RecuperarDados(objServico.ID_Produto)
'----
    objFormExibirDesignacaoOTS.Caption = objDadosProduto.Sigla & " - OTS"


    If Produto_ID = GbIDProdutoBIV Or Produto_ID = GbIDProdutoJAH Then
        
        objFormExibirDesignacaoOTS.lblOTSRedundante2.Visible = False
        objFormExibirDesignacaoOTS.lblOTSRedundante.Visible = False
        
        objFormExibirDesignacaoOTS.lblDesigRedundante2.Visible = False
        objFormExibirDesignacaoOTS.lblDesigRedundante.Visible = False
    
    End If

    'IP VPN
    
    PreencheCamposFormExibirDesignacaoOTS = False
    
    ' Instância a classe de centro funcionais
    Set objCentroFuncional = New clsgenCentroFuncional
    
    ' Informa os objetos de banco e de mensagem
    objCentroFuncional.InformaBanco frmgenMDI_SGAPlus.objBanco
    objCentroFuncional.InformaMensagem objMensagem
    
    ' Recupera os dados do centro funcional
    bRetorno = objCentroFuncional.RecuperarDados(objServico.ID_CENTRO_FUNCIONAL)
    
    If Not bRetorno Then
    
        GoTo Finalizar
    
    End If
    
    If objServico.ID_Produto = GbIDProdutoWID Then
    
        objFormExibirDesignacaoOTS.lblUBR = objCircuito.UBR
    
    End If
    
    If Not bRedundante Then
    
        ' Atribui os valores as propriedades, para os objetos de tela
        'objFormExibirDesignacaoOTS.lstOTS = objCentroFuncional.Codigo_Centro_Funcional & _
                                                     " - " & _
                                                     sSiglaProduto & _
                                                     " - " & _
                                                     sOTS
                                                     
        objFormExibirDesignacaoOTS.lstOTS.AddItem (objCentroFuncional.Codigo_Centro_Funcional & _
                                                     " - " & _
                                                     sSiglaProduto & _
                                                     " - " & _
                                                     sOTS)
                                                     
                                                     
                                                     
        objFormExibirDesignacaoOTS.lblOTSRedundante = ""
        
        'objFormExibirDesignacaoOTS.lstDesignacao = sDesignacaoCircuito
        
        objFormExibirDesignacaoOTS.lstDesignacao.AddItem (sDesignacaoCircuito)
        
        objFormExibirDesignacaoOTS.lblDesigRedundante = ""
                                                     
    Else
    
        objFormExibirDesignacaoOTS.lstOTS.AddItem (objCentroFuncional.Codigo_Centro_Funcional & _
                                                     " - " & _
                                                     sSiglaProduto & _
                                                     " - " & _
                                                     sOTS)
                                                     
        objFormExibirDesignacaoOTS.lblOTSPrincipal = ""
        
        'objFormExibirDesignacaoOTS.lblDesigRedundante = sDesignacaoCircuito
        objFormExibirDesignacaoOTS.lstDesignacao.AddItem (sDesignacaoCircuito)
        
        objFormExibirDesignacaoOTS.lblDesigPrincipal = ""
                                                 
    End If
    
Finalizar:

    Screen.MousePointer = vbDefault

    If objMensagem.TipoMsg <> "" Then

        objMensagem.ExibeMensagem
        objMensagem.TipoMsg = ""

    End If
    
    Set objCentroFuncional = Nothing
    Set objDadosProduto = Nothing
    
    On Error GoTo 0

    Exit Function

ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: PreencheCamposFormExibirDesignacaoOTS " & _
                            "da classe clsgenOC_OTSCorrecao"

    GoTo Finalizar

   
End Function

Public Sub SaidaTelaOTSCorrecao()
    
    Dim bRetorno As Boolean

    Unload objFormOTSCorrecao
    
    objOCSelecionarOTS.ExibirListaOTS
    
End Sub

Private Sub PreencherCabecalhoTelaOTSCorrecao()

    ' Preenche os dados do cabecalho da tela de analisar circuito
    '
    ' Ação
    objFormOTSCorrecao.lblAcao.Caption = objAcao.Descricao
    ' Tipo do Circuito
    objFormOTSCorrecao.lblTipoCcto.Caption = objCircuito.TipoCircuitoDescricao
    ' Nome do cliente
    objFormOTSCorrecao.lblNomeCli.Caption = objCliente.Nome_Razao_Social_Cliente
    ' Nome Fantasia
    objFormOTSCorrecao.lblNomeFantasia.Caption = objCliente.Nome_Guerra_Clientes
    ' Conta do cliente
    objFormOTSCorrecao.lblConta.Caption = objCliente.Conta_Corrente_Cliente

End Sub

Private Function PreencheCamposDadosTecnicos_BLF() As Boolean

On Error GoTo ErrorHandler:

    Dim bRetorno As Boolean
    Dim objDadosTecnicos As clsblfDadosTecnicos
    
    PreencheCamposDadosTecnicos_BLF = False
    
    ' Classe de dados técnicos
    Set objDadosTecnicos = New clsblfDadosTecnicos
    
    ' Informar o objeto do banco para as classes
    '
    ' Classe de dados técnicos
    objDadosTecnicos.InformaBanco frmgenMDI_SGAPlus.objBanco
    
    ' Informar o objeto de mensagem para as classes
    '
    ' Classe de dados cadastrais
    objDadosTecnicos.InformaMensagem objMensagem
    
    'esta ocorrendo erro na linha abaixo: vcb 12/07
    'objDadosTecnicos.Servico = objOCSelecionarOTS
    
    '=============================================================================
    'CALLBAT - 04/02/2011
    '------------------------------------------------------------------------------------------------------------
    'Otimização BLF
    '=============================================================================
    objDadosTecnicos.Servico = objServico.ID_Servico
    '=============================================================================
    '=============================================================================
    ' Executar o método para recuperar os dados cadastrais
    bRetorno = objDadosTecnicos.RecuperarDados(objServico.ID_Contrato, _
                                               objServico.NR_LOCAL_INSTA)
    
    ' Se ocorreu algum erro na DLL exibi-lo
    If Not bRetorno Then
        
        GoTo Finalizar
    
    End If
    
    ' Atualiza os campos
    objFormOTSCorrecao.lblDesigCcto.Caption = objDadosTecnicos.DesignacaoCircuito
    objFormOTSCorrecao.lblOTSassociada.Caption = objDadosTecnicos.Ots
    objFormOTSCorrecao.lblConexInternet.Caption = objDadosTecnicos.ConexaoInternet
    objFormOTSCorrecao.lblProtocolo.Caption = objDadosTecnicos.Protocolo
    objFormOTSCorrecao.lblVelocidadeCircuito.Caption = objDadosTecnicos.Velocidade_WebIP
    objFormOTSCorrecao.lblCIRSIR.Caption = objDadosTecnicos.CIR_SIR
    objFormOTSCorrecao.lblBGP.Caption = objDadosTecnicos.TipoBGP
    objFormOTSCorrecao.lblPontoCentral.Caption = objDadosTecnicos.PontoCentral
    objFormOTSCorrecao.lblCPE.Caption = sTemCPE
    objFormOTSCorrecao.lblSMTP.Caption = objDadosTecnicos.SMTP
    objFormOTSCorrecao.lblSLA.Caption = objDadosTecnicos.SLA
    
    PreencheCamposDadosTecnicos_BLF = True
    
Finalizar:

    Set objDadosTecnicos = Nothing
    
    On Error GoTo 0
    
    Exit Function
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: PreencheCamposDadosTecnicos_BLF " & _
                            "da classe clsgenOC_OTSCorrecao"
    
    GoTo Finalizar

End Function



Private Function PreencheCamposDadosTecnicos_BLD() As Boolean
    '''
On Error GoTo ErrorHandler:

    Dim bRetorno As Boolean
    Dim objDadosTecnicos As clsbldDadosTecnicos
    
    PreencheCamposDadosTecnicos_BLD = False
    
    ' Classe de dados técnicos
    Set objDadosTecnicos = New clsbldDadosTecnicos
    
    ' Informar o objeto do banco para as classes
    '
    ' Classe de dados técnicos
    objDadosTecnicos.InformaBanco frmgenMDI_SGAPlus.objBanco
    
    ' Informar o objeto de mensagem para as classes
    '
    ' Classe de dados cadastrais
    objDadosTecnicos.InformaMensagem objMensagem
    
    ' Executar o método para recuperar os dados cadastrais
    '======================================================================
    'CALBAT - 26/02/2008                                                                                     =
    'Acerto identificado para acerto efetuado no BO 451483                                   =
    '======================================================================
    '======================================================================
    'bRetorno = objDadosTecnicos.RecuperarDados(objServico.NumOE, _
                                        objServico.AnoOE, 1, objServico.ItemOE)
                                        
    bRetorno = objDadosTecnicos.RecuperarDados(objServico.NumOE, _
                                            objServico.AnoOE, 1, objServico.SubItemOE)
    '======================================================================
    '======================================================================
    ' Se ocorreu algum erro na DLL exibi-lo
    If Not bRetorno Then
        
        GoTo Finalizar
    
    End If
    
    ' Atualiza os campos
    objFormOTSCorrecao.lblDesigCcto.Caption = objDadosTecnicos.DesignacaoCircuito
    objFormOTSCorrecao.lblOTSassociada.Caption = objDadosTecnicos.Ots
    objFormOTSCorrecao.lblConexInternet.Caption = objDadosTecnicos.ConexaoInternet
    objFormOTSCorrecao.lblProtocolo.Caption = objDadosTecnicos.Protocolo
    objFormOTSCorrecao.lblVelocidadeCircuito.Caption = objDadosTecnicos.Velocidade_WebIP
    objFormOTSCorrecao.lblCIRSIR.Caption = objDadosTecnicos.CIR_SIR
    objFormOTSCorrecao.lblBGP.Caption = objDadosTecnicos.TipoBGP
    objFormOTSCorrecao.lblPontoCentral.Caption = objDadosTecnicos.PontoCentral
    objFormOTSCorrecao.lblCPE.Caption = sTemCPE
    objFormOTSCorrecao.lblSMTP.Caption = objDadosTecnicos.SMTP
    objFormOTSCorrecao.lblSLA.Caption = objDadosTecnicos.SLA
    
    PreencheCamposDadosTecnicos_BLD = True
    
Finalizar:

    Set objDadosTecnicos = Nothing
    
    On Error GoTo 0
    
    Exit Function
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: PreencheCamposDadosTecnicos_BLD " & _
                            "da classe clsgenOC_OTSCorrecao"
    
    GoTo Finalizar

End Function


Private Function PreencheCamposDadosTecnicos_WID(plServicoId As Long) As Boolean

On Error GoTo ErrorHandler:

    Dim bRetorno As Boolean
    Dim objDadosTecnicos As clswidDadosTecnicos
    
    PreencheCamposDadosTecnicos_WID = False
    
    ' Classe de dados técnicos
    Set objDadosTecnicos = New clswidDadosTecnicos
    
    ' Informar o objeto do banco para as classes
    '
    ' Classe de dados técnicos
    objDadosTecnicos.InformaBanco frmgenMDI_SGAPlus.objBanco
    
    ' Informar o objeto de mensagem para as classes
    '
    ' Classe de dados cadastrais
    objDadosTecnicos.InformaMensagem objMensagem
    
    ' Executar o método para recuperar os dados cadastrais
'VCB 08/08/20119 ---- Causa raíz WID --- performance
'    bRetorno = objDadosTecnicos.RecuperarDados(objServico.ID_Contrato, _
                                               objServico.NR_LOCAL_INSTA, _
                                               objServico.ItemOE)
    
    bRetorno = objDadosTecnicos.RecuperarDados(objServico.ID_Contrato, _
                                               objServico.NR_LOCAL_INSTA, _
                                               objServico.ItemOE, "", plServicoId)
'------------------------------------------------------------------------------------

    ' Se ocorreu algum erro na DLL exibi-lo
    If Not bRetorno Then
        
        GoTo Finalizar
    
    End If
    
    ' Atualiza os campos
    objFormOTSCorrecao.lblDesigCcto.Caption = objDadosTecnicos.DesignacaoCircuito
    objFormOTSCorrecao.lblOTSassociada.Caption = objDadosTecnicos.Ots
    objFormOTSCorrecao.lblConexInternet.Caption = objDadosTecnicos.ConexaoInternet
    objFormOTSCorrecao.lblProtocolo.Caption = objDadosTecnicos.Protocolo
    objFormOTSCorrecao.lblVelocidadeCircuito.Caption = objDadosTecnicos.Velocidade_WebIP
    objFormOTSCorrecao.lblCIRSIR.Caption = objDadosTecnicos.CIR_SIR
    objFormOTSCorrecao.lblBGP.Caption = objDadosTecnicos.TipoBGP
    objFormOTSCorrecao.lblPontoCentral.Caption = objDadosTecnicos.PontoCentral
    objFormOTSCorrecao.lblCPE.Caption = sTemCPE
    objFormOTSCorrecao.lblSMTP.Caption = objDadosTecnicos.SMTP
    objFormOTSCorrecao.lblSLA.Caption = objDadosTecnicos.SLA
    
    PreencheCamposDadosTecnicos_WID = True
    
Finalizar:

    Set objDadosTecnicos = Nothing
    
    On Error GoTo 0
    
    Exit Function
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: PreencheCamposDadosTecnicos_WID " & _
                            "da classe clsgenOC_OTSCorrecao"
    
    GoTo Finalizar

End Function

Public Function GerarOTS() As Boolean
    On Error GoTo ErrorHandler:

    Dim bRetorno                        As Boolean
    Dim lID_SERVICO_PRINCIPAL           As Long
    Dim lID_SERVICO_ANTERIOR            As Long
    Dim lID_SERVICO_REDUNDANTE          As Long
    Dim Flag_Iniciou_Transacao          As Boolean
    Dim bGeraFases                      As Boolean
    Dim sStatusServicoAnterior          As String
    '#BLD - Correção Operacional
    Dim objOrderEntry_BLD               As clsbldOrderEntry
    Dim objDadosCtrlCorrecao            As New clsGenDadosControleCorrecao
    Dim objdadosServico                 As clsGenServico
    
    Dim objAnotacao                     As clsGenAnotacao
    Dim objDetalheAnotacao              As clsGenDetalheAnotacao
    'Dim objFase                              As clsgenFase
    
    Dim lCodFaseOTSNova                 As Long
    Dim lFaseIDAux                      As Long
    Dim objDadosFase                    As clsgenFase
    Dim iSequenciaAcaoAux               As Integer
    Dim IDFaseCentroLocalOTSAnt         As Long
    Dim IDFaseCentroLocalOTSNova        As Long
    Dim objDadosAlocarCF                As clsBLF_AlocacaoCF
    Dim objdetalhecf                    As clsGen_DetalheAlocarCF
    
    Dim sFlagMotivoCorrecao             As String
    
    '================================================================
    'Callbat - 07/04/2015
    'Chamado CH-74999UGL
    '================================================================
    Dim lID_SERVICO_ACESSO              As Long
    Dim ObjAcesso                       As New clsGenAcesso
    '================================================================
    '================================================================
    
    Screen.MousePointer = vbArrowHourglass
     
    GerarOTS = False
    
    Flag_Iniciou_Transacao = False
    
    ' Zera as propriedadedes de identificação de OTS
    sOTS = ""
    
    ' Buscar a sigla do produto
    sSiglaProduto = "BF"
    
    ' Iniciar a transação
    frmgenMDI_SGAPlus.objBanco.Iniciar_Transacao
        
    Flag_Iniciou_Transacao = True
    
    ' Preenche a sigla do produto
    sSiglaProduto = ""
    If objServico.ID_Produto = GbIDProdutoWID Then
    
        sSiglaProduto = "WD"
    
    ElseIf objServico.ID_Produto = GbIDProdutoBLF Then
    
        sSiglaProduto = "BF"
        
    'IP VPN
    ElseIf objServico.ID_Produto = GbIDProdutoBIV Then
        
        sSiglaProduto = "VN"
        
    ElseIf objServico.ID_Produto = GbIDProdutoAAC Then
    
        sSiglaProduto = "AAC"
    
    ElseIf objServico.ID_Produto = GbIDProdutoVIP Then
    
        sSiglaProduto = "VLI"
    
    'JAH
    ElseIf objServico.ID_Produto = GbIDProdutoJAH Then
        
        sSiglaProduto = "JAH"
        
    'ElseIf (objServico.ID_Produto = GbIDProdutoBLD Or _
            objServico.ID_Produto = GbIDProdutoVOIP) Then
    
    ElseIf (objServico.ID_Grupo = GbIDGrupoBLD Or _
            objServico.ID_Produto = GbIDProdutoVOIP) Then
    
        sSiglaProduto = "IP"
        
        '#BLD - Correção Operacional
        Set objOrderEntry_BLD = New clsbldOrderEntry
        
        objOrderEntry_BLD.InformaBanco frmgenMDI_SGAPlus.objBanco
        objOrderEntry_BLD.InformaMensagem objMensagem
    
    End If
    
    ' Zera as variáveis de retorno
    lID_SERVICO_ANTERIOR = objServico.ID_Servico
    lID_SERVICO_REDUNDANTE = objServico.ID_SERVICO_REDUNDANTE
    lID_SERVICO_PRINCIPAL = 0
    
    objServico.ID_Servico_Anterior = objServico.ID_Servico
    sStatusServicoAnterior = objServico.StatusServico
    
    'IP VPN
    'If (objServico.ID_Produto <> GbIDProdutoBIV And _
            objServico.ID_Produto <> GbIDProdutoJAH And _
            objServico.ID_Produto <> GbIDProdutoBLD And _
            objServico.ID_Produto <> GbIDProdutoVOIP) Then
    
    If (objServico.ID_Produto <> GbIDProdutoBIV And _
            objServico.ID_Produto <> GbIDProdutoJAH And _
            objServico.ID_Grupo <> GbIDGrupoBLD And _
            objServico.ID_Produto <> GbIDProdutoVOIP) Then
            
        bRedundante = (objServico.ID_SERVICO_REDUNDANTE = 0) Or IsNull(objServico.ID_SERVICO_REDUNDANTE)
        
    End If
    
    ' Se o servico já foi corrigido
    If (sStatusServicoAnterior = "CR") Then
    
        objMensagem.TipoMsg = "A"
        objMensagem.Descricao = "Serviço já foi corrigido."
        
        GoTo Finalizar
    
    ' Se o serviço já foi cancelado
    ElseIf (sStatusServicoAnterior = "CAN") Then
    
        objMensagem.TipoMsg = "A"
        objMensagem.Descricao = "Serviço já foi cancelado."
        
        GoTo Finalizar
            
    End If
    
    ' Atribui o valor a designação do cricuito
    sDesignacaoCircuito = objCircuito.CNL_Ponta_A & _
                          "/" & _
                          sSiglaProduto & _
                          "/" & _
                          IIf(Len(CStr(objCircuito.NumeroDesignacao)) > 5, objCircuito.NumeroDesignacao, Format(objCircuito.NumeroDesignacao, "00000"))
                          
    '#MTLP - 17/11/04 - Alteração formato Designação JAh!
    If (sSiglaProduto = GbCodigo_Produto_JAH) Then
        sDesignacaoCircuito = Replace(sDesignacaoCircuito, "/", "-")
    End If
    '#MTLP - Fim
    
    ' Verificar a ação do serviço anterior, se estiver em andamento,
    ' Será gerado uma correção, se ativado gera uma alteração
    If (sStatusServicoAnterior = "AN") Then
    
        ' Ser a ação anterior for de alteração ou correção de alteração,
        ' gera uma correção de alteração
        If (objAcao.Sigla = "ALT") Or _
           (objAcao.Sigla = "ALG") Then
           
           objServico.Acao = "ALG"
           
        ' Ser a ação anterior for de ativação ou correção de ativação,
        ' gera uma correção de ativação
        ElseIf (objAcao.Sigla = "ATV") Or _
           (objAcao.Sigla = "ATG") Then
           
           objServico.Acao = "ATG"
           
        End If
        
    ' Gera uma alteração
    Else
    
        objServico.Acao = "ALT"
        
    End If
    
    bRetorno = objServico.Buscar_ID_Acao(objServico.Acao)
    
    If Not bRetorno Then
    
        GoTo Finalizar
        
    End If
    
    
    If objFormOTSCorrecao.OptMotivoCorrecao(0).Value = True Then
        
        sFlagMotivoCorrecao = "M"
    
    Else
        
        sFlagMotivoCorrecao = "A"
    
    End If
    
    objServico.FlagMotivoCorrecao = sFlagMotivoCorrecao
    
    'Gerar o serviço
    bGeraFases = True
    bRetorno = GerarServicoAtivacao(bGeraFases)

    If Not bRetorno Then
    
        GoTo Finalizar
        
    End If
    
    'Atribuir os valores de OTS
    sOTS = CStr(objServico.NumeroOTS) & "/" & CStr(Year(Date))
            
    'Se for correção ou cancelamento, paralisar o serviço anterior
    If (sStatusServicoAnterior = "AN") Then
    
        bRetorno = objServico.AtualizaStatus(lID_SERVICO_ANTERIOR, "CR")

    End If

    'Alterar o campo correção Fase para que OTSs ativadas saiam da caixa de correção
    bRetorno = objFase.GravarCorrecao(lFase_Id, "N")

    'Trata erro
    If Not bRetorno Then

        GoTo Finalizar

    End If

    Set objdadosServico = New clsGenServico

    objdadosServico.InformaBanco frmgenMDI_SGAPlus.objBanco
    objdadosServico.InformaMensagem objMensagem
    
    bRetorno = objdadosServico.GravarCorrecao(lID_SERVICO_ANTERIOR, "N")
    
    If objMensagem.TipoMsg <> "" Then
    
        GoTo ErrorHandler
        
    End If
    
    bRetorno = objdadosServico.GravarCorrecao(objServico.ID_Servico, "", sFlagMotivoCorrecao)
    
    If objMensagem.TipoMsg <> "" Then
    
        GoTo ErrorHandler
        
    End If
    
    Set objDadosCtrlCorrecao = New clsGenDadosControleCorrecao
    
    objDadosCtrlCorrecao.InformaBanco frmgenMDI_SGAPlus.objBanco
    objDadosCtrlCorrecao.InformaMensagem objMensagem
    
    objDadosCtrlCorrecao.DataExecucao = Format(Date, "dd/mm/yyyy")
    objDadosCtrlCorrecao.UsuarioIDExec = GbID_Usuario
    
    bRetorno = objDadosCtrlCorrecao.AtualizarExecucao(lFase_Id, "N")
    
    If objMensagem.TipoMsg <> "" Then

        GoTo ErrorHandler
        
    End If
       
    'Se o tipo de circuito for SHADOW ou DIVERSE
    If (objCircuito.TipoCircuito = "03") Or _
       (objCircuito.TipoCircuito >= "51") Then
    
        'Se for o redundante
        If bRedundante Then
        
            'Buscar o número do ID principal
            bRetorno = objServico.BuscarPrincipal(lID_SERVICO_ANTERIOR, _
                                                  lID_SERVICO_PRINCIPAL)
            If Not bRetorno Then

                GoTo Finalizar
                
            End If
        
            'Atualizar a tabela de servico
            bRetorno = objServico.GravaID_ServicoRedundante(lID_SERVICO_PRINCIPAL, _
                                                            objServico.ID_Servico, _
                                                            objServico.Acao)
                                                    
            If Not bRetorno Then
        
                GoTo Finalizar
            
            End If
        
        End If
        
    End If
    
    '#BLD - Correção Operacional
    'Informa ao GSI a correção operacional do circuito atualizando a OTS
    'If (objServico.ID_Produto = GbIDProdutoBLD Or _
            objServico.ID_Produto = GbIDProdutoVOIP) Then
    
      If (objServico.ID_Grupo = GbIDGrupoBLD Or _
            objServico.ID_Produto = GbIDProdutoVOIP) Then
    
        '###########################################
        'CALLBAT - 26/03/2010
        ' Ajuste feito para corrigir o problema de sempre atualizar a designação do
        'item 1 na tabela ORDER_ENTRY_SUB_ITEM
        '###########################################
        'bRetorno = objOrderEntry_BLD.AtualizarSOE(objServico.NumOE, _
                                objServico.AnoOE, _
                                True, _
                                "03", _
                                objServico.ItemOE, _
                                , _
                                , _
                                objServico.NumeroOTS, _
                                objServico.AnoOTS, _
                                sDesignacaoCircuito, _
                                objServico.ID_CENTRO_FUNCIONAL)
                                
        bRetorno = objOrderEntry_BLD.AtualizarSOE(objServico.NumOE, _
                                objServico.AnoOE, _
                                True, _
                                "03", _
                                objServico.SubItemOE, _
                                , _
                                , _
                                objServico.NumeroOTS, _
                                objServico.AnoOTS, _
                                sDesignacaoCircuito, _
                                objServico.ID_CENTRO_FUNCIONAL)
        '###########################################
        '###########################################

        If (Not bRetorno) Then
    
            GoTo Finalizar
        
        End If
        
    End If
    
    'Recupera as anotações efetuadas na OTS original
    Set objAnotacao = New clsGenAnotacao
    
    objAnotacao.InformaBanco frmgenMDI_SGAPlus.objBanco
    objAnotacao.InformaMensagem objMensagem

    Set objDetalheAnotacao = New clsGenDetalheAnotacao

    bRetorno = objAnotacao.RecuperarAnotacaoServico(lID_SERVICO_ANTERIOR, "9999")
    
    If Not (bRetorno) = True Then
    
        GoTo Finalizar
    
    End If
    
    
    For Each objDetalheAnotacao In objAnotacao.ObjColecaoAnotacao

        If Trim(objDetalheAnotacao.Anotacao) <> "" Then
            
            objDetalheAnotacao.Anotacao = TrataPlick(objDetalheAnotacao.Anotacao)
            
            lCodFaseOTSNova = objFase.ObtemIDFase(objServico.ID_Servico, objDetalheAnotacao.CodigoTipoFase)

            objDetalheAnotacao.InformaBanco frmgenMDI_SGAPlus.objBanco
            objDetalheAnotacao.InformaMensagem objMensagem

            objDetalheAnotacao.iD = lCodFaseOTSNova

            bRetorno = objDetalheAnotacao.IncluirDados("A")

            If Not (bRetorno) = True Then
                GoTo Finalizar
            End If

       End If

    Next objDetalheAnotacao
    
    Set objDadosFase = New clsgenFase
    
    objDadosFase.InformaBanco frmgenMDI_SGAPlus.objBanco
    objDadosFase.InformaMensagem objMensagem
    
    'Finaliza a fase de analise de OE
    lFaseIDAux = objDadosFase.ObtemIDFase(objServico.ID_Servico, GbCodigoFaseAnalisarOE)
    objDadosFase.Usuario_ID = GbID_Usuario
    bRetorno = objDadosFase.IniciaFase(lFaseIDAux, GbID_Usuario)
    
    If (objMensagem.TipoMsg <> "") Then
    
       GoTo ErrorHandler
       
    End If
    
    objDadosFase.Usuario_ID = GbID_Usuario
    
    bRetorno = objDadosFase.FinalizaFase(lFaseIDAux)
    
    If (objMensagem.TipoMsg <> "") Then
    
           GoTo ErrorHandler
       
    End If
    
    'Finalizar a fase Alocação centro funcional
    
    
    'Obtem o ID fase de Alocação
    lFaseIDAux = objDadosFase.ObtemIDFase(objServico.ID_Servico, GbCodigoFaseAlocacao)
    objDadosFase.Usuario_ID = GbID_Usuario
    bRetorno = objDadosFase.IniciaFase(lFaseIDAux, GbID_Usuario)
    
    If (objMensagem.TipoMsg <> "") Then
    
       GoTo ErrorHandler
       
    End If
    
    'AGRC - 04/11/2016 - SO-MKTIP-18424CTY
    Set objCircuito = New clsgenCircuito
    objCircuito.InformaBanco frmgenMDI_SGAPlus.objBanco
    objCircuito.InformaMensagem objMensagem
    'Recuperar os dados do circuito
    bRetorno = objCircuito.RecuperarDados(objServico.ID_Servico)
    If objCircuito.EBkp3G = False Then
        objDadosFase.Usuario_ID = GbID_Usuario
        
        bRetorno = objDadosFase.FinalizaFase(lFaseIDAux)
        
        If (objMensagem.TipoMsg <> "") Then
        
               GoTo ErrorHandler
           
        End If
    End If
    '--- Tratar Alocacao de centro funcional para a fase Centro Local
    'Obtem o ID fase acesso da OTS anterior
    IDFaseCentroLocalOTSAnt = objDadosFase.ObtemIDFase(objServico.ID_Servico_Anterior, GbCodigoFaseCentroLocal)
    
    'Obtem o ID fase acesso da OTS Nova
    IDFaseCentroLocalOTSNova = objDadosFase.ObtemIDFase(objServico.ID_Servico, GbCodigoFaseCentroLocal)
    
    'Calcular o período previsto da fase Tratar Acesso
    iSequenciaAcaoAux = objDadosFase.ObtemSequenciadaFase(objServico.ID_Servico, GbCodigoFaseAlocacao)
    
    Set objDadosAlocarCF = New clsBLF_AlocacaoCF

    objDadosAlocarCF.InformaMensagem objMensagem
    objDadosAlocarCF.InformaBanco frmgenMDI_SGAPlus.objBanco
     
    bRetorno = objDadosAlocarCF.MontaColecaoVariosCFPorFase(IDFaseCentroLocalOTSAnt)
     
    If bRetorno Then
     
        For Each objdetalhecf In objDadosAlocarCF.ObjColecaoVariosCFPorFase
        
            bRetorno = objDadosAlocarCF.InserirCFVarios(IDFaseCentroLocalOTSNova, objdetalhecf.iD)

        Next
        
    End If
    '--- Fim Tratar Alocacao de centro funcional para a fase Centro Local
    
    bRetorno = objDadosFase.CalculaPeriodoProximaFase(iSequenciaAcaoAux, objServico.ID_Servico)
    
    'Finaliza a fase de Alocacao de centro funcional
    lFaseIDAux = objDadosFase.ObtemIDFase(objServico.ID_Servico, GbCodigoFaseAlocacao)
    objDadosFase.Usuario_ID = GbID_Usuario
    bRetorno = objDadosFase.IniciaFase(lFaseIDAux, GbID_Usuario)
    
    If (objMensagem.TipoMsg <> "") Then
    
       GoTo ErrorHandler
       
    End If
    
    'AGRC - 04/11/2016 - SO-MKTIP-18424CTY
    If objCircuito.EBkp3G = False Then
    
        objDadosFase.Usuario_ID = GbID_Usuario
         
        bRetorno = objDadosFase.FinalizaFase(lFaseIDAux)
        
        If (objMensagem.TipoMsg <> "") Then
        
               GoTo ErrorHandler
           
        End If
    End If
    If objFormOTSCorrecao.OptMotivoCorrecao(1).Value = True Then
        
        If objServico.TipoAcessoCCTO <> "C" Then
        
        
        
          '================================================================
          'Callbat - 07/04/2015
          'Chamado CH-74999UGL
          '================================================================
            'Verifica o id do ultimo caso de acesso valido para replicar na tabela de servico
            'bRetorno = TrataSimulacaoFaseAcessoEBT(objServico.ID_Servico, objServico.ID_Servico_Anterior, IIf(objFormOTSCorrecao.OptMotivoCorrecao(1).Value = True, "A", "M"))
            
            ObjAcesso.InformaBanco frmgenMDI_SGAPlus.objBanco
            ObjAcesso.InformaMensagem objMensagem
            
            lID_SERVICO_ACESSO = ObjAcesso.ObtemUltIDAcesso(objServico.ID_CIRCUITO)
            
            
          '================================================================
          '================================================================
           
           
           bRetorno = TrataSimulacaoFaseAcessoEBT(objServico.ID_Servico, lID_SERVICO_ACESSO, IIf(objFormOTSCorrecao.OptMotivoCorrecao(1).Value = True, "A", "M"))
        
        Else
           
           bRetorno = TrataSimulacaoFaseAcessoCLI()
        
        End If
            
        If Not bRetorno Then
            
            GoTo Finalizar
                
        End If
            
    End If
        
    
    'Gravar atributo na tabela SERVICO
    'Obrigar a seleção de uma opção da tela
    'acertar tratamento de erro
    '
    
    ' Finalizar transação
    frmgenMDI_SGAPlus.objBanco.Finalizar_Transacao
    
    Flag_Iniciou_Transacao = False
    
    ' Chama a tela e exibição de designação e OTS
    bRetorno = TelaExibirDesignacaoOTS()
    
        
    If Not bRetorno Then
        
        GoTo Finalizar
            
    End If
        
    GerarOTS = True
    
Finalizar:

    '#BLD - Correção Operacional
    Set objOrderEntry_BLD = Nothing
    
    Set objAnotacao = Nothing
    Set objDetalheAnotacao = Nothing
    
    
    Screen.MousePointer = vbDefault
    
    If Flag_Iniciou_Transacao Then
    
        ' Cancelar Transação
        frmgenMDI_SGAPlus.objBanco.Cancelar_Transacao
            
    End If
    
    If objMensagem.TipoMsg <> "" Then
    
        objMensagem.ExibeMensagem
        
        objMensagem.TipoMsg = ""

    End If
    
    On Error GoTo 0
    
    Exit Function
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: GerarOTS " & _
                            "da classe clsgenOC_OTSCorrecao"
    
    GoTo Finalizar

End Function

Public Function GerarServicoAtivacao(ByVal pbGeraFases As Boolean) As Boolean

On Error GoTo ErrorHandler:

    Dim bRetorno                        As Boolean
    Dim lIDFaseAux                      As Long
    Dim objDadosCircuito                As clsgenCircuito
    
    Dim lServicoAnteriorValido          As Long
    
    '''===FABIMENDONCA - 04/2021 - SO-SOLEN-28213JNU - Acesso 4G - Rede Wireless
    Dim sIndWireless                    As String
    Dim objOrderEntry_BLD               As clsbldOrderEntry
    Dim sSEVWirelessAux                 As String
    ''''-----------------------------------------
    
    Screen.MousePointer = vbArrowHourglass
    
    GerarServicoAtivacao = False
    
    ' Excutar o método de criação de número de OTS
    bRetorno = objServico.GerarNumeroOTS()
    
    If Not bRetorno Then
        
        GoTo Finalizar
        
    End If
    
    ' Executar o método de gravação dos dados de servico
    objServico.ID_USUARIO = GbID_Usuario
    
    bRetorno = objServico.GravarDados(objServico.NumOE, _
                                      objServico.AnoOE, _
                                      objServico.ItemOE, _
                                      objServico.NumContrato, _
                                      objServico.AnoContrato, _
                                      objServico.VersaoContrato, _
                                      objServico.ID_CIRCUITO, _
                                      objServico.ID_Contrato, _
                                      0, _
                                      "SGA", _
                                      "", _
                                      "", _
                                      objServico.DiretorRegional, _
                                      objServico.Possui_CPE, _
                                      objServico.NR_LOCAL_INSTA, _
                                      objServico.TipoAcessoCCTO, _
                                      objServico.Possui_BGP, objServico.FlagMotivoCorrecao)
                                      
                                          
    If Not bRetorno Then
    
        GoTo Finalizar
    
    End If

    ' Executa o método para buscar o número do serviço
    bRetorno = objServico.BuscarNumeroServicoAtual()
    
    If Not bRetorno Then
        
        GoTo Finalizar
        
    End If
    
    
    Set objDadosCircuito = New clsgenCircuito
    
    objDadosCircuito.InformaBanco frmgenMDI_SGAPlus.objBanco
    objDadosCircuito.InformaMensagem objMensagem
    
    '=================================================================
    'CALLBAT - 21/05/2012 - BSoD Net Light
    'Inclusão do ultimo usuário valido para a designação.
    '=================================================================
    If objServico.Acao = "CAN" Then
    
        lServicoAnteriorValido = objServico.ObtemServicoAnteriorValido(objServico.ID_CIRCUITO, objServico.ID_Servico)
        
        bRetorno = objDadosCircuito.AtualizarUltimoServicoID(objServico.ID_CIRCUITO, lServicoAnteriorValido)
        
    Else
        bRetorno = objDadosCircuito.AtualizarUltimoServicoID(objServico.ID_CIRCUITO, objServico.ID_Servico)
    End If
    '=================================================================
    '=================================================================
    
    
    
    
    If Not bRetorno Then
        
        GoTo Finalizar
        
    End If
    
    ' Verifica se realmente é para gerar as fases
    If pbGeraFases Then
    
        'Verifica se a OTS anterior possui CPE
        lIDFaseAux = objFase.ObtemIDFase(objServico.ID_Servico_Anterior, "22") 'captura o ID da fase de CPE
        
        '''===============================================
        '''FabiMendonca - Projeto Rede Wireless - 4G
        '''Nao gerar mais fase de CPE na correção, estava criando a fase CPE mesmo em casos q nao havia CPE marcado
        ''If lIDFaseAux <> 0 then
        If lIDFaseAux <> 0 Or lIDFaseAux = 0 Then
        
            'objServico.Possui_CPE = "S"
            'TBASTOS CH-41732KAC Não gera mais fase de CPE para OTS de correção.
            objServico.Possui_CPE = "N"
            
        End If
            
        'Verifica se a OTS anterior possui CPE
        lIDFaseAux = objFase.ObtemIDFase(objServico.ID_Servico_Anterior, "11") 'captura o ID da fase de BGP
        
        '''===============================================
        '''FabiMendonca - Projeto Rede Wireless - 4G
        '''Nao gerar mais fase de CPE na correção, estava criando a fase CPE mesmo em casos q nao havia CPE marcado
        ''If lIDFaseAux <> 0 then
        If lIDFaseAux <> 0 Or lIDFaseAux = 0 Then
        
            'objServico.Possui_CPE = "S"
            'TBASTOS CH-41732KAC Não gera mais fase de CPE para OTS de correção.
            objServico.Possui_CPE = "N"
            
        End If
            
        ' Criar as fases do Workflow
        bRetorno = objFase.CriarFases(objServico.ID_Produto, _
                                      objServico.ID_Servico, _
                                      objServico.ID_Servico_Anterior, _
                                      objServico.Id_Acao, _
                                      objServico.ID_CENTRO_FUNCIONAL, _
                                      objServico.Acao, _
                                      "", _
                                      "SGA", _
                                      objServico.Possui_CPE, _
                                      objServico.Possui_BGP, _
                                      objServico.ID_CENTRO_FUNCIONAL)
                                      
        If Not bRetorno Then
        
            GoTo Finalizar
            
        End If
        
        '=================================================================
        '''===FABIMENDONCA - 04/2021 - SO-SOLEN-28213JNU - Acesso 4G - Rede Wireless
        '''===Quando for Wireless NAO devera ter FASES :  codigo_tipo_fase = 8 (TRATAR CENTRO LOCAL)
        '''=== Nem codigo_tipo_fase = 3 (PRÉ-ATIVAÇÃO)
        '=================================================================
        sIndWireless = ""
        
        Set objOrderEntry_BLD = New clsbldOrderEntry
        objOrderEntry_BLD.InformaBanco frmgenMDI_SGAPlus.objBanco
        objOrderEntry_BLD.InformaMensagem objMensagem
        
        sIndWireless = objOrderEntry_BLD.VerificaIndicWireless(objServico.ID_Servico)
                                                               
        If sIndWireless = "S" Then
            'captura o ID da fase  codigo tipo fase = 8 (TRATAR CENTRO LOCAL)
            '''
            lIDFaseAux = objFase.ObtemIDFase(objServico.ID_Servico, "8")
            ''
            If lIDFaseAux <> 0 Then
                ''Verificar e excluir se tiver registro filho na tab centro_funcional
                bRetorno = objFase.ExcluiFaseCentrosFunc(lIDFaseAux)
                If (Not bRetorno) Then
                    GoTo Finalizar
                End If
                
                ''Excluir a fase de Centro Local
                bRetorno = objFase.ApagaFase(lIDFaseAux)
                If (Not bRetorno) Then
                    GoTo Finalizar
                End If
            End If
            'captura o ID da fase  codigo tipo fase =3 (PRÉ-ATIVAÇÃO)
            lIDFaseAux = objFase.ObtemIDFase(objServico.ID_Servico, "3")
            '
            bRetorno = objFase.ApagaFase(lIDFaseAux)
            
            If (Not bRetorno) Then
                GoTo Finalizar
            End If
            
            '=================================================================
            '''---- Gravar o numero SEV na tabela SERVICO
            sSEVWirelessAux = ""
            bRetorno = objServico.Buscar_SEV_Wireless(objServico.ID_Servico_Anterior, sSEVWirelessAux)
            
            If Not bRetorno Then
    
                GoTo ErrorHandler
    
            End If
            ''
            If sSEVWirelessAux <> "" Then
                bRetorno = objServico.AtualizaSEV_Wireless(objServico.ID_Servico, sSEVWirelessAux)
                
                If Not bRetorno Then
                    GoTo ErrorHandler
                End If
            End If
                                
        End If
        '=================================================================
    
    End If
        
    GerarServicoAtivacao = True
    
Finalizar:

    Screen.MousePointer = vbDefault
    
    On Error GoTo 0
    
    Exit Function
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: GerarServicoAtivacao " & _
                            "da classe clsgenOC_OTSCorrecao"
    
    GoTo Finalizar

End Function

Public Function InformaOCSelecionarOTS(ByVal oobjOCSelecionarOTS As clsgenOC_SelecionarOTSCorrecao)

On Error GoTo ErrorHandler:

    InformaOCSelecionarOTS = False
    
    Set objOCSelecionarOTS = oobjOCSelecionarOTS
    
    InformaOCSelecionarOTS = True
    
    GoTo Finalizar
    
Finalizar:

    Exit Function
    
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no Método:" & _
                            "InformaOCSelecionarOTS da classe clsGenOC_OTSCorrecao"
    
    GoTo Finalizar

End Function
Private Function PreencheCamposDadosTecnicos_IVN(ByVal psTipoItem As String, _
    Optional ByVal pbExibeDesativados As Boolean) As Boolean
    'IP VPN
On Error GoTo ErrorHandler:

    Dim bRetorno                As Boolean
    Dim objSites                As clsivnListaSitesTopologia
    Dim objBloco                As clsivnListaIPMascara
    Dim objListaSitesExtranet   As clsivnListaSitesExtranet
    Dim vSitesExtranet          As String
    Dim objDadosTecnicos        As clsivnDadosTecnicos
    
    PreencheCamposDadosTecnicos_IVN = False
    
    ' Classe de dados técnicos
    Set objDadosTecnicos = New clsivnDadosTecnicos
    
    objDadosTecnicos.InformaBanco frmgenMDI_SGAPlus.objBanco
    objDadosTecnicos.InformaMensagem objMensagem
    
    ' Executar o método para recuperar os dados cadastrais
    bRetorno = objDadosTecnicos.RecuperarDados_Site(objServico.ID_Contrato, _
                                               objServico.NR_LOCAL_INSTA, _
                                               objServico.AcaoSOE, _
                                               CodigoProduto)
    
    ' Se ocorreu algum erro na DLL exibi-lo
    If (Not bRetorno) Then
        
        GoTo Finalizar
    
    End If
    
    ' Atualiza os campos
    objFormOTSCorrecao.LblOTSTopologia_Site.Caption = objDadosTecnicos.OTSTopologia

    objFormOTSCorrecao.LblDesignacaoCircuito_Site = objDadosTecnicos.DesignacaoCircuito
    objFormOTSCorrecao.LblNomeSite_Site = objDadosTecnicos.NomeSite
    objFormOTSCorrecao.LblNomeVPN_Site = objDadosTecnicos.NomeVPN
    
    If UCase(Trim(objDadosTecnicos.Protocolo)) = "ATM" Then
        'Valores ATM
        objFormOTSCorrecao.LblProtocolo_ATM.Caption = objDadosTecnicos.Protocolo
        objFormOTSCorrecao.LblVelocidade_ATM.Caption = objDadosTecnicos.Velocidade_WebIP
        objFormOTSCorrecao.LblVPI_ATM.Caption = objDadosTecnicos.NR_VPI
        objFormOTSCorrecao.LblVCI_ATM.Caption = objDadosTecnicos.NR_VCI
        objFormOTSCorrecao.LblSIR_ATM.Caption = objDadosTecnicos.CIR_SIR
        
    ElseIf (UCase(Trim(objDadosTecnicos.Protocolo)) = "FR" Or _
                UCase(Trim(objDadosTecnicos.Protocolo)) = "FRAMERELAY") Then
        
        'Valores FR
        objFormOTSCorrecao.FramATM.Visible = False
        objFormOTSCorrecao.FramFR.Visible = True
        
        objFormOTSCorrecao.LblProtocolo_FR.Caption = objDadosTecnicos.Protocolo
        objFormOTSCorrecao.LblVelocidade_FR.Caption = objDadosTecnicos.Velocidade_WebIP
        objFormOTSCorrecao.LblSIR_FR.Caption = objDadosTecnicos.CIR_SIR
        
        'LMI
        If (objDadosTecnicos.IC_TIPO_LMI = "V") Then
            
            objFormOTSCorrecao.LblLMI_FR.Caption = "V.Forum"
            
        ElseIf objDadosTecnicos.IC_TIPO_LMI = "I" Then
        
            objFormOTSCorrecao.LblLMI_FR.Caption = "ITU-T"
        
        ElseIf objDadosTecnicos.IC_TIPO_LMI = "A" Then
        
            objFormOTSCorrecao.LblLMI_FR.Caption = "Ansi"
        
        ElseIf objDadosTecnicos.IC_TIPO_LMI = "C" Then
        
            objFormOTSCorrecao.LblLMI_FR.Caption = "CCITT"
        
        Else
            
            objFormOTSCorrecao.LblLMI_FR.Caption = objDadosTecnicos.IC_TIPO_LMI
        
        End If
        
        objFormOTSCorrecao.LblDLCI_FR.Caption = objDadosTecnicos.NR_DLCI
    Else
        'Outros protocolos
        objFormOTSCorrecao.LblProtocolo_ATM.Caption = objDadosTecnicos.Protocolo
        objFormOTSCorrecao.LblVelocidade_ATM.Caption = objDadosTecnicos.Velocidade_WebIP
        objFormOTSCorrecao.LblVPI_ATM.Caption = objDadosTecnicos.NR_VPI
        objFormOTSCorrecao.LblVCI_ATM.Caption = objDadosTecnicos.NR_VCI
        objFormOTSCorrecao.LblSIR_ATM.Caption = objDadosTecnicos.CIR_SIR
    
    End If
    
    
    If objDadosTecnicos.TipoSite = "H" Then
        objFormOTSCorrecao.optTPSiteHub_Site.Value = True
        objFormOTSCorrecao.optTPSiteSpoke_Site.Value = False
    
    ElseIf objDadosTecnicos.TipoSite = "S" Then
        objFormOTSCorrecao.optTPSiteHub_Site.Value = False
        objFormOTSCorrecao.optTPSiteSpoke_Site.Value = True
    
    Else
        objFormOTSCorrecao.optTPSiteHub_Site.Value = False
        objFormOTSCorrecao.optTPSiteSpoke_Site.Value = False
    
    End If
    
    If objDadosTecnicos.TipoVPN = "H" Then
        objFormOTSCorrecao.optTPVPNHubSpoke_Site.Value = True
        objFormOTSCorrecao.optTPVPNFull_Site.Value = False
        objFormOTSCorrecao.optTPVPNMista_Site.Value = False
    
    ElseIf objDadosTecnicos.TipoVPN = "F" Then
        objFormOTSCorrecao.optTPVPNHubSpoke_Site.Value = False
        objFormOTSCorrecao.optTPVPNFull_Site.Value = True
        objFormOTSCorrecao.optTPVPNMista_Site.Value = False
    
    ElseIf objDadosTecnicos.TipoVPN = "M" Then
        objFormOTSCorrecao.optTPVPNHubSpoke_Site.Value = False
        objFormOTSCorrecao.optTPVPNFull_Site.Value = False
        objFormOTSCorrecao.optTPVPNMista_Site.Value = True
    
    Else
        objFormOTSCorrecao.optTPVPNHubSpoke_Site.Value = False
        objFormOTSCorrecao.optTPVPNFull_Site.Value = False
        objFormOTSCorrecao.optTPVPNMista_Site.Value = False
    
    End If
    
    
    If objDadosTecnicos.TemBalanceamento = "S" Then
        objFormOTSCorrecao.optBalanceamentoSim_Site.Value = True
        objFormOTSCorrecao.optBalanceamentoNao_Site.Value = False

    Else
        objFormOTSCorrecao.optBalanceamentoSim_Site.Value = False
        objFormOTSCorrecao.optBalanceamentoNao_Site.Value = True

    End If
    
    'Exibe os Sites e Designações Extranet
    If (objDadosTecnicos.objSitesExtra.Count > 0) Then
    
        Set objListaSitesExtranet = New clsivnListaSitesExtranet
        
        For Each objListaSitesExtranet In objDadosTecnicos.objSitesExtra
            
            vSitesExtranet = ""
            
            If (Trim(objListaSitesExtranet.Designacao) <> "") Then
                
                vSitesExtranet = Trim(objListaSitesExtranet.Designacao) & _
                            "     -      " & objListaSitesExtranet.NomeSite
                    
            Else
                vSitesExtranet = objListaSitesExtranet.NomeSite
            End If
            
            objFormOTSCorrecao.LstDesigExtranetSites_Site.AddItem vSitesExtranet
            
        Next objListaSitesExtranet

    End If
    
    
    If objDadosTecnicos.Extranet = "S" Then
        
        objFormOTSCorrecao.optExtranetSim_Site.Value = True
        objFormOTSCorrecao.optExtranetNao_Site.Value = False
    
    Else
        objFormOTSCorrecao.optExtranetSim_Site.Value = False
        objFormOTSCorrecao.optExtranetNao_Site.Value = True

    End If
    
    If objDadosTecnicos.Roteamento = "E" Then
        objFormOTSCorrecao.optRoteamentoEstatico_Site.Value = True
        objFormOTSCorrecao.optRoteamentoBGP_Site.Value = False

    ElseIf objDadosTecnicos.Roteamento = "B" Then
        objFormOTSCorrecao.optRoteamentoEstatico_Site.Value = False
        objFormOTSCorrecao.optRoteamentoBGP_Site.Value = True
    
    Else
        objFormOTSCorrecao.optRoteamentoEstatico_Site.Value = False
        objFormOTSCorrecao.optRoteamentoBGP_Site.Value = False
    
    End If
    
    If (objDadosTecnicos.Gpa <> "" And _
            objDadosTecnicos.Gpa <> "N") Then
        
        objFormOTSCorrecao.LbGPA_Site = "S"
        objFormOTSCorrecao.LbTipoGPA_Site = objDadosTecnicos.TipoGPA

    Else
        objFormOTSCorrecao.LbGPA_Site = "N"
        objFormOTSCorrecao.LbTipoGPA_Site = ""
    End If
    
    'Bloco IP
    If objDadosTecnicos.objBlocoIP.Count > 0 Then
    
        Set objBloco = New clsivnListaIPMascara
        
        For Each objBloco In objDadosTecnicos.objBlocoIP
            objFormOTSCorrecao.LstBlocosEndereco_Site.AddItem Trim(objBloco.NumeroIP) & _
                "/" & objBloco.NumeroMascara
            
        Next objBloco
        
    End If
    
    objFormOTSCorrecao.LblSiteBalanceamento_Site = objDadosTecnicos.SiteBalanceamento

    objFormOTSCorrecao.LblIPPeerBPG_Site.Caption = objDadosTecnicos.IPPeerBPG
    objFormOTSCorrecao.LblValorBGP.Caption = objDadosTecnicos.TipoBGP
    objFormOTSCorrecao.LbSLA_Site.Caption = objDadosTecnicos.SLA
    
    objFormOTSCorrecao.LbDtPrevATV_Site.Caption = IIf(IsDate(objDadosTecnicos.DataPrevATV), _
            Format(objDadosTecnicos.DataPrevATV, "dd/mm/yyyy"), "")

    If (objDadosTecnicos.TipoCircuito = "VL") Then
        
        objFormOTSCorrecao.lblTipoCcto.Caption = "Light"
    
    ElseIf objDadosTecnicos.TipoCircuito = "VP" Then
        
        objFormOTSCorrecao.lblTipoCcto.Caption = "Premium"
        
    Else
        
        objFormOTSCorrecao.lblTipoCcto.Caption = ""
            
    End If
        
    PreencheCamposDadosTecnicos_IVN = True
    
Finalizar:
    Set objSites = Nothing
    Set objBloco = Nothing
    Set objListaSitesExtranet = Nothing
    Set objDadosTecnicos = Nothing
    
    On Error GoTo 0
    
    Exit Function
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: PreencheCamposDadosTecnicos_IVN " & _
                            "da classe clsgen_OTSCorrecao"
    GoTo Finalizar

End Function




Private Function PreencheCamposDadosTecnicos_JAH() As Boolean
    'JAH
On Error GoTo ErrorHandler:

    Dim bRetorno                As Boolean
    Dim objDadosTecnicos        As clsjahDadosTecnicos
    
    PreencheCamposDadosTecnicos_JAH = False
    
    ' Classe de dados técnicos
    Set objDadosTecnicos = New clsjahDadosTecnicos
    
    objDadosTecnicos.InformaBanco frmgenMDI_SGAPlus.objBanco
    objDadosTecnicos.InformaMensagem objMensagem
    
    ' Executar o método para recuperar os dados cadastrais
    bRetorno = objDadosTecnicos.RecuperarDados(objServico.ID_Contrato, _
                                               objServico.NR_LOCAL_INSTA, _
                                               objServico.AcaoSOE, _
                                               CodigoProduto)
    ' Se ocorreu algum erro na DLL exibi-lo
    If (Not bRetorno) Then
        
        GoTo Finalizar
    
    End If
    
    objFormOTSCorrecao.LblDesignacao_JAH.Caption = objDadosTecnicos.DesignacaoCircuito
    objFormOTSCorrecao.LblUsernameJAH.Caption = objDadosTecnicos.UserName
    objFormOTSCorrecao.LblVelocidadeJAH.Caption = objDadosTecnicos.Velocidade_WebIP
    
    
    If (objServico.TipoItem = GbDADOSVOZ) Then
        
        objFormOTSCorrecao.lblTipoCcto.Caption = "VOZ"
    
    ElseIf (objServico.TipoItem = GbDADOSBANDALARGA) Then
        
        objFormOTSCorrecao.lblTipoCcto.Caption = "BANDA LARGA"
        
    ElseIf (objServico.TipoItem = GbDADOSVOZ_BANDALARGA) Then
        
        objFormOTSCorrecao.lblTipoCcto.Caption = "VOZ E BANDA LARGA"
        
    Else
        
        objFormOTSCorrecao.lblTipoCcto.Caption = ""
            
    End If
        
    PreencheCamposDadosTecnicos_JAH = True
    
Finalizar:
    Set objDadosTecnicos = Nothing
    
    On Error GoTo 0
    
    Exit Function
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: PreencheCamposDadosTecnicos_JAH " & _
                            "da classe clsgen_OTSCorrecao"
    GoTo Finalizar

End Function

'Inicio - PTS 1244 - Joao Carlos Barros - 05/05/2010
Public Function GerarOTSCan() As Boolean
    On Error GoTo ErrorHandler:

    Dim bRetorno                        As Boolean
    Dim lID_SERVICO_PRINCIPAL           As Long
    Dim lID_SERVICO_ANTERIOR            As Long
    Dim lID_SERVICO_REDUNDANTE          As Long
    Dim Flag_Iniciou_Transacao          As Boolean
    Dim bGeraFases                      As Boolean
    Dim sStatusServicoAnterior          As String
    Dim objOrderEntry_BLD               As clsbldOrderEntry
    Dim objDadosCtrlCorrecao            As New clsGenDadosControleCorrecao
    Dim objdadosServico                 As clsGenServico
    Dim objAnotacao                     As clsGenAnotacao
    Dim objDetalheAnotacao              As clsGenDetalheAnotacao
    Dim lCodFaseOTSNova                 As Long
    Dim lID_SERVICO_ANTERIOR_ANTERIOR   As Long
    Dim lsStatus_SERVICO_ANTERIOR_ANTERIOR   As String
    Dim lIDFase                         As Long
    Dim lSequencia                      As Long
    Dim objDadosIntegracao              As New clsBlf_IntegracaoWebIP
    Dim objDadosCircuito                As clsgenCircuito
        
    Screen.MousePointer = vbArrowHourglass
     
    GerarOTSCan = False
    
    Flag_Iniciou_Transacao = False
    
    If objServico.Origem = "SOE" Then
    
        objMensagem.TipoMsg = "A"
        objMensagem.Descricao = "Não é possível gerar OTS de cancelamento para OTSs geradas a partir do aceite de uma Order Entry!" & vbCrLf & " Favor providenciar a geração de OE de cancelamento no GSI."
        
        GoTo Finalizar
        
    End If
    
    ' Zera as propriedadedes de identificação de OTS
    sOTS = ""
    
    ' Preenche a sigla do produto
    sSiglaProduto = ""
    If objServico.ID_Produto = GbIDProdutoWID Then
        sSiglaProduto = "WD"
    ElseIf objServico.ID_Produto = GbIDProdutoBLF Then
        sSiglaProduto = "BF"
    'IP VPN
    ElseIf objServico.ID_Produto = GbIDProdutoBIV Then
        sSiglaProduto = "VN"
    ElseIf objServico.ID_Produto = GbIDProdutoAAC Then
        sSiglaProduto = "AAC"
    ElseIf objServico.ID_Produto = GbIDProdutoVIP Then
        sSiglaProduto = "VLI"
    'JAH
    ElseIf objServico.ID_Produto = GbIDProdutoJAH Then
        sSiglaProduto = "JAH"
    ElseIf (objServico.ID_Grupo = GbIDGrupoBLD Or objServico.ID_Produto = GbIDProdutoVOIP) Then
        sSiglaProduto = "IP"
        
        Set objOrderEntry_BLD = New clsbldOrderEntry
        
        objOrderEntry_BLD.InformaBanco frmgenMDI_SGAPlus.objBanco
        objOrderEntry_BLD.InformaMensagem objMensagem
    End If
    
    ' Zera as variáveis de retorno
    
    bRetorno = objServico.Buscar_Status(objServico.ID_Servico_Anterior)
    If Not bRetorno Then
        GoTo Finalizar
    End If
    lID_SERVICO_ANTERIOR_ANTERIOR = objServico.ID_Servico_Anterior
    lsStatus_SERVICO_ANTERIOR_ANTERIOR = objServico.Status
    
    lID_SERVICO_ANTERIOR = objServico.ID_Servico
    lID_SERVICO_REDUNDANTE = objServico.ID_SERVICO_REDUNDANTE
    lID_SERVICO_PRINCIPAL = 0
    objServico.ID_Servico_Anterior = objServico.ID_Servico
    sStatusServicoAnterior = objServico.StatusServico
    
    If (objServico.ID_Produto <> GbIDProdutoBIV And _
            objServico.ID_Produto <> GbIDProdutoJAH And _
            objServico.ID_Grupo <> GbIDGrupoBLD And _
            objServico.ID_Produto <> GbIDProdutoVOIP) Then
            
        bRedundante = (objServico.ID_SERVICO_REDUNDANTE = 0) Or IsNull(objServico.ID_SERVICO_REDUNDANTE)
        
    End If
    
    ' Atribui o valor a designação do cricuito
    sDesignacaoCircuito = objCircuito.CNL_Ponta_A & _
                          "/" & _
                          sSiglaProduto & _
                          "/" & _
                          IIf(Len(CStr(objCircuito.NumeroDesignacao)) > 5, objCircuito.NumeroDesignacao, Format(objCircuito.NumeroDesignacao, "00000"))
                          
    If (sSiglaProduto = GbCodigo_Produto_JAH) Then
        sDesignacaoCircuito = Replace(sDesignacaoCircuito, "/", "-")
    End If
       
    objServico.Acao = "CAN"
    
    bRetorno = objServico.Buscar_ID_Acao(objServico.Acao)
    
    If Not bRetorno Then
        GoTo Finalizar
    End If
    
    ' Iniciar a transação
    frmgenMDI_SGAPlus.objBanco.Iniciar_Transacao
        
    Flag_Iniciou_Transacao = True
    
    'Gerar o serviço
    bGeraFases = False
    bRetorno = GerarServicoAtivacao(bGeraFases)
    
    If Not bRetorno Then
        GoTo Finalizar
    End If
    
    'Gera a fase especifica de Analise de Order Entry, finalizada
    '0 - Nao iniciada, 1 - Iniciada, 2 - Finalizada
    bRetorno = objFase.CriarFaseEspecifica(objServico.ID_Servico, objServico.ID_Produto, objServico.Id_Acao, GbCodigoFaseAnalisarOE, objServico.ID_CENTRO_FUNCIONAL, objServico.ID_USUARIO, 2)
    If Not bRetorno Then
        GoTo Finalizar
    End If
    
    'Atribuir os valores de OTS
    sOTS = CStr(objServico.NumeroOTS) & "/" & CStr(Year(Date))
            
    'Cancelar o serviço anterior
    bRetorno = objServico.AtualizaStatus(lID_SERVICO_ANTERIOR, "CAN")
    'Trata erro
    If Not bRetorno Then
        GoTo Finalizar
    End If
    
    Set objDadosCtrlCorrecao = New clsGenDadosControleCorrecao
    
    objDadosCtrlCorrecao.InformaBanco frmgenMDI_SGAPlus.objBanco
    objDadosCtrlCorrecao.InformaMensagem objMensagem
    
    'Verifica se o servico anterior estava finalizado , caso positivo não faz nada.
    If lsStatus_SERVICO_ANTERIOR_ANTERIOR = "FF" Then
       'não altera status
    Else
        'Voltar status do servico anterior ao cancelado para Em Andamento
        bRetorno = objServico.AtualizaStatus(lID_SERVICO_ANTERIOR_ANTERIOR, "AN")
        'Trata erro
        If Not bRetorno Then
            GoTo Finalizar
        End If
        
        lIDFase = objDadosCtrlCorrecao.ObtemIdFaseCorrecao(lID_SERVICO_ANTERIOR_ANTERIOR)
        
        bRetorno = objDadosCtrlCorrecao.RecuperarDados(lIDFase, "C")
        
        objFase.Data_Inicio_Previsto = objDadosCtrlCorrecao.DataInicioPrevAnt
        objFase.Data_Fim_Previsto = objDadosCtrlCorrecao.DataFimPrevAnt
        
        bRetorno = objFase.AtualizaPeriodoPrevisto(lIDFase)
        
        bRetorno = objFase.GravarCorrecao(lIDFase, "")
        If Not bRetorno Then
           GoTo Finalizar
        End If
        
        'ATUALIZA INTEGRACAO_BLF
        objDadosIntegracao.InformaBanco frmgenMDI_SGAPlus.objBanco
        objDadosIntegracao.InformaMensagem objMensagem
        bRetorno = objDadosIntegracao.AtualizarProcessado(lID_SERVICO_ANTERIOR_ANTERIOR)
        If Not bRetorno Then
           GoTo Finalizar
        End If
        'DELETAR REGISTRO DA CORREÇÃO ANTERIOR
        bRetorno = objDadosIntegracao.Deletar_Correcao_Cancelada(lID_SERVICO_ANTERIOR)
        If Not bRetorno Then
           GoTo Finalizar
        End If
      
    End If
    
    'Atualiza na tabela circuito com o ID da última OTS válida
    Set objDadosCircuito = New clsgenCircuito
    
    objDadosCircuito.InformaBanco frmgenMDI_SGAPlus.objBanco
    objDadosCircuito.InformaMensagem objMensagem
    
    bRetorno = objDadosCircuito.AtualizarUltimoServicoID(objCircuito.ID_CIRCUITO, lID_SERVICO_ANTERIOR_ANTERIOR)
    If Not bRetorno Then
       GoTo Finalizar
    End If

    
    'Finalizar o cancelamento
    bRetorno = objServico.FinalizaServico(objServico.ID_Servico)
    'Trata erro
    If Not bRetorno Then
        GoTo Finalizar
    End If
    
    'Alterar o campo correção Fase para que OTSs ativadas saiam da caixa de correção
    bRetorno = objFase.GravarCorrecao(lIDFase, "N")
    'Trata erro
    If Not bRetorno Then
        GoTo Finalizar
    End If

    Set objdadosServico = New clsGenServico

    objdadosServico.InformaBanco frmgenMDI_SGAPlus.objBanco
    objdadosServico.InformaMensagem objMensagem

    bRetorno = objdadosServico.GravarCorrecao(lID_SERVICO_ANTERIOR, "N")
    
    If objMensagem.TipoMsg <> "" Then
    
        GoTo ErrorHandler
        
    End If
    
    objDadosCtrlCorrecao.DataExecucao = Format(Date, "dd/mm/yyyy")
    objDadosCtrlCorrecao.UsuarioIDExec = GbID_Usuario
    
    bRetorno = objDadosCtrlCorrecao.AtualizarExecucao(lFase_Id, "N")
    
    If objMensagem.TipoMsg <> "" Then

        GoTo ErrorHandler
        
    End If
       
    'Se o tipo de circuito for SHADOW ou DIVERSE
    If (objCircuito.TipoCircuito = "03") Or _
       (objCircuito.TipoCircuito >= "51") Then
    
        'Se for o redundante
        If bRedundante Then
        
            'Buscar o número do ID principal
            bRetorno = objServico.BuscarPrincipal(lID_SERVICO_ANTERIOR, _
                                                  lID_SERVICO_PRINCIPAL)
            If Not bRetorno Then

                GoTo Finalizar
                
            End If
        
            'Atualizar a tabela de servico
            bRetorno = objServico.GravaID_ServicoRedundante(lID_SERVICO_PRINCIPAL, _
                                                            objServico.ID_Servico, _
                                                            objServico.Acao)
                                                    
            If Not bRetorno Then
        
                GoTo Finalizar
            
            End If
        
        End If
        
    End If
    
'    If (objServico.ID_Grupo = GbIDGrupoBLD Or _
'            objServico.ID_PRODUTO = GbIDProdutoVOIP) Then
'
'        bRetorno = objOrderEntry_BLD.AtualizarSOE(objServico.NumOE, _
'                                objServico.AnoOE, _
'                                True, _
'                                "03", _
'                                objServico.SubItemOE, _
'                                , _
'                                , _
'                                objServico.NumeroOTS, _
'                                objServico.AnoOTS, _
'                                sDesignacaoCircuito, _
'                                objServico.ID_CENTRO_FUNCIONAL)
'
'        If (Not bRetorno) Then
'
'            GoTo Finalizar
'
'        End If
'
'    End If
    
    'Recupera as anotações efetuadas na OTS original
'    Set objAnotacao = New clsGenAnotacao
'
'    objAnotacao.InformaBanco frmgenMDI_SGAPlus.objBanco
'    objAnotacao.InformaMensagem objMensagem
'
'    Set objDetalheAnotacao = New clsGenDetalheAnotacao
'
'    bRetorno = objAnotacao.RecuperarAnotacaoServico(lID_SERVICO_ANTERIOR, "9999")
'
'    If Not (bRetorno) = True Then
'
'        GoTo Finalizar
'
'    End If
'
'    For Each objDetalheAnotacao In objAnotacao.ObjColecaoAnotacao
'
'        If Trim(objDetalheAnotacao.Anotacao) <> "" Then
'
'            objDetalheAnotacao.Anotacao = TrataPlick(objDetalheAnotacao.Anotacao)
'
'            lCodFaseOTSNova = objFase.ObtemIDFase(objServico.ID_SERVICO, objDetalheAnotacao.CodigoTIpoFase)
'            '
'            objDetalheAnotacao.InformaBanco frmgenMDI_SGAPlus.objBanco
'            objDetalheAnotacao.InformaMensagem objMensagem
'
'            objDetalheAnotacao.iD = lCodFaseOTSNova
'
'            bRetorno = objDetalheAnotacao.IncluirDados("A")
'
'            If Not (bRetorno) = True Then
'                GoTo Finalizar
'            End If
'
'       End If
'
'    Next objDetalheAnotacao
    
    ' Finalizar transação
    frmgenMDI_SGAPlus.objBanco.Finalizar_Transacao
    
    Flag_Iniciou_Transacao = False
    
    ' Chama a tela e exibição de designação e OTS
    bRetorno = TelaExibirDesignacaoOTS()
        
    If Not bRetorno Then
        
        GoTo Finalizar
            
    End If
        
    GerarOTSCan = True
    
Finalizar:

    '#BLD - Correção Operacional
    Set objOrderEntry_BLD = Nothing
    
    Set objAnotacao = Nothing
    Set objDetalheAnotacao = Nothing
    
    
    Screen.MousePointer = vbDefault
    
    If Flag_Iniciou_Transacao Then
    
        ' Cancelar Transação
        frmgenMDI_SGAPlus.objBanco.Cancelar_Transacao
            
    End If
    
    If objMensagem.TipoMsg <> "" Then
    
        objMensagem.ExibeMensagem
        
        objMensagem.TipoMsg = ""

    End If
    
    On Error GoTo 0
    
    Exit Function
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: GerarOTSCan " & _
                            "da classe clsgenOC_OTSCorrecao"
    
    GoTo Finalizar

End Function

Public Function ValidarCancelamento() As Boolean

    On Error GoTo ErrorHandler:

    Dim objFase                  As clsgenFase
    Dim objDadosIntegracao       As New clsBlf_IntegracaoWebIP

    ValidarCancelamento = False
    
    If objServico.StatusServico = "AN" Then
        
        If objServico.Acao = "ATG" Or objServico.Acao = "ALG" Or (objServico.Acao = "ALT" And objServico.Origem = "SGA") Then
            
            Set objFase = New clsgenFase
            
            objFase.InformaBanco frmgenMDI_SGAPlus.objBanco
            objFase.InformaMensagem objMensagem
            
            objFase.ID_Servico = objServico.ID_Servico
           'Comentada para não verificar pela fase Tratar acesso , esse item será descomentado
           'quando entrar a interface com o CLA.
           ' objFase.Id_Fase = objServico.ObtemIDFase(objFase.ID_Servico, GbCodigoFaseAcesso)
           'objFase.Id_Fase = objServico.ObtemIDFase(objFase.ID_Servico, GbCodigoFaseConfiguracao)
           
           'Conforme definido por conta da interface do CLA, a restrição para o cancelamento será a fase de acesso e não mais a fase
           'de configuração. Vagner 22/10/2010
           
           If objServico.FlagMotivoCorrecao = "M" Then '--- Migracao de acesso
              
              objFase.Id_Fase = objServico.ObtemIDFase(objFase.ID_Servico, GbCodigoFaseAcesso)
           
           Else
              
              objFase.Id_Fase = objServico.ObtemIDFase(objFase.ID_Servico, GbCodigoFaseConfiguracao)
           
           End If
           
           If objFase.VerificaFaseIniciada(objFase.Id_Fase) Then
              
              objMensagem.TipoMsg = "A"
              
              If objServico.FlagMotivoCorrecao = "M" Then '--- Migracao de acesso
                 
                 objMensagem.Descricao = "Essa OTS não pode ser cancelada pois a fase de Acesso já foi iniciada e/ou finalizada!"
              
              Else
                 objMensagem.Descricao = "Essa OTS não pode ser cancelada pois a fase de Configuração (SAIP) já foi iniciada e/ou finalizada!"
              
              End If
           
           Else
              ValidarCancelamento = True
              GoTo Finalizar
           
           End If
           
           'If Not (objFase.FaseFinalizada(objFase.Id_Fase) Or objFase.FaseIniciada(objFase.Id_Fase)) Then
           '     ValidarCancelamento = True
           '     GoTo Finalizar
           ' Else
           
           If objFase.Id_Fase = 0 Then
           
                ValidarCancelamento = True
                GoTo Finalizar
                
'           Else
'                objDadosIntegracao.InformaBanco frmgenMDI_SGAPlus.objBanco
'                objDadosIntegracao.InformaMensagem objMensagem
'                If objDadosIntegracao.ExisteProcessadoIntegracao(objServico.ID_Servico) Or objFase.FaseFinalizada(objFase.Id_Fase) Then
'                   objMensagem.TipoMsg = "A"
'                   objMensagem.Descricao = "Essa OTS não pode ser cancelada pois a fase de Configuração(SAIP) já foi iniciada ou finalizada!"
'                Else
'                   ValidarCancelamento = True
'                   GoTo Finalizar
'                End If
           End If
        Else
            objMensagem.TipoMsg = "A"
            objMensagem.Descricao = "Essa OTS não pode ser cancelada!"
        End If
    Else
        objMensagem.TipoMsg = "A"
        objMensagem.Descricao = "Essa OTS não pode ser cancelada, pois não está em andamento!"
        
    End If
    
Finalizar:

    Set objFase = Nothing
    
    If objMensagem.TipoMsg <> "" Then
    
        objMensagem.ExibeMensagem
        
        objMensagem.TipoMsg = ""

    End If
    
    On Error GoTo 0
    
    Exit Function
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: ValidarCancelamento " & _
                            "da classe clsgenOC_OTSCorrecao"
    
    GoTo Finalizar

End Function
'Fim - PTS 1244 - Joao Carlos Barros - 05/05/2010

Public Function TrataSimulacaoFaseAcessoEBT(plServicoId As Long, plServicoIDAnterior As Long, psFlagMotivo As String) As Boolean

'Obter o ID do serviço anterior
'Obter o ID da fase do Servico anterior
'Recuperar dados da contingencia_acesso
    
    Dim objDadosProduto              As clsGenProduto
    Dim sSituacaoOTS                 As String
    Dim numUltimaFaseRealizada       As Long
    Dim IdOTS                        As Long
    Dim IDFaseOTSAnt                 As Long
    Dim lFaseIDOTSNova               As Long
    Dim lFaseIDAux                   As Long
    Dim sAcaoAux                     As String
    Dim sSiglaProduto                As String
    Dim objDadosEqptoNGN             As VIP_Equipamento_NGN
    Dim objDadosLogSGAP              As clsBLD_LogSGAPCLA
    Dim iAux                         As Long
    Dim objAcessoCLA                 As clsGenAcessoCLA
    Dim objListaAcessosCLA           As clsGenListaAcessoCLA
    Dim sVar                         As Variant
    Dim bRetorno                     As Boolean
    Dim objDadosFase                 As clsgenFase
    Dim iSequenciaAcaoAux            As Integer
    Dim objWEBIP                     As clsInterOC_TrataWebIP
    Dim objLogSGAPCLA                As clsBLD_LogSGAPCLA
    Dim sStatusAnterior              As String
    Dim objdadosServico              As clsGenServico
    Dim lServicoIdCla                As Long
    Dim lCircuitoID                  As Long
    Dim objCircuito                  As clsgenCircuito
    
    TrataSimulacaoFaseAcessoEBT = False
    
    If objMensagem Is Nothing Then
        
        Set objMensagem = New clsgenMensagem
        
    End If
    
    TrataSimulacaoFaseAcessoEBT = False

    Set objDadosFase = New clsgenFase
    
    objDadosFase.InformaBanco frmgenMDI_SGAPlus.objBanco
    objDadosFase.InformaMensagem objMensagem
    
    
    
    'Obtem o ID fase acesso da OTS anterior
    IDFaseOTSAnt = objDadosFase.ObtemIDFase(plServicoIDAnterior, GbCodigoFaseAcesso)
    
    'Obtem o ID fase acesso da OTS Nova
    lFaseIDOTSNova = objDadosFase.ObtemIDFase(plServicoId, GbCodigoFaseAcesso)
    
    Set objAcessoCLA = New clsGenAcessoCLA
     
    objAcessoCLA.InformaBanco frmgenMDI_SGAPlus.objBanco
    objAcessoCLA.InformaMensagem objMensagem
     
    bRetorno = objAcessoCLA.RecuperarAcessos(IDFaseOTSAnt, objMensagem)
     
    If bRetorno Then
         
        If objAcessoCLA.CollectAcessos.Count > 0 Then
            
            For Each objListaAcessosCLA In objAcessoCLA.CollectAcessos
                
                bRetorno = RecuperaDadosAcessoPorID(CLng(objListaAcessosCLA.iD), lFaseIDOTSNova, IDFaseOTSAnt)
                    
            Next
            
            If objMensagem.TipoMsg <> "" Then
                
                GoTo ErrorHandler
            
            End If
            
            If psFlagMotivo = "A" Then
                
                'Inserir registro na log_sgap para a OTS de correção
                Set objLogSGAPCLA = New clsBLD_LogSGAPCLA
                
                objLogSGAPCLA.InformaBanco frmgenMDI_SGAPlus.objBanco
                objLogSGAPCLA.InformaMensagem objMensagem
                
                If objLogSGAPCLA.RecuperarDados(plServicoIDAnterior) Then
                    
                    Set objdadosServico = New clsGenServico
                    
                    objdadosServico.InformaBanco frmgenMDI_SGAPlus.objBanco
                    objdadosServico.InformaMensagem objMensagem
                    
                    lCircuitoID = objdadosServico.ObtemIDCircuito(plServicoId)
                    
                    lServicoIdCla = objLogSGAPCLA.ObtemServicoEnviadoAoCLA(lCircuitoID)
                   
                   bRetorno = objLogSGAPCLA.ReplicaRegistro(plServicoIDAnterior, plServicoId, lServicoIdCla)
                    
                   If Not bRetorno Then
                        
                        GoTo ErrorHandler
                   
                   End If
                   
                   'AGRC - CH-83814ZEK - 18/01/2017
                   Set objCircuito = New clsgenCircuito
                   objCircuito.InformaBanco frmgenMDI_SGAPlus.objBanco
                   objCircuito.InformaMensagem objMensagem
                   'Recuperar os dados do circuito
                   bRetorno = objCircuito.RecuperarDados(plServicoId)
                        
                    If Not bRetorno Then
                    
                        GoTo Finalizar
                        
                    End If
                   If objCircuito.EBkp3G = False Then
                        objDadosFase.Usuario_ID = GbID_Usuario
                        bRetorno = objDadosFase.IniciaFase(lFaseIDOTSNova, GbID_Usuario)
                         
                        If (objMensagem.TipoMsg <> "") Then
                         
                         GoTo ErrorHandler
                            
                        End If
                    End If
                End If
            
            End If
        
        End If
     
     End If
     
     TrataSimulacaoFaseAcessoEBT = True
     
Finalizar:

    Set objDadosEqptoNGN = Nothing
    Set objDadosLogSGAP = Nothing
    Set objAcessoCLA = Nothing
    Set objListaAcessosCLA = Nothing
    Set objDadosFase = Nothing
    Set objWEBIP = Nothing
    Set objLogSGAPCLA = Nothing
    
    Exit Function

ErrorHandler:

    If objMensagem.TipoMsg = "" And Err.Number <> 0 Then

        objMensagem.TipoMsg = "E"
        objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                                "Houve um erro de execução no método: TrataSimulacaoFaseAcessoEBT " & _
                                "da classe clsgenOC_OTSCorrecao"
    
    End If
    
    GoTo Finalizar

End Function

Private Function TrataSimulacaoFaseAcessoCLI()

'Obter o ID do serviço anterior
'Obter o ID da fase do Servico anterior
'Recuperar dados da contingencia_acesso

    Dim objDadosAcessoCliente  As clsblf_TratarAcessoCliente
    Dim IDFaseOTSAnt           As Long
    Dim lFaseIDOTSNova         As Long
    Dim lFaseIDAux             As Long
    Dim iAux                   As Long
    Dim bRetorno               As Boolean
    Dim objDadosFase           As clsgenFase
    Dim iSequenciaAcaoAux      As Integer
    Dim objWEBIP               As clsInterOC_TrataWebIP
    
    Set objDadosFase = New clsgenFase

    '-- VAGNERR 19/09
    TrataSimulacaoFaseAcessoCLI = False
    '--
    objDadosFase.InformaBanco frmgenMDI_SGAPlus.objBanco
    objDadosFase.InformaMensagem objMensagem
    
    'Obtem o ID fase acesso da OTS anterior
    IDFaseOTSAnt = objDadosFase.ObtemIDFase(objServico.ID_Servico_Anterior, GbCodigoFaseAcesso)
    
    'Obtem o ID fase acesso da OTS anterior
    lFaseIDOTSNova = objDadosFase.ObtemIDFase(objServico.ID_Servico, GbCodigoFaseAcesso)
    
    Set objDadosAcessoCliente = New clsblf_TratarAcessoCliente
    
    objDadosAcessoCliente.InformaMensagem objMensagem
    objDadosAcessoCliente.InformaBanco frmgenMDI_SGAPlus.objBanco
     
    If objDadosAcessoCliente.RecuperarDados(IDFaseOTSAnt) Then
        
        bRetorno = objDadosAcessoCliente.IncluirDados(lFaseIDOTSNova)
    
        If (objMensagem.TipoMsg <> "") Then
        
           GoTo ErrorHandler
           
        End If
        objDadosFase.Usuario_ID = GbID_Usuario
        bRetorno = objDadosFase.IniciaFase(lFaseIDOTSNova, GbID_Usuario)
    
        If (objMensagem.TipoMsg <> "") Then
        
           GoTo ErrorHandler
           
        End If
        
        'Calcular o periodo previsto da fase Acesso
        iSequenciaAcaoAux = objDadosFase.ObtemSequenciadaFase(objServico.ID_Servico, GbCodigoFaseAlocacao)
        
        bRetorno = objDadosFase.CalculaPeriodoProximaFase(iSequenciaAcaoAux, objServico.ID_Servico)
    
    
        If objDadosFase.ObtemDataFimReal(IDFaseOTSAnt) <> "" Then
            
            bRetorno = objDadosFase.FinalizaFase(lFaseIDOTSNova)
        
            If (objMensagem.TipoMsg <> "") Then
            
                   GoTo ErrorHandler
               
            End If
        
            'Calcular o periodo previsto da fase Configuracao
            iSequenciaAcaoAux = objDadosFase.ObtemSequenciadaFase(objServico.ID_Servico, GbCodigoFaseAcesso)
            
            bRetorno = objDadosFase.CalculaPeriodoProximaFase(iSequenciaAcaoAux, objServico.ID_Servico)
        
            Set objWEBIP = New clsInterOC_TrataWebIP
            
            objWEBIP.Grupo_ID = lGrupo_ID
                      
            bRetorno = objWEBIP.GravaWEBIPIntegracao(objServico.ID_Servico, _
                            lFaseIDOTSNova, _
                            "C", _
                            Produto_ID, _
                            objMensagem, _
                            objServico.Acao)
        
            If (objMensagem.TipoMsg <> "") Then
        
                GoTo ErrorHandler
                
            End If
        
        End If
        
    End If
    
        TrataSimulacaoFaseAcessoCLI = True
Finalizar:

    Set objDadosAcessoCliente = Nothing
    Set objDadosFase = Nothing
    Set objWEBIP = Nothing
    
    Exit Function

ErrorHandler:

    If objMensagem.TipoMsg = "" And Err.Number <> 0 Then

        objMensagem.TipoMsg = "E"
        objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                                "Houve um erro de execução no método: TrataSimulacaoFaseAcessoCLI " & _
                                "da classe clsgenOC_OTSCorrecao"
    
    End If
    
    GoTo Finalizar


End Function

Public Function RecuperaDadosAcessoPorID(plId As Long, plFaseIDOTSNova As Long, plIDFaseOTSAnt As Long) As Boolean
        
    Dim bRetorno As Boolean
    
    Dim objDadosPropriedadeEBT  As clsGenPropriedadeEmbratel
    Dim objDadosPropriedadeTerc As clsGenPropriedadeTerceiros
    Dim lContingenciaAcesso_ID  As Long
    Dim lContingenciaIDAnt      As Long
    Dim sCaracteristicaAcesso   As String
    
    Dim objDadosTipoRadio       As clsGenTipoAcessoRadio
    Dim objDadosTipoFibra       As clsGenTipoAcessoFibra
    Dim objdadosTipoADE         As clsGenTipoAcessoADE
    Dim objDadosTipoCabo        As clsGenTipoAcessoCabo
    Dim objdadosTipoAcessoMenor As clsGenTipoAcessoMenor2M
    Dim objdadosTipoAcessoMaior As clsGenTipoAcessoMaior2M
    Dim objdadosTipoAcessoRede  As clsGenTipoAcessoRedeDeterm
    Dim objdadosTipoBSod        As ClsGenTipoAcessoBSOD
    Dim objdadosTipoGPON        As clsGenTipoAcessoGPON
    Dim objDadosAcesso          As clsGenAcesso
    Dim objListaAcessos         As clsGenAcessoCLA
    Dim objdadosTipoFOEthernet  As clsGenTipoAcessoFOEtherNET
    
    '=================================================================
    'Callbat - 06/05/2014
    'Alterado para atender ao chamado CH-61100EVY do Gati
    '=================================================================
    Dim objDadosTipoFO_EDD As clsGenTipoAcessoFO_EDD
    '=================================================================
    '--=================================================================
    '-- 18/02/2019 -  ADE DSLAM - "M" '--SO-SOLEN-08414BBJ_SGAP
    '--=================================================================
    Dim objDadosTipoADE_DSLAM As clsGenTipoAcessoADE_DSLAM
    '=================================================================
    
    '=================================================================
    'CALLBAT - 13/09/2013 - Radio IP
    'SO- TVASS-47468LFW
    '=================================================================
    Dim objdadosTipoAcessoRadioIP As clsGenTipoAcessoRadioIP
    '=================================================================
    '=================================================================
    Dim objdadosTipoAcessoTerceiroETH As New clsGenTipoAcessoTerceiroETH
    On Error GoTo ErrorHandler

    RecuperaDadosAcessoPorID = False
    
    Set objDadosAcesso = New clsGenAcesso
    
    objDadosAcesso.InformaMensagem objMensagem
    objDadosAcesso.InformaBanco frmgenMDI_SGAPlus.objBanco

'    'Não resgata os dados da OTS anterior quando se altera o
'    'endereço de instalação. IP VPN
'    If (Sigla_Acao <> "ATV" And bResgataAnterior) Then
'
'        If TipoOperacao = "I" Then
'
'           pServico_ID = Servico_ID_Anterior
'
'           'GbCodigoFaseAcesso - ID da fase na tabela Tipo_Fase
'           pFase_ID = objdadosServico.ObtemIDFase(Servico_ID_Anterior, GbCodigoFaseAcesso)
'
'           If objmensagem.TipoMsg <> "" Then
'
'              GoTo ErrorHandler
'
'           End If
'
'        End If
'
'    End If
    
    sCaracteristicaAcesso = "P"
    
    If Not objDadosAcesso.RecuperarDadosPorID(plId, objMensagem) Then
    
       If objMensagem.TipoMsg <> "" Then
           
          objMensagem.ExibeMensagem
          
          GoTo Finalizar
        
       End If
       
    Else

        bRetorno = objDadosAcesso.IncluirDados(plFaseIDOTSNova)
        
        If objDadosAcesso.IdFisico <> "" Then
             
             lContingenciaIDAnt = objDadosAcesso.ObtemIDPorAcessoFisico(plIDFaseOTSAnt, objDadosAcesso.IdFisico, objDadosAcesso.CaracteristicaAcesso)
             
             objDadosAcesso.iD = objDadosAcesso.ObtemIDPorAcessoFisico(plFaseIDOTSNova, objDadosAcesso.IdFisico, objDadosAcesso.CaracteristicaAcesso)
        
        Else
            
            lContingenciaIDAnt = objDadosAcesso.ObtemIDContingencia(plIDFaseOTSAnt, objDadosAcesso.CaracteristicaAcesso)
            
            objDadosAcesso.iD = objDadosAcesso.ObtemIDContingencia(plFaseIDOTSNova, objDadosAcesso.CaracteristicaAcesso)

        End If
        
        'Obtem os dados de Equipamento Complementar
        bRetorno = objDadosAcesso.RecuperarDadosEquipamentos(lContingenciaIDAnt, objMensagem)
          
        If objMensagem.TipoMsg <> "" Then
            
           GoTo ErrorHandler
         
        End If
        
        If objDadosAcesso.objColecaoEqpto.Count > 0 Then
           
           bRetorno = GravarEqptosComplementar(objDadosAcesso, plFaseIDOTSNova)
        
        End If
        
        'Caso o Acesso já esteja Cadastrado, Recuperar os dados Conforme a Propriedade do
        'Acesso
        
        If objDadosAcesso.PropridedadeAcesso = "E" Then   'Propriedade Acesso = Embratel
        
           Set objDadosPropriedadeEBT = New clsGenPropriedadeEmbratel
           
           objDadosPropriedadeEBT.InformaBanco frmgenMDI_SGAPlus.objBanco
           
           bRetorno = objDadosPropriedadeEBT.RecuperarDados(plId, objMensagem)
            
           If objMensagem.TipoMsg <> "" Then
              
              GoTo ErrorHandler
           
           End If
           
           If bRetorno Then
                
                bRetorno = objDadosPropriedadeEBT.IncluirDados(objDadosAcesso.iD, objMensagem)
                
                If objMensagem.TipoMsg <> "" Then
                   
                   GoTo ErrorHandler
                
                End If
                
           End If
          
           'Carrega a propriedade Objeto Propriedade Embratel
           If objDadosPropriedadeEBT.TipoAcesso = "R" Then
               
              Set objDadosTipoRadio = New clsGenTipoAcessoRadio
              
              objDadosTipoRadio.InformaBanco frmgenMDI_SGAPlus.objBanco
           
              bRetorno = objDadosTipoRadio.RecuperarDados(plId, objMensagem)
              
              If objMensagem.TipoMsg <> "" Then
                 
                 GoTo ErrorHandler
                 
              End If

              bRetorno = objDadosTipoRadio.IncluirDados(objDadosAcesso.iD, objMensagem)

           ElseIf objDadosPropriedadeEBT.TipoAcesso = "F" Then  'Fibra Otica
              
              Set objDadosTipoFibra = New clsGenTipoAcessoFibra
              
              objDadosTipoFibra.InformaBanco frmgenMDI_SGAPlus.objBanco
           
              bRetorno = objDadosTipoFibra.RecuperarDados(plId, objMensagem)
              
              If objMensagem.TipoMsg <> "" Then
                 
                 GoTo ErrorHandler
                 
              End If
              
              bRetorno = objDadosTipoFibra.IncluirDados(objDadosAcesso.iD, objMensagem)
              
           ElseIf objDadosPropriedadeEBT.TipoAcesso = "A" Then  'ADE
              
              Set objdadosTipoADE = New clsGenTipoAcessoADE
              
              objdadosTipoADE.InformaBanco frmgenMDI_SGAPlus.objBanco
           
              bRetorno = objdadosTipoADE.RecuperarDados(plId, objMensagem)
              
              If objMensagem.TipoMsg <> "" Then
                 
                 GoTo ErrorHandler
                 
              End If
              
              bRetorno = objdadosTipoADE.IncluirDados(objDadosAcesso.iD, objMensagem, False)
           
           ElseIf objDadosPropriedadeEBT.TipoAcesso = "C" Then  'Cabo Interno
              
              Set objDadosTipoCabo = New clsGenTipoAcessoCabo
              
              objDadosTipoCabo.InformaBanco frmgenMDI_SGAPlus.objBanco
           
              bRetorno = objDadosTipoCabo.RecuperarDados(plId, objMensagem)
              
              If objMensagem.TipoMsg <> "" Then
                 
                 GoTo ErrorHandler
                 
              End If
              
              bRetorno = objDadosTipoCabo.IncluirDados(objDadosAcesso.iD, objMensagem)
           
           ElseIf objDadosPropriedadeEBT.TipoAcesso = "G" Then  'GPON
              
              Set objdadosTipoGPON = New clsGenTipoAcessoGPON
              
              objdadosTipoGPON.InformaBanco frmgenMDI_SGAPlus.objBanco
           
              bRetorno = objdadosTipoGPON.RecuperarDados(plId, objMensagem)
              
              If objMensagem.TipoMsg <> "" Then
                 
                 GoTo ErrorHandler
                 
              End If
              
              bRetorno = objdadosTipoGPON.IncluirDados(objDadosAcesso.iD, objMensagem)
              
           '=================================================================
           'Callbat - 06/05/2014
           'Alterado para atender ao chamado CH-61100EVY do Gati
           '=================================================================
           '-- O código abaixo foi inserido para replicar as informações do tipo de acesso FO EDD
           ElseIf objDadosPropriedadeEBT.TipoAcesso = "D" Then   'FO EDD
           '=================================================================
           '=================================================================
              
              Set objDadosTipoFO_EDD = New clsGenTipoAcessoFO_EDD
              
              objDadosTipoFO_EDD.InformaBanco frmgenMDI_SGAPlus.objBanco
           
              bRetorno = objDadosTipoFO_EDD.RecuperarDados(plId, objMensagem)
              
              If objMensagem.TipoMsg <> "" Then
                 
                 GoTo ErrorHandler
                 
              End If
              'objDadosTipoFO_EDD.iD = objServico.ID_Servico
              bRetorno = objDadosTipoFO_EDD.IncluirDados(objDadosAcesso.iD, objMensagem)
           '=================================================================
           
           '--=================================================================
           '-- 18/02/2019 -  ADE DSLAM - "M" '--SO-SOLEN-08414BBJ_SGAP
           '--=================================================================
           ElseIf objDadosPropriedadeEBT.TipoAcesso = "M" Then   'ADE DSLAM
           '=================================================================
           '=================================================================
              
              Set objDadosTipoADE_DSLAM = New clsGenTipoAcessoADE_DSLAM
              
              objDadosTipoADE_DSLAM.InformaBanco frmgenMDI_SGAPlus.objBanco
           
              bRetorno = objDadosTipoADE_DSLAM.RecuperarDados(plId, objMensagem)
              
              If objMensagem.TipoMsg <> "" Then
                 
                 GoTo ErrorHandler
                 
              End If
              'objDadosTipoFO_EDD.iD = objServico.ID_Servico
              bRetorno = objDadosTipoADE_DSLAM.IncluirDados(objDadosAcesso.iD, objMensagem)
           '=================================================================
           '=================================================================
           
           End If
           
           Set objDadosPropriedadeEBT = Nothing
           
           

           
        Else
           
           'Propriedade Acesso = Terceiros
           Set objDadosPropriedadeTerc = New clsGenPropriedadeTerceiros
           
           objDadosPropriedadeTerc.InformaBanco frmgenMDI_SGAPlus.objBanco
           
           bRetorno = objDadosPropriedadeTerc.RecuperarDados(plId, objMensagem)
            
           If objMensagem.TipoMsg <> "" Then
              
              GoTo ErrorHandler
           
           End If
           
           If bRetorno Then
                
                bRetorno = objDadosPropriedadeTerc.IncluirDados(objDadosAcesso.iD, objMensagem)
           
           End If
           
          'Carrega a propriedade Objeto Propriedade Terceiros
           If objDadosPropriedadeTerc.TipoAcesso = "<" Then   'Tipo de Acesso Menor 2M
               
              Set objdadosTipoAcessoMenor = New clsGenTipoAcessoMenor2M
              
              objdadosTipoAcessoMenor.InformaBanco frmgenMDI_SGAPlus.objBanco
           
              bRetorno = objdadosTipoAcessoMenor.RecuperarDados(plId, objMensagem)
              
              If objMensagem.TipoMsg <> "" Then
                 
                 GoTo ErrorHandler
                 
              End If

              bRetorno = objdadosTipoAcessoMenor.IncluirDados(objDadosAcesso.iD, objMensagem)

           ElseIf objDadosPropriedadeTerc.TipoAcesso = ">" Then  'Tipo de Acesso Maior 2M
              
              Set objdadosTipoAcessoMaior = New clsGenTipoAcessoMaior2M
              
              objdadosTipoAcessoMaior.InformaBanco frmgenMDI_SGAPlus.objBanco
           
              bRetorno = objdadosTipoAcessoMaior.RecuperarDados(plId, objMensagem)
              
              If objMensagem.TipoMsg <> "" Then
                 
                 GoTo ErrorHandler
                 
              End If
              
              bRetorno = objdadosTipoAcessoMaior.IncluirDados(objDadosAcesso.iD, objMensagem)
           
           ElseIf objDadosPropriedadeTerc.TipoAcesso = "R" Then  'Tipo de Acesso Rede Determ
              
              Set objdadosTipoAcessoRede = New clsGenTipoAcessoRedeDeterm
              
              objdadosTipoAcessoRede.InformaBanco frmgenMDI_SGAPlus.objBanco
           
              bRetorno = objdadosTipoAcessoRede.RecuperarDados(plId, objMensagem)
              
              If objMensagem.TipoMsg <> "" Then
                 
                 GoTo ErrorHandler
                 
              End If
              
              bRetorno = objdadosTipoAcessoRede.IncluirDados(objDadosAcesso.iD, objMensagem)
           
            '=================================================================
            'CALLBAT - 21/05/2012 - BSoD Net Light
            '=================================================================
           'ElseIf objDadosPropriedadeTerc.TipoAcesso = "V" Or objDadosPropriedadeTerc.TipoAcesso = "N" Then  'VIA ou NET
           
           'ElseIf objDadosPropriedadeTerc.TipoAcesso = "V" Or objDadosPropriedadeTerc.TipoAcesso = "N" Or _
           '           objDadosPropriedadeTerc.TipoAcesso = "L" Then   'VIA / NET / BSoD Light
                      
           ElseIf objDadosPropriedadeTerc.TipoAcesso = "V" Or objDadosPropriedadeTerc.TipoAcesso = "N" Or _
                  objDadosPropriedadeTerc.TipoAcesso = "L" Or objDadosPropriedadeTerc.TipoAcesso = "U" Then   'VIA / NET / BSoD Light / Unificado
                      
            '=================================================================
            '=================================================================
              
              Set objdadosTipoBSod = New ClsGenTipoAcessoBSOD
              
              objdadosTipoBSod.InformaBanco frmgenMDI_SGAPlus.objBanco
           
              bRetorno = objdadosTipoBSod.RecuperarDados(plId, objMensagem)
              
              If objMensagem.TipoMsg <> "" Then
                 
                 GoTo ErrorHandler
                 
              End If
              
              bRetorno = objdadosTipoBSod.IncluirDados(objDadosAcesso.iD, objMensagem)

           
           '-- O código abaixo foi inserido para replicar as informações do tipo de acesso FO Ethernet
           ElseIf objDadosPropriedadeTerc.TipoAcesso = "E" Then   'FO Ethernet
            '=================================================================
            '=================================================================
              
              Set objdadosTipoFOEthernet = New clsGenTipoAcessoFOEtherNET
              
              objdadosTipoFOEthernet.InformaBanco frmgenMDI_SGAPlus.objBanco
           
              bRetorno = objdadosTipoFOEthernet.RecuperarDados(plId, objMensagem)
              
              If objMensagem.TipoMsg <> "" Then
                 
                 GoTo ErrorHandler
                 
              End If
              'objdadosTipoFOEthernet.iD = objServico.ID_SERVICO
              bRetorno = objdadosTipoFOEthernet.IncluirDados(objDadosAcesso.iD, objMensagem)
           
           '=================================================================
           'CALLBAT - 13/09/2013 - Radio IP
           'SO- TVASS-47468LFW
           '=================================================================
           ElseIf objDadosPropriedadeTerc.TipoAcesso = "I" Then   'Radio IP

              
              Set objdadosTipoAcessoRadioIP = New clsGenTipoAcessoRadioIP
              
              objdadosTipoAcessoRadioIP.InformaBanco frmgenMDI_SGAPlus.objBanco
           
              bRetorno = objdadosTipoAcessoRadioIP.RecuperarDados(plId, objMensagem)
              
              If objMensagem.TipoMsg <> "" Then
                 
                 GoTo ErrorHandler
                 
              End If
              
              
              bRetorno = objdadosTipoAcessoRadioIP.IncluirDados(objDadosAcesso.iD, objMensagem)
           
           '=================================================================
           '=================================================================
            'AGRC - SO-TVASS-32377EYC - 25/08/2017
           ElseIf objDadosPropriedadeTerc.TipoAcesso = "H" Then   'Terceiro ETH

              
              Set objdadosTipoAcessoTerceiroETH = New clsGenTipoAcessoTerceiroETH
              
              objdadosTipoAcessoTerceiroETH.InformaBanco frmgenMDI_SGAPlus.objBanco
           
              bRetorno = objdadosTipoAcessoTerceiroETH.RecuperarDados(plId, objMensagem)
              
              If objMensagem.TipoMsg <> "" Then
                 
                 GoTo ErrorHandler
                 
              End If
              
              
              bRetorno = objdadosTipoAcessoTerceiroETH.IncluirDados(objDadosAcesso.iD, objMensagem)
           
           '=================================================================
           '=================================================================
           
           End If
        
        
           Set objDadosPropriedadeTerc = Nothing
        
        End If
    
    End If
    
    RecuperaDadosAcessoPorID = True
    
Finalizar:
    
    On Error GoTo 0
    
    Exit Function
    
ErrorHandler:
       
       If objMensagem.TipoMsg = "" Then
          
          objMensagem.TipoMsg = "E"
          objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                             vbCrLf & "Houve um erro de execução no método: " & vbCrLf & vbCrLf & "RecuperadadosAcessoPorID do Objeto Controle " & _
                             "clsGenOC_OTSCorrecao"
       End If
       
'       objMensagem.ExibeMensagem
       
       GoTo Finalizar

End Function

Private Function GravarEqptosComplementar(pobjAcesso As clsGenAcesso, plFaseId As Long) As Boolean

    Dim ObjEqpto                    As clsGenEqptoComplementar
    Dim objEqptoModem               As clsGenEqptosAcessoModem
    Dim ObjEqptoOrigem              As clsGenEqptoComplementar
    Dim ObjEqptoDestino             As clsGenEqptoComplementar
    Dim aLista()                    As Variant

    Dim iQtEquipamento              As Integer
    Dim bRetorno                    As Boolean
    Dim lIDEquip                    As Long
    Dim lIDEqpto                    As Long
    Dim iCont                       As Integer
    
    If Produto_ID <> GbIDProdutoVIP Then  'Tratamento abaixo serve para todos os
                                                  'produtos exceto VIPLINE IP
        For Each ObjEqpto In pobjAcesso.objColecaoEqpto
                    
            iQtEquipamento = iQtEquipamento + 1
            
            ObjEqpto.InformaBanco frmgenMDI_SGAPlus.objBanco
            
            If Val(ObjEqpto.TipoEquipamento) <> 0 Then
                        
                ObjEqpto.iD = 0
                
                If ObjEqpto.iD = 0 Then 'Inclusao
                   
                   ObjEqpto.Fase_ID = plFaseId
                   
                   bRetorno = ObjEqpto.IncluirDados("E", pobjAcesso.iD, objMensagem)
                 
                   If bRetorno Then
                      
                      lIDEquip = ObjEqpto.ObtemID(iQtEquipamento, pobjAcesso.iD, objMensagem)
                   
                   End If

                End If
                
            End If
                    
            If objMensagem.TipoMsg <> "" Then
             
'               GoTo ErrorHandler
            
            End If
                    
        Next
    
    Else
        If Grupo_ID = GbIDGrupoVLI Then   'Se vipline IP
            
            'Obter o ID da tabela contingencia referente a OTS anterior
        
            For Each ObjEqptoOrigem In pobjAcesso.objColecaoEqpto
    
                'Set ObjEqptoDestino = New clsGenEqptoComplementar
    
                Set ObjEqptoDestino = ObjEqptoOrigem
                
                ObjEqptoDestino.InformaBanco frmgenMDI_SGAPlus.objBanco
                
                ObjEqptoDestino.Fase_ID = plFaseId
                
                'Tratar aqui eqpto tipo Modem
                bRetorno = ObjEqptoDestino.IncluirDados("", pobjAcesso.iD, objMensagem)
                
                If Not ObjEqptoDestino.objTipoAcesso Is Nothing Then
                    
'                        Set objEqptoModem = New clsGenEqptosAcessoModem
                    
                    Set objEqptoModem = ObjEqptoDestino.objTipoAcesso
                    
                    objEqptoModem.InformaBanco frmgenMDI_SGAPlus.objBanco
                    
                    objEqptoModem.EquipamentosAcessoID = pobjAcesso.iD
                    
                    bRetorno = objEqptoModem.IncluirDados(objMensagem)
                    
                    If Not bRetorno Then Exit For
                    
                End If
                
            Next
        
            If objMensagem.TipoMsg <> "" Then
               
'               GoTo Finalizar
            
            End If
            
            
        End If
        
        If objMensagem.TipoMsg <> "" Then
           
'           GoTo ErrorHandler
        
        End If
    
    End If

End Function
