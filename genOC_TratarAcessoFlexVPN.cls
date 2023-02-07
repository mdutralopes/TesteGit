VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsgenOC_TratarAcessoFlexVPN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public objDadosAcesso           As New clsGenAcesso 'ACESSO EMBRATEL



Public objOCSelecionarOTS       As clsgenOC_SelecionarOTS
Public vMensagemListaMsg        As Variant
Public bFlagControlaAlteracao   As Boolean
Public objMensagem             As New clsgenMensagem

Public objDadosPendencia        As New clsblf_Pendencia
Dim objDadosAcessoCliente       As New clsblf_TratarAcessoCliente

'merge
Dim objDadosAcessoFlexVPN       As New clsblf_TratarAcessoFlexVPN
Public objForm                  As frmgenTratarAcessoFlexVPN


Public vMaskProvedor            As Variant

Public objDadosFaseVariosCfs    As clsgenFase
Public objOcFaseVariosCfs       As clsInterOC_AlocarCF
Public objDadosVariosCFs        As clsBLF_AlocacaoCF
Public objAlocCFs               As clsGen_DetalheAlocarCF

    
Dim sTipoOperacao               As String
Dim sSigla_Acao                 As String
Dim sStatus_Fase                As String
Dim lServico_ID                 As Long
Dim lIdCircuito                 As Long
Dim lServico_ID_Anterior        As Long
Dim lFase_Id                    As Long
Dim lAcao_ID                    As Long
Dim lProduto_ID                 As Long
Dim lGrupo_ID                   As Long
Dim sDataFinalizacao            As String
Dim lSequencia_Acao             As Long
Dim sTextoAnotacao              As String
Dim lNrLocalInsta               As Long
Dim lContratoID                 As Long
Dim lCentroFuncionalCL_ID       As Long
Dim sPermiteExecucao            As String
Dim lCentroFuncionalID          As Long
Dim Circuito_ID                 As Long

Dim lNumeroOE                     As Long
Dim lAnoOE                        As Long
Dim lNumeroITemOE                 As Long

Dim sFamilia_Produto            As String
Dim oobjPendencia               As clsDNI_DadosPendencia

Public Property Get objPendencia() As clsDNI_DadosPendencia

Set objPendencia = oobjPendencia

End Property

Public Property Let objPendencia(ByVal vNewValue As clsDNI_DadosPendencia)

Set oobjPendencia = vNewValue

End Property

Public Property Get Familia_Produto() As String

    Familia_Produto = sFamilia_Produto

End Property

Public Property Let Familia_Produto(ByVal sNewValue As String)

    sFamilia_Produto = sNewValue
    
End Property

Public Property Get ContratoID() As Long

ContratoID = lContratoID

End Property

Public Property Let ContratoID(ByVal lNewValue As Long)

lContratoID = lNewValue

End Property

Public Property Get NrLocalInsta() As Long

NrLocalInsta = lNrLocalInsta

End Property

Public Property Let NrLocalInsta(ByVal lNewValue As Long)

lNrLocalInsta = lNewValue

End Property


Public Property Get NumeroOE() As Long

NumeroOE = lNumeroOE

End Property

Public Property Let NumeroOE(ByVal lNewValue As Long)

lNumeroOE = lNewValue

End Property

Public Property Get AnoOE() As Long

AnoOE = lAnoOE

End Property

Public Property Let AnoOE(ByVal lNewValue As Long)

lAnoOE = lNewValue

End Property

Public Property Get NumeroItemOE() As Long

NumeroItemOE = lNumeroITemOE

End Property

Public Property Let NumeroItemOE(ByVal lNewValue As Long)

lNumeroITemOE = lNewValue

End Property

Public Property Get TextoAnotacao() As String

TextoAnotacao = sTextoAnotacao

End Property

Public Property Let TextoAnotacao(ByVal sNewValue As String)

sTextoAnotacao = sNewValue

End Property

Public Property Get Sequencia_Acao() As Long

Sequencia_Acao = lSequencia_Acao

End Property

Public Property Let Sequencia_Acao(ByVal lNewValue As Long)

lSequencia_Acao = lNewValue

End Property

Public Property Get CentroFuncionalCL_ID() As Long

CentroFuncionalCL_ID = lCentroFuncionalCL_ID

End Property

Public Property Let CentroFuncionalCL_ID(ByVal lNewValue As Long)

lCentroFuncionalCL_ID = lNewValue

End Property

Public Property Get CentroFuncionalID() As Long

CentroFuncionalID = lCentroFuncionalID

End Property

Public Property Let CentroFuncionalID(ByVal lNewValue As Long)

lCentroFuncionalID = lNewValue

End Property

Public Property Get Produto_ID() As Long

Produto_ID = lProduto_ID

End Property

Public Property Let Produto_ID(ByVal lNewValue As Long)

lProduto_ID = lNewValue

End Property

Public Property Get Grupo_ID() As Long

Grupo_ID = lGrupo_ID

End Property

Public Property Let Grupo_ID(ByVal lNewValue As Long)

lGrupo_ID = lNewValue

End Property

Public Property Get Servico_ID() As Long

Servico_ID = lServico_ID

End Property

Public Property Let Servico_ID(ByVal vNewValue As Long)

lServico_ID = vNewValue

End Property

Public Property Get Servico_ID_Anterior() As Long

Servico_ID_Anterior = lServico_ID_Anterior

End Property

Public Property Let Servico_ID_Anterior(ByVal vNewValue As Long)

lServico_ID_Anterior = vNewValue

End Property

Public Property Get Fase_ID() As Long

Fase_ID = lFase_Id

End Property

Public Property Let Fase_ID(ByVal vNewValue As Long)

lFase_Id = vNewValue

End Property

Public Function CarregarForm() As Boolean
    
    Dim objdadosServico        As New clsGenServico
    Dim objAutenticacao        As New clsgenAutenticacao
    Dim objDadosFase           As New clsgenFase
    Dim objDadosProduto        As New clsGenProduto
    Dim objPendenciaDNI        As New clsDNI_DadosPendencia
    Dim objDniDadosTecnico     As New clsDNI_DadosTecnicos
    Dim paPermissoes()         As Variant
    Dim sSituacaoOTS           As String
    Dim numUltimaFaseRealizada As Long
    Dim IdOTS                  As Long
    Dim lFaseAcessoID          As Long
    Dim cColecao               As Collection
    Dim flUsuarioDoCF          As Boolean
    Dim objOC_AcessoEBT        As New clsGenOC_TratarAcessoEmbratel
    Dim objDadosAcessoSOE      As clsblfDadosAcesso
    
    'Não resgata os dados da OTS anterior quando se altera o
    'endereço de instalação. IP VPN
    Dim sCNLAtual              As String
    Dim sCNLAnterior           As String
    Dim objDadosCircuito       As New clsgenCircuito
    Dim bResgataAnterior       As Boolean
    
    Dim sTabela_Interface       As String   'JAH
    
    On Error GoTo ErrorHandler
    
    CarregarForm = False
    
    'Circuito
    Set objDadosCircuito = New clsgenCircuito
    
    objDadosCircuito.InformaMensagem objMensagem
    objDadosCircuito.InformaBanco frmgenMDI_SGAPlus.objBanco

    objDadosAcesso.InformaMensagem objMensagem
    objDadosAcesso.InformaBanco frmgenMDI_SGAPlus.objBanco
    
    'Verfica se a Fase/OTS está em OK para execução
    '=================================================================================
    objdadosServico.InformaBanco frmgenMDI_SGAPlus.objBanco
    objdadosServico.InformaMensagem objMensagem
    
    'Obtem dados servico e verifica se a OTS selecionada ainda encontra-se em andamento
    bRetorno = objdadosServico.RecuperarDados(Servico_ID)
    
  
    
    If objMensagem.TipoMsg <> "" Then
       objMensagem.ExibeMensagem
       Set objdadosServico = Nothing
       
       GoTo Finalizar
    End If
    
    sSituacaoOTS = objdadosServico.StatusServico
    
    If sSituacaoOTS = "CAN" Or sSituacaoOTS = "CR" Then
       
       objMensagem.TipoMsg = "A"
       objMensagem.Descricao = "Esta OTS foi paralisada pelo SOE!"
       objMensagem.ExibeMensagem
       
       Set objdadosServico = Nothing
       
       GoTo Finalizar
    End If
             
    'Verifica se a Fase ainda continua Em Andamento
    objDadosFase.InformaBanco frmgenMDI_SGAPlus.objBanco
    objDadosFase.InformaMensagem objMensagem
    
    If Not objDadosFase.RecuperarSituacaoFase(Fase_ID) And DataFinalizacao = "" Then

       objMensagem.ExibeMensagem

       GoTo Finalizar
    Else
       
       objMensagem.TipoMsg = ""
       objMensagem.Descricao = ""
    
    End If
             
    objPendenciaDNI.InformaBanco frmgenMDI_SGAPlus.objBanco
    objPendenciaDNI.InformaMensagem objMensagem
             
    objDadosPendencia.InformaBanco frmgenMDI_SGAPlus.objBanco
    objDadosPendencia.InformaMensagem objMensagem
             
    'Obtem a Sequencia_Acao da última fase realizada
    numUltimaFaseRealizada = objDadosFase.ObtemSequenciaUltimaFaseRealizada(Servico_ID, _
                                             GbCodigoFaseProgramarCPE, Familia_Produto)
    
    If objMensagem.TipoMsg <> "" Then
       
       objMensagem.ExibeMensagem
       
       Set objdadosServico = Nothing
       
       GoTo Finalizar
    End If
    
    'Verifica as permissoes do usuario
    objAutenticacao.IDUsuario = frmgenMDI_SGAPlus.objAutenticacao.IDUsuario
    objAutenticacao.Login = frmgenMDI_SGAPlus.objAutenticacao.Login
    objAutenticacao.InformaMensagem objMensagem
    
    Call objAutenticacao.InformaBanco(frmgenMDI_SGAPlus.objBanco)
    
    'GbCodigoFaseAcesso - ID da fase Acesso cadastrado na tabela Tipo fase
    Call objAutenticacao.RetornarPermissoesFase(objAutenticacao.IDUsuario, _
                                                paPermissoes, _
                                                GbCodigoFaseAcesso, _
                                                Produto_ID, _
                                                Acao_ID, _
                                                CentroFuncionalID)
    'Coluna 3 - Checa Permissao de consulta
    'Coluna 4 - Checa Permissao de Alteracao
    'Coluna 5 - Checa Permissao de Exclusao
    
    flUsuarioDoCF = objAutenticacao.VerificaUsuarioNoCF( _
                                                        objAutenticacao.IDUsuario, _
                                                        CentroFuncionalID)
    Set objAutenticacao = Nothing

    If UBound(paPermissoes) = 0 Then
        objMensagem.TipoMsg = "E"
        objMensagem.Descricao = "O usuário não possui permissão de consulta!"
        objMensagem.ExibeMensagem

        GoTo Finalizar

    End If
   
    If (Produto_ID = GbIDProdutoDiginet Or Produto_ID = GbIDProdutoEmbratelIPL) And _
       objdadosServico.DataRealizacao = "" Then
            
        bRetorno = VerificaUltimaVersaoContratoDNI(Servico_ID, objdadosServico.VersaoContrato, objMensagem)
        
        If Not bRetorno Then
        
            GoTo Finalizar
            
        End If
        
    End If
   
   '======================================================================================
   'Instancia o formulario, obtem os dados da fase e carrega os campos da tela
   '======================================================================================
    Set objForm = New frmgenTratarAcessoFlexVPN
    
     'recupera os dados Tecnicos pra preencher o Form
    Set objDadosTecnicos = New clsbldDadosTecnicos
    objDadosTecnicos.InformaBanco frmgenMDI_SGAPlus.objBanco
    objDadosTecnicos.InformaMensagem objMensagem
    bRetorno = objDadosTecnicos.RecuperarDados(objdadosServico.NumOE, _
    objdadosServico.AnoOE, _
     1, _
    objdadosServico.SubItemOE)
    
    
    'Trata erro
    If (objMensagem.TipoMsg <> "") Then
    
        GoTo ErrorHandler
    
    End If
   'Call CarregaCabecalho_BLF(objForm.UsrBLFCabecalho_Acesso1, objBlfCabecalho)
    
    'TBASTOS FLEX VPN
    'objForm.UsrGenDadosFlexVPN1.VRFFLX = objDadosTecnicos.VRFFLX
    'objForm.UsrGenDadosFlexVPN1.IPVRFFLX = objDadosTecnicos.IPVRFFLX
    'objForm.UsrGenDadosFlexVPN1.IPLNSFLX = objDadosTecnicos.IPLNSFLX
    'objForm.UsrGenDadosFlexVPN1.BlocoFLX = objDadosTecnicos.BlocoFLX
    'objForm.UsrGenDadosFlexVPN1.IPCPEFLX = objDadosTecnicos.IPCPEFLX
    'objForm.UsrGenDadosFlexVPN1.IPTTFLX = objDadosTecnicos.IPTTFLX
    
    objForm.txtVRF.Text = objDadosTecnicos.VRFFLX
    objForm.txtIpLoopbackLnsVrfCliente.Text = objDadosTecnicos.IPVRFFLX
    'objForm.TxtIPPublicoLoopbackLns = objDadosTecnicos.BlocoFLX
    'objForm.txtIpLoopbackClienteTT = objDadosTecnicos.IPTTFLX
    objForm.TxtBlocoIpLanCliente.Text = objDadosTecnicos.IPCPEFLX
    objForm.txtIpLoopbackClienteCpe = objDadosTecnicos.IPLOOPBACKCLICPE
    
    
    
    
    
    
    objForm.UsrGenDadosFlexVPN1.MostrarDados
   
    If objMensagem.TipoMsg <> "" Then
       
       GoTo ErrorHandler
       
    End If
    
    objForm.LblAcao_Status = Status_Fase
    
    objDadosProduto.InformaBanco frmgenMDI_SGAPlus.objBanco
    objDadosProduto.InformaMensagem objMensagem
    
    bRetorno = objDadosProduto.RecuperarDados(Produto_ID)
    
    objForm.lblProduto.Caption = objDadosProduto.Descricao
    objForm.Caption = objDadosProduto.Sigla & " - Tratar Acesso Flex VPN"
    
    'Carrega os Cabecalho
    If Familia_Produto = "I" Then
        
       'objForm.UsrDNICabecalhoAcesso.Visible = False
       
       If (Produto_ID = GbIDProdutoBIV) Then
       
            'IP VPN
            objForm.CmdConsultarOE.Left = 2895
            
'            objForm.UsrBLFCabecalho_Acesso.Visible = False
            
            'objForm.usrBIPVPNCabecalho.Visible = True
        
            'Call CarregaCabecalho_BIV(objForm.usrBIPVPNCabecalho, objBlfCabecalho)
       Else
            '''objForm.CmdConsultarOE.Visible = False
            If (objForm.CmdTrataPendencia.Visible = False) Then
            
                objForm.CmdConsultarOE.Left = 2895
            
            End If
            
            Call CarregaCabecalho_BLF(objForm.UsrBLFCabecalho_Acesso1, objBlfCabecalho)
       
       End If
    
       'Carregar a combo de Velocidade
       objOC_AcessoEBT.Produto_ID = Produto_ID
       

       
       
        '=================================================================
        'CALLBAT - 03/08/2013 - Resolução 600
        'SO-TELCO-79171SQV-04
        '=================================================================
        bRetorno = objdadosServico.RecuperaIdentificacaoSNOA(Servico_ID)
        
        If objMensagem.TipoMsg <> "" Then
           objMensagem.ExibeMensagem
           
           Set objForm = Nothing
           GoTo Finalizar
        End If
        
'        If objdadosServico.Indic_PMS = "S" Then
'            objForm.chkInterligacaoPMS.Value = 1
'        Else
'            objForm.chkInterligacaoPMS.Value = 0
'        End If
'
  '      objForm.lblNumPedidoSNOA.Text = objdadosServico.Num_Ped_Venda_SNOA
        
        '=================================================================
        '=================================================================
       
        'CALBAT - 05/10/2007 =====================================
        'Tratar Acesso Paralelismo.                                                          =
        '=========================================================
       ' bRetorno = CarregarCombos(objDadosAcesso, "INTERFACECLI", objForm.CboPadraoInterfaceEBT)
        
        If objMensagem.TipoMsg <> "" Then
        
            objMensagem.ExibeMensagem
            
            Unload objForm
            Set objForm = Nothing
            
            GoTo Finalizar
        
        End If
        '=========================================================
        '=========================================================
        
    ElseIf Familia_Produto = "D" Then
        
        objForm.cmdPendencia.Visible = False
        
        'objForm.UsrDNICabecalhoAcesso.Visible = True
        
        'objForm.UsrBLFCabecalho_Acesso.Visible = False
        
        objForm.CmdTrataPendencia.Left = 2190
        
        objForm.CmdTrataPendencia.Visible = True
        objForm.CmdConsultarOE.Visible = True
        objForm.CmdConsultarOE.Left = 4095
       
       'Recupera dados do Cabecalho
'        Call CarregaCabecalho_Diginet(objForm.UsrDNICabecalhoAcesso, _
'                                  Servico_ID, Fase_ID, objMensagem, GbCodigoFaseAcesso)
    
        'Carregar a combo de Velocidade
        objOC_AcessoEBT.Produto_ID = Produto_ID
        
        '''bRetorno = objOC_AcessoEBT.CarregarComboVelocidade_DNI(objDadosAcesso, _
                                           objForm.UsrDNICabecalhoAcesso.Velocidade, _
                                           "VELOCIDADE_DNI", objForm.CboVelocidadeAcesso)
    
        Set objDniDadosTecnico = New clsDNI_DadosTecnicos
        
        objDniDadosTecnico.InformaMensagem objMensagem
        objDniDadosTecnico.InformaBanco frmgenMDI_SGAPlus.objBanco
        
        bRetorno = objDniDadosTecnico.RecuperarDados(NumeroOE, AnoOE, NumeroItemOE, _
                                                     ContratoID)
    
        If bRetorno Then
           
           'Carrega a combo de Provedor Urbano
'            bRetorno = objOC_AcessoEBT.CarregarCombos(objDadosAcesso, "PROVEDOR", _
'                                                      objForm.CboProvedorUrbano)
        
            If objMensagem.TipoMsg <> "" Then
        
               objMensagem.ExibeMensagem
        
               Unload objForm
               Set objForm = Nothing
               
               GoTo Finalizar
        
            End If
            
            'Número do Acesso provedor
       '     objForm.SpdNumeroAcessoUrbano.Text = objDniDadosTecnico.NumeroAcessoProvedor
            
            'Descricao do equipamento
          '  objForm.txtDescrEqptoCliente.Text = objDniDadosTecnico.DescricaoEqpto
            
            'Cabeacao
            'objForm.txtCabeamentoInternoCliente.Text = objDniDadosTecnico.CabeacaoInterna
            
            'Velocidade
            'objForm.txtCabeamentoInternoCliente.Text = objDniDadosTecnico.CabeacaoInterna
            
            'Carregar a combo de Velocidade
'            bRetorno = objOC_AcessoEBT.CarregarComboVelocidade_DNI(objDadosAcesso, _
'                                           objDniDadosTecnico.Cd_Veloc, _
'                                           "VELOCIDADE_DNI", objForm.CboVelocidadeAcesso)
                                           
           ' objForm.CboVelocidadeAcesso.Enabled = False
            
        End If
    
    End If
    
    'Verifica se existe alguma fase posterior a esta já iniciada, em caso afirmativo,
    'desabilitar os campos da tela e os botões de ações ou caso o usuário não possua
    'permissão de atualização
    
    If Familia_Produto = "D" Then
    
    'CALBAT - 21/09/2007
    '====================================================
        'If (numUltimaFaseRealizada > Sequencia_Acao) Or _
            (paPermissoes(4, 1) = 0) Or (Not flUsuarioDoCF) Then
        
        If (lProduto_ID = GbIDProdutoVBL) Or _
            (paPermissoes(4, 1) = 0) Or (Not flUsuarioDoCF) Then
    '====================================================
           
           Call DesabilitaObjetosTela
           
           objForm.CmdTrataPendencia.Enabled = False
           
        End If

    Else
    
        If (DataFinalizacao <> "" And PermiteExecucao <> "S") Or _
            paPermissoes(4, 1) = 0 Or Not flUsuarioDoCF Then
        
           Call DesabilitaObjetosTela
            
        End If
    
    End If
    
    If Sigla_Acao = "DES" Or Sigla_Acao = "CAN" Then
            
        Call DesabilitaObjetosTela
    
        objForm.CmdSalvar.Enabled = True
        objForm.CmdLimpar.Enabled = False
            
    End If
    
    'Carrega a combo de Provedor
    'bRetorno = objOC_AcessoEBT.CarregarCombos(objDadosAcesso, "PROVEDOR", objForm.CboProvedorUrbano)

    If objMensagem.TipoMsg <> "" Then

       objMensagem.ExibeMensagem

       Unload objForm
       Set objForm = Nothing
       
       GoTo Finalizar

    End If
    
    vMaskProvedor = objOC_AcessoEBT.vMaskProvedor
    
    
    objDadosAcessoFlexVPN.InformaBanco frmgenMDI_SGAPlus.objBanco
    objDadosAcessoFlexVPN.InformaMensagem objMensagem
    
    '#BLD
    If (Produto_ID = GbIDProdutoBLF Or _
            Produto_ID = GbIDProdutoBIV Or _
            Produto_ID = GbIDProdutoJAH Or _
            Produto_ID = GbIDProdutoBLD Or _
            Produto_ID = GbIDProdutoVOIP Or Produto_ID = GbIDProdutoVBL) Then
        
        Set objDadosAcessoSOE = New clsblfDadosAcesso
    
        objDadosAcessoSOE.InformaBanco frmgenMDI_SGAPlus.objBanco
        objDadosAcessoSOE.InformaMensagem objMensagem
        
        'Recupera Dados de Acesso
        sTabela_Interface = ""
        
        If (Produto_ID = GbIDProdutoJAH) Then
    
            sTabela_Interface = "ATEN_ACESS_FACIL"
        
        End If
        
        bRetorno = objDadosAcessoSOE.RecuperarDados(ContratoID, _
                        NrLocalInsta, _
                        , _
                        sTabela_Interface)
        
        If (objMensagem.TipoMsg <> "") Then
        
            GoTo ErrorHandler
            
        End If
        
        objDadosAcesso.Cd_Veloc = objDadosAcessoSOE.CdVelocidadeAcesso
    
    End If
    
    'Caso exista Servico Anterior e a fase ainda não foi iniciada,
    'Obtem os dados da OTS anterior
    
'    If (Servico_ID_Anterior <> 0) Then
'        'Não resgata os dados da OTS anterior quando se altera o
'        'endereço de instalação. IP VPN
'
'        bResgataAnterior = True
'
'        'Local de instalação da nova OTS
'        If (Produto_ID = GbIDProdutoBIV) And _
'            (objdadosServico.Id_Acao = GbIDAcaoAlteracao Or _
'                objdadosServico.Id_Acao = GbIDAcaoCorrecaoAtivacao Or _
'                objdadosServico.Id_Acao = GbIDAcaoCorrecaoAlteracao) Then
'
'            'OTS atual
'            bRetorno = objdadosCircuito.RecuperarDados(Servico_ID)
'
'            If (Not bRetorno) Then
'               objMensagem.ExibeMensagem
'
'               Unload objForm
'               Set objForm = Nothing
'
'               GoTo Finalizar
'            End If
'
'            'OTS atual
'            sCNLAtual = objdadosCircuito.CNL_Ponta_A
'
'
'            'OTS anterior
'            bRetorno = objdadosCircuito.RecuperarDados(Servico_ID_Anterior)
'
'            If (Not bRetorno) Then
'               objMensagem.ExibeMensagem
'
'               Unload objForm
'               Set objForm = Nothing
'
'               GoTo Finalizar
'            End If
'
'            'OTS anterior
'            sCNLAnterior = objdadosCircuito.CNL_Ponta_A
'
'            If (sCNLAtual <> sCNLAnterior) Then
'
'                bResgataAnterior = False
'
'            End If
'
'        End If
'
''       If (Not objDadosAcessoCliente.FaseExistente(Fase_ID) And _
''                bResgataAnterior) Then
''
''           IdOTS = Servico_ID_Anterior
''
''           'Problema identificado durante os testes do BLD
''           If (Sigla_Acao = "DES" Or Sigla_Acao = "CAN") Then
''
''                'GbCodigoFaseAcesso - ID da fase na tabela Tipo_Fase
''                IDFase = objdadosServico.ObtemIDFase_Recursivo(Servico_ID_Anterior, _
''                            GbCodigoFaseAcesso)
''           Else
''                'GbCodigoFaseAcesso - ID da fase na tabela Tipo_Fase
''                IDFase = objdadosServico.ObtemIDFase(Servico_ID_Anterior, _
''                        GbCodigoFaseAcesso)
''           End If
''
''           If objMensagem.TipoMsg <> "" Then
''              GoTo ErrorHandler
''           End If
''
''       Else
''           IdOTS = Servico_ID
''           IDFase = Fase_ID
''       End If
'
'    Else
'
'       IdOTS = Servico_ID
'       IDFase = Fase_ID
'
'    End If
        
    'Recupera os dados de Acesso Cliente e popula os objetos da tela
    If Not objDadosAcessoFlexVPN.RecuperarDados(IDFase) Then
    
        If objMensagem.TipoMsg <> "" Then
           
           objMensagem.ExibeMensagem
           
           Set objForm = Nothing
           
           GoTo Finalizar
        
        End If
        
        Set cColecao = New Collection
        
        objDadosPendencia.ColecaoPendencia = cColecao
        
        objPendencia = objPendenciaDNI
        
        'objForm.FlControlaAlteracao = False
        
        TipoOperacao = "I"     'Indica que a fase está sendo executada pela 1a vez
        
        objForm.informa_OC Me
        
'        If Produto_ID = GbIDProdutoVBL Then
'            Call SelecionaCombo(objForm.CboVelocidadeAcesso, _
'                                CLng("10093"))
'        End If
        objForm.Show
        
        GoTo Finalizar
    
    End If
        
    'Caso a fase não tenha sido iniciada e seja uma OTS de Desativacao ou Cancelamento,
    'Exibe os dados da OTS original
    
    TipoOperacao = "A"     'Indica que a fase já foi iniciada anteriormente
       
    objForm.informa_OC Me
    
    'Popula os campos da tela
    Call PopularCamposTela
    
    
    'Caso exista pendencia para a fase obtem os dados de pendencia e carrega a
    'propriedade do formulário com o objeto Pendencia
    
    If objForm.cmdPendencia.Enabled = True And Familia_Produto = "I" Then
           
        If Not objDadosPendencia.RecuperarDados(Fase_ID) Then
    
           If objMensagem.TipoMsg <> "" Then
              
              objMensagem.ExibeMensagem
    
              Set objForm = Nothing
    
              GoTo Finalizar
    
           End If
           
           'Caso não exista pendencia, cria uma colecao  vazia no objeto pendencia
           Set cColecao = New Collection
           
           objDadosPendencia.ColecaoPendencia = cColecao
    
        End If
    
    End If
    
    If Familia_Produto = "D" Then
    
       If Not objPendenciaDNI.RecuperarDados(Fase_ID) Then
            
          If objMensagem.TipoMsg <> "" Then
             
             objMensagem.ExibeMensagem
             
             GoTo Finalizar
             
          End If
       
       End If
    
        If objPendenciaDNI.iD <> 0 Then
           
           objForm.CmdTrataPendencia.Caption = "&Retirar Pendência"
        
        End If
        
        objForm.CmdTrataPendencia.Visible = True
        
        objPendencia = objPendenciaDNI
        
    End If
    
    'Caso o Servico já esteja Finalizado, exibir os campos desabilitados
    If objdadosServico.DataRealizacao <> "" And Familia_Produto = "I" Then
       
       Call DesabilitaObjetosTela

    End If
    
    'objForm.FlControlaAlteracao = False
    
    objForm.Show
    
    'Executa método para popular os campos do Controle Cabecalho
    CarregarForm = True

Finalizar:
    
    Set objdadosServico = Nothing
    Set objAutenticacao = Nothing
    Set objDadosFase = Nothing
    Set objDadosProduto = Nothing
    Set objPendenciaDNI = Nothing
    Set objDadosCircuito = Nothing
    
    Exit Function
    
ErrorHandler:

    If Err.Number <> 0 Then
       objMensagem.TipoMsg = "E"
       objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                   "Houve um erro de execução no Método: " & vbCrLf & vbCrLf & "CarregarForm " & _
                   "da classe clsblfOC_TratarAcessoCliente"
    End If
       
    objMensagem.ExibeMensagem
       
    GoTo Finalizar
       
End Function
Public Function CarregarCombos(objDadosAcesso As clsGenAcesso, _
                               RefCombo As String, _
                               cboLista As ComboBox) As Boolean

On Error GoTo ErrorHandler

    Dim bRetorno As Boolean
    Dim iCont As Integer
    Dim aLista() As Variant
    
    CarregarCombos = False
    
    ' Limpa a combo
    bRetorno = objDadosAcesso.MontarListasAcesso(RefCombo, aLista)
    
    If Not bRetorno Then

        GoTo ErrorHandler

    End If

    'Ler o array para preencher a combo
    
    cboLista.Clear
    cboLista.AddItem ""
    
    'Ao carregar a combo de provedores, carrega as informacoes de máscara de Digitacao
    'da designacao do Provedor
    If RefCombo = "PROVEDOR" Then
    
        ReDim vMaskProvedor(UBound(aLista, 2))
        
    End If
    
    For iCont = 1 To UBound(aLista, 2)

        If Not IsEmpty(aLista(0, iCont)) Then

            If RefCombo = "PROVEDOR" Then
            
                cboLista.AddItem Trim(aLista(1, iCont)) & _
                  Space(25 - Len(Trim(aLista(1, iCont)))) & Trim(aLista(4, iCont))
                
            Else
            
                cboLista.AddItem aLista(1, iCont)
            
            End If
            
            cboLista.ItemData(iCont) = aLista(0, iCont)
            
            If RefCombo = "PROVEDOR" Then
            
               If Trim(aLista(3, iCont)) <> "" Then
               
                   vMaskProvedor(iCont) = Trim(aLista(3, iCont))
                   
               Else
               
                   vMaskProvedor(iCont) = Trim(aLista(2, iCont))
               
               End If
               
            End If
        
        Else
        
            Exit For
            
        End If

    Next
        
    CarregarCombos = True
       
Finalizar:
    
    On Error GoTo 0
    
    Exit Function
        
ErrorHandler:

    If objMensagem.TipoMsg = "" Then
       
       objMensagem.TipoMsg = "E"
       objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                   "Houve um erro de execução no Método: " & vbCrLf & vbCrLf & "CarregarCombos " & _
                   "da classe clsGenOC_TratarAcessoEmbratel"
    
    End If
    
    GoTo Finalizar

End Function

Public Function GravarDados(pbFlagDescarregaFormulario As Boolean, _
                            Optional Flag_Atualiza_Pendencia As Boolean) As Boolean

    Dim objdadosServico             As New clsGenServico
    Dim objDadosFase                As New clsgenFase
    Dim objWEBIP                    As clsInterOC_TrataWebIP
    Dim objBldOrderEntry            As clsbldOrderEntry
    
    Dim objDadosAtivacaoFinal       As clsblf_AtivacaoFinalizacao
    
    'gerar servico de desativação
    Dim sDesignacaoAtual            As String
    Dim sDesignacaoAnterior         As String
    
    Dim sSituacaoOTS                As String
    Dim bRetorno                    As Boolean
    Dim FlagIniciouTransacao        As Boolean
    Dim numUltimaFaseRealizada      As Long
    Dim FlagTemPendenciaEmAberto    As Boolean
    Dim FlagFaseFinalizadaOK        As Boolean
    Dim lSeqUltimaFaseOTS           As Long
    Dim sDesignacaoSOE              As String
    
    Dim FlagTemPendenciaCliente     As Boolean
    Dim FlagManterPeriodo           As Boolean
    
    Dim lIdFaseCentroLocal          As Long
    Dim lIDCentroFuncionalCL        As Long
    Dim sCodigoCFCentroLocal        As String
    Dim lID_Fase_Seguinte           As Long
    
    'VAGNER
    Dim iSeqAux                     As Integer
    
    'Checa se o produto pode alocar vários CFs para a fase Tratar Centro Local
    Dim objDadosDuracaoFase         As clsgenDuracaoFase
    Dim ocAcessoEmbratel            As clsGenOC_TratarAcessoEmbratel
    Dim ocAcessoFlexVPN             As clsgenOC_TratarAcessoFlexVPN
    
    
    On Error GoTo ErrorHandler
    
    FlagIniciouTransacao = False
    FlagTemPendenciaCliente = False
    
    GravarDados = False
    
    
    If Acao_ID <> 8 And Acao_ID <> 9 Then   'Desativacao ou Cancelamento
    
        If Not CamposPreenchidosOK Then
        
           GoTo Finalizar
        
        End If
    
    End If
    'Verfica se a OTS selecionada ainda encontra-se em andamento
    objdadosServico.InformaBanco frmgenMDI_SGAPlus.objBanco
    objdadosServico.InformaMensagem objMensagem
    
    'Obtem dados servico e verifica se a OTS selecionada ainda encontra-se em andamento
    bRetorno = objdadosServico.RecuperarDados(Servico_ID)
    
    If Trim(objMensagem.TipoMsg) <> "" Then
       objMensagem.ExibeMensagem
       Set objdadosServico = Nothing
       
       GoTo Finalizar
    End If
    
    sSituacaoOTS = objdadosServico.StatusServico
    
    If (sSituacaoOTS = "CAN") Or _
            (sSituacaoOTS = "CR") Then
       
       objMensagem.TipoMsg = "A"
       objMensagem.Descricao = "Esta OTS foi paralisada pelo SOE!"
       objMensagem.ExibeMensagem
       
       Set objdadosServico = Nothing
       
       GoTo Finalizar
    End If

    objDadosFase.InformaBanco frmgenMDI_SGAPlus.objBanco
    objDadosFase.InformaMensagem objMensagem

    If (Not objDadosFase.RecuperarSituacaoFase(Fase_ID)) And _
        (DataFinalizacao = "") Then

       Set objDadosFase = Nothing

       objMensagem.ExibeMensagem

       GravarDados = True

       GoTo Finalizar
    
    Else
       
       objMensagem.TipoMsg = ""
    
    End If

    DataFinalizacao = objDadosFase.ObtemDataFimReal(Fase_ID)

    'Obtem o número da última fase realizada
    numUltimaFaseRealizada = objDadosFase.ObtemSequenciaUltimaFaseRealizada(Servico_ID, _
                                          GbCodigoFaseProgramarCPE, Familia_Produto)
    
    If (objMensagem.TipoMsg <> "") Then
       
       objMensagem.ExibeMensagem
       GoTo Finalizar
    
    End If
    
    If (numUltimaFaseRealizada > Sequencia_Acao And _
            PermiteExecucao <> "S") And lProduto_ID = GbIDProdutoVBL Then
    '====================================================
    
       objMensagem.TipoMsg = "A"
       objMensagem.Descricao = "Os dados desta fase não poderão ser alterados! Pois, a fase seguinte já foi iniciada."
       objMensagem.ExibeMensagem
       
       GravarDados = True
       
       GoTo Finalizar
    
    End If

    If (Produto_ID = GbIDProdutoDiginet Or _
            Produto_ID = GbIDProdutoEmbratelIPL) And _
       (objdadosServico.DataRealizacao = "") Then
            
        bRetorno = VerificaUltimaVersaoContratoDNI(Servico_ID, objdadosServico.VersaoContrato, objMensagem)
        
        If (Not bRetorno) Then
            GoTo Finalizar
        End If
        
    End If

    'Checa se o produto pode alocar vários CFs para a fase Tratar Centro Local
    Set objDadosDuracaoFase = New clsgenDuracaoFase
    
    objDadosDuracaoFase.InformaBanco frmgenMDI_SGAPlus.objBanco
    objDadosDuracaoFase.InformaMensagem objMensagem
    
    'Carrega as propriedades da classe de dados e executa o método de Inclusao ou alteracao
    ' Call CarregarAtributosAcesso
    
    bRetorno = frmgenMDI_SGAPlus.objBanco.Iniciar_Transacao()
    
      objDadosFase.Usuario_ID = GbID_Usuario
         
    If (Not bRetorno) Then
        
        GoTo ErrorHandler
    
    End If
    
    FlagIniciouTransacao = True
    
    'Atualiza dados da fase Acesso Cliente
    If (objDadosAcessoFlexVPN.FaseExistente(Fase_ID)) Then
       
       Gravar = objDadosAcessoFlexVPN.AtualizarDados(Fase_ID)
    
    Else
       
       Gravar = objDadosAcessoFlexVPN.IncluirDados(Fase_ID)
    
    End If
   
    If (objMensagem.TipoMsg <> "") Then
       
       GoTo ErrorHandler
       
    End If
    
    If (Familia_Produto = "I") Then
        
        'Verifica se Existe pendencia a ser incluida
        If (objDadosPendencia.ColecaoPendencia.Count <> 0) Then
           
           bRetorno = objDadosPendencia.AtualizarDados(Fase_ID, Servico_ID, _
                                                       Produto_ID, _
                                                       FlagTemPendenciaEmAberto, _
                                                       objdadosServico.DiretorRegional, _
                                                       objdadosServico.NumOE, _
                                                       objdadosServico.AnoOE, _
                                                       objdadosServico.ItemOE)
           
           If (Not bRetorno) Then
              
              GoTo ErrorHandler
              
           End If
        
        End If
    
    ElseIf Familia_Produto = "D" Then
    
        'Verifica se Existe pendencia a ser incluida
        If (objPendencia.Tipo_Pendencia_ID <> 0) And _
                (Flag_Atualiza_Pendencia) Then
            
           objPendencia.Fase_ID = Fase_ID
            
           objPendencia.Servico_ID = Servico_ID
           
           objPendencia.Codigo_Tipo_Fase = GbCodigoFaseAcesso
           '---------------------------------------------------
           objPendencia.InformaBanco frmgenMDI_SGAPlus.objBanco
           objPendencia.InformaMensagem objMensagem
        
           If (objPendencia.Operacao = "R") Then
               
              bRetorno = objPendencia.Retirar(objPendencia.iD)
              
              FlagTemPendenciaEmAberto = False
            
           Else
               
              objPendencia.Servico_ID = Servico_ID
              
              bRetorno = objPendencia.IncluirDados
              
              bRetorno = objPendencia.RecuperarDados(Fase_ID)
              
              FlagTemPendenciaEmAberto = True
            
           End If
            
           If (Not bRetorno) Then
              
              GoTo ErrorHandler
              
           End If
            
        Else
            
            If (objPendencia.Tipo_Pendencia_ID <> 0) Then
            
                FlagTemPendenciaEmAberto = True
                
            End If
            
        End If
    
    End If
    
    'Grava Anotacao da fase
    If Trim(TextoAnotacao) <> "" Then
       
       bRetorno = objOCSelecionarOTS.GravarAnotacaoGen(False)
         
       If objMensagem.TipoMsg <> "" Then
            
          GoTo ErrorHandler
      
       End If
    
    End If
    objDadosFase.Usuario_ID = GbID_Usuario

    bRetorno = objDadosFase.IniciaFase(Fase_ID, GbID_Usuario)
    
    If (objMensagem.TipoMsg <> "") Then
       
       GoTo ErrorHandler
       
    End If
    
    FlagFaseFinalizadaOK = False
    
    
    If VerificaFinalizacaoOK(Sigla_Acao) And _
            (Not FlagTemPendenciaEmAberto) Then
       
       'Caso a fase não tenha sido finalizada e não exista pendencia em Aberto,
       'Finaliza a fase
        '''
        'gerar servico de desativação
        'inicio
        
        If (Produto_ID = GbIDProdutoBIV) And _
            (objdadosServico.ID_ACAO = GbIDAcaoAlteracao) Then

            'Recupera os dados de Circuito da OTS que está alterando
            bRetorno = objdadosServico.DevolveDesignacaoOTS(objdadosServico.ID_Servico, _
                                    Produto_ID, _
                                    True)
            If (Not bRetorno) Then
                GoTo ErrorHandler
            End If
            
            sDesignacaoAtual = ("" & objdadosServico.DesignacaoCircuito)
            
            'Recupera os dados de Circuito da OTS que está sendo alterada
            bRetorno = objdadosServico.DevolveDesignacaoOTS(objdadosServico.ID_Servico_Anterior, _
                                    Produto_ID, _
                                    True)
            If (Not bRetorno) Then
                GoTo ErrorHandler
            End If
            
            sDesignacaoAnterior = ("" & objdadosServico.DesignacaoCircuito)

            'só irá verificar se existe necessidade de gerar serviço de desativação, se as designações forem diferentes
            If (sDesignacaoAtual <> sDesignacaoAnterior) Then
                        
                objMensagem.TipoMsg = "Q"
                objMensagem.Descricao = "Deseja Gerar uma OTS Desativando " & _
                                        "a OTS alterada"
            
                If (objMensagem.ExibeMensagem = "S") Then

                    objdadosServico.geraDesativacao = "S"
            
                ElseIf (objMensagem.ExibeMensagem = "N") Then
                   
                    objdadosServico.geraDesativacao = "N"
            
                End If
        
                bRetorno = objdadosServico.AtualizarGeraDesativacao(objdadosServico.geraDesativacao)
                
                If (objMensagem.TipoMsg <> "") Then
                    GoTo ErrorHandler
                End If
                
            End If
            
        End If
        'fim

       If (DataFinalizacao = "") Then
          'Para OTS de Cancelamento e Desativacao, não será necessário calcular o periodo
          'para a fase seguinte já que todas as fases após a fase de Alocacao, são executadas
          'em paralelo
          
          If (Sigla_Acao <> "CAN" And Sigla_Acao <> "DES") And _
                (Familia_Produto = "I") Then
            
              'Calcula o período para a proxima fase
              bRetorno = objDadosFase.CalculaPeriodoProximaFase(Sequencia_Acao, _
                            Servico_ID)
        
              If (objMensagem.TipoMsg <> "") Then
                 GoTo ErrorHandler
              End If

              'Uma OTS de cancelamento pode não gerar Fase Tratar Centro Local
              'Aquino IP VPN
              lIdFaseCentroLocal = objDadosFase.ObtemIDFase(Servico_ID, _
                            GbCodigoFaseCentroLocal)
            
              If (objMensagem.TipoMsg <> "") Then
                GoTo ErrorHandler
              End If
            
              'Checa se o produto pode alocar vários CFs para a fase Tratar Centro Local
'              If (objDadosDuracaoFase.ChecaPermissaoMultiplosCF(Produto_ID, _
'                                                       GbCodigoFaseCentroLocal)) Then
'
'                If objDadosFase.ObtemSequenciaFaseFinalizada(Servico_ID, _
'                                                   GbCodigoFaseCentroLocal) = 0 And _
'                                                   (Familia_Produto = "I" And _
'                                                   lIdFaseCentroLocal > 0) Then
'
'                  'Carrega a tela Alocação de Centro Funcional e atualiza os CFs alocados
'                  Set ocAcessoEmbratel = New clsGenOC_TratarAcessoEmbratel
'
'                  'Informa o ID Produto
'                  ocAcessoEmbratel.Produto_ID = Produto_ID
'
'                  'Informa a Controladora Selecionar OTS
'                  Call ocAcessoEmbratel.InformaOC_SelecionarOTS(objOCSelecionarOTS)
'
'                  bRetorno = ocAcessoEmbratel.CarregarFormCFsDaFase(lIdFaseCentroLocal)
'
'                  If (Not bRetorno) Then
'                      GoTo ErrorHandler
'                  End If
'
'                End If 'ObtemSequenciaFaseFinalizada
'
'              End If 'Checa multiplos

          ElseIf (Familia_Produto = "D") Then
    
              'Calcula o período para a proxima fase
              bRetorno = objDadosFase.CalculaPeriodoProximaFase(Sequencia_Acao, Servico_ID, Familia_Produto)
        
              If (objMensagem.TipoMsg <> "") Then
              
                 GoTo ErrorHandler
                 
              End If
          
          End If    'Familia_Produto <> "I"
          
          'Quando for uma OTS de Cancelamento ou Desativacao, verifica se
          'as fases que estão sendo executadas em paralelo já foram finalizadas
          'caso Todas as outras fases já estejam finalizadas, o sistema deverá finalizar
          'o Servico.
          
          'Preenche o periodo realizado indicando que a fase foi finalizada
          bRetorno = objDadosFase.FinalizaFase(Fase_ID)
    
          If (objMensagem.TipoMsg <> "") Then
             
             GoTo ErrorHandler
             
          End If
          
          If (Sigla_Acao = "CAN" Or Sigla_Acao = "DES") And _
                (Familia_Produto = "I") Then

             bRetorno = VerificaFinalizacaoDesativacao(objDadosFase, _
                                                       objdadosServico, _
                                                       Servico_ID, Fase_ID, _
                                                       Sigla_Acao, _
                                                       objMensagem)
          End If
       
       End If
       
       'Produtos de Dados
       If (Familia_Produto = "D") Then
              
            'Obtem a sequencia_acao da última fase existente para a OTS
            lSeqUltimaFaseOTS = objDadosFase.ObtemSequenciaUltimaFase(Servico_ID, Familia_Produto)
            
            If (objMensagem.TipoMsg <> "") Then
            
                objMensagem.ExibeMensagem
                 
                GoTo Finalizar
            
            End If
              
            If (Sequencia_Acao = lSeqUltimaFaseOTS) Then
                 
                Set objDadosAnalisarOE = New clsDNI_DadosAnalisarOE

                objDadosAnalisarOE.InformaMensagem objMensagem
                objDadosAnalisarOE.InformaBanco frmgenMDI_SGAPlus.objBanco
               
               ' sDesignacaoSOE = Trim(objForm.UsrDNICabecalhoAcesso.CNLPontaA) & "/" & _
               '                   Trim(objForm.UsrDNICabecalhoAcesso.CNLPontaB) & "/" & _
               '                   Trim(objForm.UsrDNICabecalhoAcesso.CodigoNP) & "/"
                                  
                If Produto_ID = GbIDProdutoDiginet Then
               
                    'sDesignacaoSOE = sDesignacaoSOE & _
                    'Format(objForm.UsrDNICabecalhoAcesso.SequencialDesig, "0000")
               
                Else
                   
                   ' sDesignacaoSOE = sDesignacaoSOE & _
                   '         Format(objForm.UsrDNICabecalhoAcesso.SequencialDesig, "0000000")
               
                End If

                bRetorno = objDadosAnalisarOE.EncerraWorkFlow(Servico_ID, _
                                Servico_ID_Anterior, _
                                AnoOE, NumeroOE, NumeroItemOE, _
                                "02", _
                                GbUsuario, _
                                sDesignacaoSOE, _
                                Sigla_Acao, _
                                objdadosServico)
              
                Set objDadosAnalisarOE = Nothing
                
                If (Sigla_Acao = "DES" Or _
                        Sigla_Acao = "CAN") Then
                
                    objdadosServico.FinalizaServico (Servico_ID)
                    
                End If
                
            End If    'Sequencia_Acao
            
       End If   'Produtos de Dados
       
       FlagFaseFinalizadaOK = True
       pbFlagDescarregaFormulario = True
       
       
    Else
       'Caso a fase já tenha sido finalizada e uma pendencia tenha sido incluida,
       'A fase deverá ser considerada como Em Andamento
       
       objDadosFase.ID_Servico = Servico_ID
       objDadosFase.Id_Fase = Fase_ID
          
       'Quando for OTS de Desativacao ou Cancelamento, o periodo previsto para a fase seguinte
       'deverá ser mantido já que as fases para estas ações serão executadas em paralelo.
       If (DataFinalizacao <> "") Then
            
          If (Sigla_Acao = "DES") Or _
                (Sigla_Acao = "CAN") Then
                
             FlagManterPeriodo = True
                 
          Else
              FlagManterPeriodo = False
          End If
               
          bRetorno = objDadosFase.RetirarFinalizacaoFase(Servico_ID, Fase_ID, Sequencia_Acao, _
                                                    FlagManterPeriodo, Familia_Produto)

          If (objMensagem.TipoMsg <> "") Then
                
             Set objDadosFase = Nothing
             
             GoTo ErrorHandler
               
          End If
          
       End If
        
    End If
    
     'Obtem ID_fase da Alocacao Configuraçao
    lIdFaseAlocacaoConfiguracao = objDadosFase.ObtemIDFase(objdadosServico.ID_Servico, GbCodigoFaseConfiguracao)
    
    If objMensagem.TipoMsg <> "" Then
      
         GoTo Finalizar
    
    End If
     
    If lIdFaseAlocacaoConfiguracao <> 0 Then
    
         bRetorno = objDadosFase.VerificaFaseFinalizada(lIdFaseAlocacaoConfiguracao)
     
         If bRetorno Then
            
            lIdFaseCentroLocal = objDadosFase.ObtemIDFase(objdadosServico.ID_Servico, GbCodigoFaseCentroLocal)
            
            'bRetorno = objDadosFase.CalculaPeriodoProximaFase(lIdFaseCentroLocal, objdadosServico.ID_SERVICO)
             'iSeqAux = objDadosFase.ObtemSequenciadaFase(objdadosServico.ID_SERVICO, GbCodigoFaseConfigE1)
    
             'bRetorno = objDadosFase.CalculaPeriodoProximaFase(5, Servico_ID)
             bRetorno = objDadosFase.CalculaPeriodoProximaFase(Sequencia_Acao + 2, Servico_ID)
            
            If objMensagem.TipoMsg <> "" Then
            
                 GoTo Finalizar
            
            End If
         
         End If
   
          If (objWEBIP Is Nothing) Then
              Set objWEBIP = New clsInterOC_TrataWebIP
          End If
          
          objWEBIP.Grupo_ID = lGrupo_ID
          
         '=================================================================
         'CALLBAT - 03/08/2013 - Resolução 600
         'SO-TELCO-79171SQV-04
         '=================================================================
         'objWEBIP.Indic_PMS = IIf((objForm.chkInterligacaoPMS.Value) = 1, "S", "N")
         'objWEBIP.Num_Ped_Venda_SNOA = objForm.lblNumPedidoSNOA
         '=================================================================
         '=================================================================
        
         
         'AGRC - 07/01/2018 - Flex VPN Fase 2
         Set objBldOrderEntry = New clsbldOrderEntry
          objBldOrderEntry.InformaBanco frmgenMDI_SGAPlus.objBanco
          bRetorno = objBldOrderEntry.atualizarDadosFlexVPN( _
              objForm.txtVRF.Text, _
              objForm.txtIpLoopbackLnsVrfCliente.Text, _
              objForm.TxtBlocoIpLanCliente.Text, _
              objForm.txtIpLoopbackClienteCpe.Text, _
              objdadosServico.SiglaUD, _
              objdadosServico.NumContrato, _
              objdadosServico.AnoContrato, _
              objdadosServico.VersaoContrato _
              )

          If (objMensagem.TipoMsg <> "") Then
            
            GoTo ErrorHandler
          
          End If
        
        
        
          bRetorno = objWEBIP.GravaWEBIPIntegracao(Servico_ID, _
                      Fase_ID, _
                      "C", _
                      Produto_ID, _
                      objMensagem, _
                      Sigla_Acao)
        
          If (objMensagem.TipoMsg <> "") Then
            
            GoTo ErrorHandler
          
          End If
         
         
         'AGRC - 07/01/2018 - Flex VPN Fase 2
         '                             ByVal vOrgaoContrato As String, _
         '                             ByVal vNumeroContrato As String, _
         '                             ByVal vAnoContrato As String, _
         '                             ByVal vVersaoContrato As String _
          'TBASTOS - FLEX VPN - Atualiza dados da Interface com o GSI.
'          Set objBldOrderEntry = New clsbldOrderEntry
'          objBldOrderEntry.InformaBanco frmgenMDI_SGAPlus.objBanco
'          bRetorno = objBldOrderEntry.atualizarDadosFlexVPN( _
'              objForm.txtVRF.Text, _
'              objForm.txtIpLoopbackLnsVrfCliente.Text, _
'              objForm.TxtIPPublicoLoopbackLns.Text, _
'              objForm.TxtBlocoIpLanCliente.Text, _
'              objForm.txtIpLoopbackClienteCpe.Text, _
'              objForm.txtIpLoopbackClienteTT.Text, _
'              objdadosServico.SiglaUD, _
'              objdadosServico.NumContrato, _
'              objdadosServico.AnoContrato, _
'              objdadosServico.VersaoContrato _
'              )
'
'          If (objMensagem.TipoMsg <> "") Then
'
'            GoTo ErrorHandler
'
'          End If
          
          
          
          
    End If

    '=========================================================
    '=========================================================
     
     'CALBAT - 05/10/2007 =======================================
     'Tratar Acesso Paralelismo.                                                              =
     'Verifica se a fase seguinte está preenchida, de Configuração está  =
     'OK, caso esteja, inicia a fase de Tratar Centro Local                      =
     '===========================================================
     
     lID_Fase_Seguinte = objDadosFase.ObtemIDFaseSeguinte(Servico_ID, Sequencia_Acao)

     If objMensagem.TipoMsg <> "" Then
       
         GoTo ErrorHandler
     
     End If
    
     If objDadosFase.ObtemDataFimPrevisto(lID_Fase_Seguinte) = "" Then
     
        bRetorno = objDadosFase.CalculaPeriodoProximaFase(Sequencia_Acao, Servico_ID)
        
        If objMensagem.TipoMsg <> "" Then
          
            GoTo ErrorHandler
        
        End If
        
     End If

    bRetorno = frmgenMDI_SGAPlus.objBanco.Finalizar_Transacao
    
    
    If (Not bRetorno) Then
        
        GoTo ErrorHandler
    
    End If
    
'    objForm.FlControlaAlteracao = False
    bFlagControlaAlteracao = False
    
    GravarDados = True

        objMensagem.TipoMsg = "A"

    If FlagFaseFinalizadaOK = True And Not FlagTemPendenciaEmAberto Then
    

        objMensagem.Descricao = "Os Dados de Acesso Flex VPN foram " & _
                               "atualizados com Sucesso!"

'#########################
'adeil RN-SGAD-81746VNS 01/10/2003
'inicio


    ElseIf Not FlagFaseFinalizadaOK And Not FlagTemPendenciaEmAberto Then

       objMensagem.Descricao = "Os Dados de Acesso Flex VPN foram atualizados com Sucesso. " & vbCrLf & _
                               " Porém a fase continua em Andamento," & vbCrLf & _
                               " pois campos obrigatórios não foram preenchidos."
    
    
    ElseIf Not FlagFaseFinalizadaOK And FlagTemPendenciaEmAberto Then

       objMensagem.Descricao = "Os Dados de Acesso Flex VPN foram atualizados com Sucesso. " & vbCrLf & _
                               " Porém a fase continua em andamento, devido a existência de pendência!"
    
    
    End If
    
    'FlagDescarregaForm = True
    
    pbFlagDescarregaFormulario = True
    
'fim
'#########################
    
    objMensagem.ExibeMensagem
    
    GoTo Finalizar
    
Finalizar:
    
    Set objdadosServico = Nothing
    Set objDadosFase = Nothing
    Set objWEBIP = Nothing
    Set objDadosDuracaoFase = Nothing
    Set ocAcessoEmbratel = Nothing
    
    Exit Function
    
ErrorHandler:
    ' Cancelar Transação
    If (FlagIniciouTransacao) Then
       
       bRetorno = frmgenMDI_SGAPlus.objBanco.Cancelar_Transacao
    
    End If
    
    If (Err.Number <> 0) Then
       
       objMensagem.TipoMsg = "E"
       objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                   "Houve um erro de execução no Método: " & vbCrLf & vbCrLf & "GravarDados " & _
                   "da classe clsblfOC_TratarAcessoFlexVPN"
    End If
       
    objMensagem.ExibeMensagem
    
    GoTo Finalizar
           
End Function
Public Property Get TipoOperacao() As String

TipoOperacao = sTipoOperacao

End Property

Public Property Let TipoOperacao(ByVal vNewValue As String)

sTipoOperacao = vNewValue

End Property

Public Property Get IdCircuito() As Long

IdCircuito = lIdCircuito

End Property

Public Property Let IdCircuito(ByVal lNewValue As Long)

lIdCircuito = lNewValue

End Property

Public Property Get Sigla_Acao() As String

Sigla_Acao = sSigla_Acao

End Property

Public Property Let Sigla_Acao(ByVal sNewValue As String)

sSigla_Acao = sNewValue

End Property

Public Property Get PermiteExecucao() As String

PermiteExecucao = sPermiteExecucao

End Property

Public Property Let PermiteExecucao(ByVal sNewValue As String)

sPermiteExecucao = sNewValue

End Property

Public Property Get Status_Fase() As String

Status_Fase = sStatus_Fase

End Property

Public Property Let Status_Fase(ByVal vNewValue As String)

sStatus_Fase = vNewValue

End Property

Public Property Get Acao_ID() As Long

Acao_ID = lAcao_ID

End Property

Public Property Let Acao_ID(ByVal lNewValue As Long)

lAcao_ID = lNewValue


End Property


Public Property Get DataFinalizacao() As String

DataFinalizacao = sDataFinalizacao

End Property

Public Property Let DataFinalizacao(ByVal sNewValue As String)

sDataFinalizacao = sNewValue

End Property

Public Function VerificaFinalizacaoOK(Sigla_Acao As String) As Boolean

VerificaFinalizacaoOK = False

    If objForm.txtVRF.Text <> "" And _
    objForm.txtIpLoopbackLnsVrfCliente.Text <> "" And _
    objForm.TxtBlocoIpLanCliente.Text <> "" And _
    objForm.txtIpLoopbackClienteCpe <> "" Then
    
    VerificaFinalizacaoOK = True
Else
    
    If Sigla_Acao = "DES" Or Sigla_Acao = "CAN" Then
    
       VerificaFinalizacaoOK = True
        
    End If
    
End If

End Function

Private Function DesabilitaObjetosTela()
Dim obj As Object

For Each obj In objForm
    If TypeOf obj Is TextBox Then
        obj.Enabled = False
    End If
Next

    objForm.CmdSalvar.Enabled = False
    objForm.CmdLimpar.Enabled = False
    
End Function

Public Function InformaOCSelecionarOTS(ByVal oobjOCSelecionarOTS As clsgenOC_SelecionarOTS)

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
                            "InformaOCSelecionarOTS da classe clsblfOC_TratarAcessoCliente"
    
    GoTo Finalizar

End Function

Public Function AtualizarListaOTS() As Boolean

    objOCSelecionarOTS.ExibirListaOTS
        
End Function

Private Sub Class_Terminate()

    Set objDadosPendencia = Nothing
    Set objMensagem = Nothing
    Set objDadosAcessoCliente = Nothing
    Set objForm = Nothing
    
End Sub
Public Function Imprime_OTS() As Boolean
    
'    Dim objnOTS As clsblf_OTS
'
'    Set objnOTS = New clsblf_OTS
'
'    objnOTS.InformaBanco frmgenMDI_SGAPlus.objBanco
'    objnOTS.ImprimeOTS (lIdServico)
'
'    Imprime_OTS = True

End Function


Private Function PopularCamposTela()

'    Dim iAux                              As Integer
'    Dim objOC_AcessoEBT        As New clsGenOC_TratarAcessoEmbratel
'
'    'CALBAT - 05/10/2007 =====================================
'    'Tratar Acesso Paralelismo.                                                          =
'    '=========================================================
'    Dim objDadosIntegracao      As New clsBlf_IntegracaoWebIP
'
'    objDadosIntegracao.InformaBanco frmgenMDI_SGAPlus.objBanco
'    objDadosIntegracao.InformaMensagem objMensagem
'
'    If (objDadosIntegracao Is Nothing) Then
'        Set objDadosIntegracao = New clsBlf_IntegracaoWebIP
'    End If
'
'
'    bRetorno = objDadosIntegracao.VerificaIntegracao(lServico_ID)
'
'    If bRetorno = True Then
'
'        'Bloqueia os campos já encaminhado para o SAIP
'        objForm.CboPadraoInterfaceEBT.Locked = True
'        objForm.txtEstacaoEntregaConfiguracao.Locked = True
'        objForm.optTipoRede(0).Enabled = False
'        objForm.optTipoRede(1).Enabled = False
'        objForm.optTipoRede(2).Enabled = False
'        objForm.TxtDescricaoRedeDeterministica.Enabled = False
'        objForm.txtTSI.Enabled = False
'        objForm.txtTSF.Enabled = False
'
'        'Muda as cores dos campos bloqueado
'        objForm.CboPadraoInterfaceEBT.BackColor = &H80000004
'        objForm.txtEstacaoEntregaConfiguracao.BackColor = &H80000004
'        objForm.TxtDescricaoRedeDeterministica.BackColor = &H80000004
'        objForm.txtTSI.BackColor = &H80000004
'        objForm.txtTSF.BackColor = &H80000004
'
'
'        'Informa ao usuário o motivo do bloqueio
'        objForm.txtEstacaoEntregaConfiguracao.ToolTipText = "Este campo não pode ser editado, em função " & Chr(13) & _
'                                                                                             "de já ter enviado as informações para o SAIP"
'
'        objForm.txtEstacaoEntregaConfiguracao.ToolTipText = "Este campo não pode ser editado, em função " & Chr(13) & _
'                                                                                             "de já ter enviado as informações para o SAIP"
'
'        objForm.optTipoRede(0).ToolTipText = "Este campo não pode ser editado, em função " & Chr(13) & _
'                                                                                             "de já ter enviado as informações para o SAIP"
'
'        objForm.optTipoRede(1).ToolTipText = "Este campo não pode ser editado, em função " & Chr(13) & _
'                                                                                             "de já ter enviado as informações para o SAIP"
'
'        objForm.optTipoRede(2).ToolTipText = "Este campo não pode ser editado, em função " & Chr(13) & _
'                                                                                             "de já ter enviado as informações para o SAIP"
'
'
'    Else
'
'        'Bloqueia os campos já encaminhado para o SAIP
'        objForm.CboPadraoInterfaceEBT.Locked = False
'        objForm.txtEstacaoEntregaConfiguracao.Locked = False
'        objForm.optTipoRede(0).Enabled = True
'        objForm.optTipoRede(1).Enabled = True
'        objForm.optTipoRede(2).Enabled = False
'        objForm.TxtDescricaoRedeDeterministica.Enabled = True
'        objForm.txtTSI.Enabled = True
'        objForm.txtTSF.Enabled = True
'
'        'Muda as cores dos campos bloqueado
'        objForm.CboPadraoInterfaceEBT.BackColor = &H80000005
'        objForm.txtEstacaoEntregaConfiguracao.BackColor = &H80000005
'        objForm.TxtDescricaoRedeDeterministica.BackColor = &H80000004
'        objForm.txtTSI.BackColor = &H80000004
'        objForm.txtTSF.BackColor = &H80000004
'
'
'        'Informa ao usuário o motivo do bloqueio
'        objForm.txtEstacaoEntregaConfiguracao.ToolTipText = ""
'
'        objForm.txtEstacaoEntregaConfiguracao.ToolTipText = ""
'        objForm.optTipoRede(0).ToolTipText = ""
'        objForm.optTipoRede(1).ToolTipText = ""
'        objForm.optTipoRede(2).ToolTipText = ""
'
'    End If
'
'
'    Set objDadosIntegracao = Nothing
'
'    '=========================================================
'    '=========================================================
'
'    With objForm
'
'        If objDadosAcessoCliente.CodigoProvedorAcesso <> "" Then
'
'           Call SelecionaCombo(objForm.CboProvedorUrbano, CLng(objDadosAcessoCliente.CodigoProvedorAcesso))
'
'        End If
'
'        .TxtEstacaoEntregaEbt.Text = "" & objDadosAcessoCliente.EstacaoEntregaEBT
'
'        If objDadosAcessoCliente.FacilidadeUtilizada <> "" Then
'
'           For iAux = 0 To .cmbFacilidadeUtilizada.ListCount - 1
'
'               If .cmbFacilidadeUtilizada.List(iAux) = _
'                   objDadosAcessoCliente.FacilidadeUtilizada Then
'
'                   .cmbFacilidadeUtilizada.ListIndex = iAux
'                   Exit For
'               End If
'
'           Next iAux
'
'        End If
'
'        .txtEstacaoEntregaConfiguracao.Text = "" & objDadosAcessoCliente.EstacaoEntregaConfiguracao
'        .TxtDesignacaoAcessoEBT.Text = "" & objDadosAcessoCliente.DesignacaoAcessoEBT
'        .TxtNumeroOTSTX.Text = "" & objDadosAcessoCliente.NumeroOTSTX
'        .TxtItemOTSTx.Text = "" & objDadosAcessoCliente.ItemOTSTX
'
'        .txtDescrEqptoCliente.Text = "" & objDadosAcessoCliente.DescrEqptoCliente
'
'        .txtCabeamentoInternoCliente.Text = "" & objDadosAcessoCliente.CabeamentoInternoCliente
'        .txtEntroncamento.Text = "" & objDadosAcessoCliente.Entroncamento
'
'        '.txtNumeroNo.Text = "" & objDadosAcessoCliente.NumeroNo
'        '.TxtDominio.Text = "" & objDadosAcessoCliente.Dominio
'        '.txtNo.Text = "" & objDadosAcessoCliente.NO
'        '.TxtSlot.Text = "" & objDadosAcessoCliente.Slot
'        '.txtPorta.Text = "" & objDadosAcessoCliente.Porta
'
'        .TxtDescricaoRedeDeterministica.Text = objDadosAcessoCliente.DescricaoRedeDeterministica
'
'
'        .txtTSI.Text = "" & objDadosAcessoCliente.TSI
'        .txtTSF.Text = "" & objDadosAcessoCliente.TSF
'
'        .SpdNumeroAcessoUrbano.Text = "" & objDadosAcessoCliente.NumeroAcessoProvedor
'
'
'        .txtNumeroCaboEBT.Text = "" & objDadosAcessoCliente.NumeroCaboEBT
'        .txtNumeroParEBT.Text = "" & objDadosAcessoCliente.NumeroParEBT
'
'        .SpdDataAtivacaoAcessoCliente.Row = 1
'        .SpdDataAtivacaoAcessoCliente.Col = 1
'        .SpdDataAtivacaoAcessoCliente.Text = IIf(Not IsNull(objDadosAcessoCliente.DtAtivacaoAcessoCliente), _
'                                Format(objDadosAcessoCliente.DtAtivacaoAcessoCliente, "dd/mm/yyyy"), "")
'
'        If Familia_Produto = "I" Then
'
'            If Val(objDadosAcessoCliente.Cd_Veloc) <> 0 Then
'
'               Call SelecionaCombo(objForm.CboVelocidadeAcesso, _
'                                   CLng(objDadosAcessoCliente.Cd_Veloc))
'
'            End If
'
'
'        'CALBAT - 05/10/2007 =====================================
'        'Tratar Acesso Paralelismo.                                                          =
'        '=========================================================
'            If objDadosAcessoCliente.Interface_ID > 0 Then
'
'               Call SelecionaCombo(objForm.CboPadraoInterfaceEBT, _
'                                   CLng(CLng(objDadosAcessoCliente.Interface_ID)))
'
'            End If
'
'
'            If objDadosAcessoCliente.Tipo_Rede = "E" Then
'
'                .optTipoRede(0) = True
'
'            ElseIf objDadosAcessoCliente.Tipo_Rede = "T" Then
'
'                .optTipoRede(1) = True
'
'            ElseIf objDadosAcessoCliente.Tipo_Rede = "D" Then
'
'                .optTipoRede(2) = True
'
'            End If
'
'
'
'        '=========================================================
'        '=========================================================
'
'        ElseIf Familia_Produto = "D" Then
'
'            Call objOC_AcessoEBT.SelecionaComboTexto_DNI(objForm.CboVelocidadeAcesso, objDadosAcessoCliente.Cd_Veloc)
'
'        End If
'
'    End With
'
End Function

Private Function CarregarAtributosAcesso()

'    With objForm
'
'        If .CboProvedorUrbano.ListIndex > -1 Then
'
'           objDadosAcessoCliente.CodigoProvedorAcesso = _
'                                .CboProvedorUrbano.ItemData(.CboProvedorUrbano.ListIndex)
'
'        Else
'
'           objDadosAcessoCliente.CodigoProvedorAcesso = ""
'
'        End If
'
'        objDadosAcessoCliente.EstacaoEntregaEBT = TrataPlick(.TxtEstacaoEntregaEbt.Text)
'        objDadosAcessoCliente.DesignacaoAcessoEBT = TrataPlick(.TxtDesignacaoAcessoEBT.Text)
'        objDadosAcessoCliente.NumeroOTSTX = TrataPlick(.TxtNumeroOTSTX.Text)
'        objDadosAcessoCliente.ItemOTSTX = TrataPlick(.TxtItemOTSTx.Text)
'
'        If .cmbFacilidadeUtilizada.ListIndex > -1 Then
'
'            objDadosAcessoCliente.FacilidadeUtilizada = _
'                        Trim(.cmbFacilidadeUtilizada.List(.cmbFacilidadeUtilizada.ListIndex))
'
'        Else
'
'            objDadosAcessoCliente.FacilidadeUtilizada = ""
'
'        End If
'
'        objDadosAcessoCliente.EstacaoEntregaConfiguracao = TrataPlick( _
'                                                       .txtEstacaoEntregaConfiguracao.Text)
'
'        objDadosAcessoCliente.DescrEqptoCliente = TrataPlick(.txtDescrEqptoCliente.Text)
'
'        objDadosAcessoCliente.CabeamentoInternoCliente = _
'                                                TrataPlick(.txtCabeamentoInternoCliente.Text)
'        objDadosAcessoCliente.Entroncamento = TrataPlick(.txtEntroncamento.Text)
'
''        objDadosAcessoCliente.NumeroNo = TrataPlick(.txtNumeroNo.Text)
''        objDadosAcessoCliente.Dominio = .TxtDominio.Text
''        objDadosAcessoCliente.NO = TrataPlick(.txtNo.Text)
''        objDadosAcessoCliente.Slot = TrataPlick(.TxtSlot.Text)
''        objDadosAcessoCliente.Porta = TrataPlick(.txtPorta.Text)
'
'        objDadosAcessoCliente.Interface = .CboPadraoInterfaceEBT.Text
'
'        objDadosAcessoCliente.DescricaoRedeDeterministica = TrataPlick(.TxtDescricaoRedeDeterministica.Text)
'
'
'        objDadosAcessoCliente.TSI = TrataPlick(.txtTSI.Text)
'        objDadosAcessoCliente.TSF = TrataPlick(.txtTSF.Text)
'
'        objDadosAcessoCliente.NumeroAcessoProvedor = _
'                                                    TrataPlick(.SpdNumeroAcessoUrbano.Text)
'
'        'objDadosAcessoCliente.Interface = TrataPlick(.txtInterface.Text)
'        objDadosAcessoCliente.NumeroCaboEBT = TrataPlick(.txtNumeroCaboEBT.Text)
'        objDadosAcessoCliente.NumeroParEBT = TrataPlick(.txtNumeroParEBT.Text)
'
'        .SpdDataAtivacaoAcessoCliente.Row = 1
'        .SpdDataAtivacaoAcessoCliente.Col = 1
'
'        If Trim(.SpdDataAtivacaoAcessoCliente.Text) <> "" Then
'
'           objDadosAcessoCliente.DtAtivacaoAcessoCliente = _
'                                                          .SpdDataAtivacaoAcessoCliente.Text
'
'        Else
'
'           objDadosAcessoCliente.DtAtivacaoAcessoCliente = ""
'
'        End If
'
'        'CALBAT - 05/10/2007 =====================================
'        'Tratar Acesso Paralelismo.                                                          =
'        '=========================================================
'        If .optTipoRede(0).Value = True Then
'
'            objDadosAcessoCliente.Tipo_Rede = "E"
'
'        ElseIf .optTipoRede(1).Value = True Then
'
'            objDadosAcessoCliente.Tipo_Rede = "T"
'
'        ElseIf .optTipoRede(2).Value = True Then
'
'            objDadosAcessoCliente.Tipo_Rede = "D"
'
'        End If
'        '=========================================================
'        '=========================================================
'
'        If Familia_Produto = "I" Then
'
'            If .CboVelocidadeAcesso.ListIndex > -1 Then
'
'               objDadosAcessoCliente.Cd_Veloc = .CboVelocidadeAcesso.ItemData( _
'                                                                 .CboVelocidadeAcesso.ListIndex)
'            End If
'
'        ElseIf Familia_Produto = "D" Then
'
'            If objForm.CboVelocidadeAcesso.ListIndex > 0 Then
'
'                objDadosAcessoCliente.Cd_Veloc = vVelocidadeDNI(objForm.CboVelocidadeAcesso.ListIndex).CodigoAlpha
'
'            Else
'
'                objDadosAcessoCliente.Cd_Veloc = ""
'
'            End If
'
'        End If
'
'    End With
    
End Function

'Private Function CarregaCabecalhoAcesso_BLF(Usr As UsrBLFCabecalho_Acesso, objBlfCabecalho As clsBlf_Cabecalho) As Boolean
'
'    With objBlfCabecalho
'
'        Usr.ContaCorrente = .ContaCorrente
'        Usr.DataFinalizacao = .DataFinalizacao
'        Usr.DesignacaoPrincipal = .DesignacaoPrincipal
'        Usr.DesignacaoSecundaria = .DesignacaoSecundaria
'        Usr.NomeCliente = .NomeCliente
'        Usr.OtsPrincipal = .OtsPrincipal
'        Usr.OtsSecundaria = .OTSAssociada
'        Usr.TipoCircuito = .TipoCircuito
'        Usr.NumeroOE = .NumeroOE
'        Usr.AnoOE = .AnoOE
'        Usr.ItemOE = .ItemOE
'        Usr.Contrato = .Contrato
'        Usr.CNL = .CNL
'
'    End With
'
'End Function


Private Function CarregarFormAlocacaoCF(psCodigoCFCentroLocal As String) As Long

    Dim objFormAlocacao    As New frmInterCFCtrLocal
    Dim bRetorno           As Boolean
    
    CarregarFormAlocacaoCF = 0
    
    Load objFormAlocacao
    
    If lProduto_ID = GbIDProdutoWID Then
    
        ' Dedicated
        objFormAlocacao.Caption = "MCI Internet Dedicated"
        
    ElseIf lProduto_ID = GbIDProdutoVPN_RA Then
    
        ' VPN RA
        objFormAlocacao.Caption = "VPN RA"
        
    ElseIf lProduto_ID = GbIDProdutoVPN_DIRECT Then
    
        ' VPN Direct
        objFormAlocacao.Caption = "VPN Direct"
        
    ElseIf lProduto_ID = GbIDProdutoBLF Then
    
        ' BLF
        objFormAlocacao.Caption = "Business Link Flex"
        
    ElseIf (lProduto_ID = GbIDProdutoBLD) Then
        
        ' BLD
        objFormAlocacao.Caption = "Business Link Direct"
        
    ElseIf (lProduto_ID = GbIDProdutoVOIP) Then
        
        ' BLD
        objFormAlocacao.Caption = "Internet Voice"
        
    End If
    
    objFormAlocacao.lblCentroFuncionalAlocado.Caption = psCodigoCFCentroLocal
    
    bRetorno = objOCSelecionarOTS.ExibirListaCentroFuncional( _
                         objFormAlocacao.cboCentroFuncional, _
                         GbCodigoFaseCentroLocal, _
                         True)
        
    objFormAlocacao.informa_OC Me
    
    objFormAlocacao.Show vbModal
    
    CarregarFormAlocacaoCF = CentroFuncionalCL_ID
    
    Set objFormAlocacao = Nothing
    
End Function

Public Function CamposPreenchidosOK() As Boolean

    Dim objDadosFase    As clsgenFase
    Dim bRetorno        As Boolean
    Dim sDataFim        As String
    
    CamposPreenchidosOK = False
    
   'Se algum campo estiver vazio.
   If objForm.TxtBlocoIpLanCliente.Text = "" Or _
      objForm.txtIpLoopbackClienteCpe.Text = "" Or _
      objForm.txtIpLoopbackLnsVrfCliente.Text = "" Or _
      objForm.txtVRF.Text = "" Then
       MsgBox "O preenchimento de todos os campos é obrigatório.", vbExclamation
      CamposPreenchidosOK = False
      Exit Function
   End If

    'validação formato de IP
    If Not IsValidIP(objForm.TxtBlocoIpLanCliente.Text, "S") Then
       MsgBox "O Campo Bloco IP Lan Cliente possui um endereço inválido.", vbExclamation
       objForm.TxtBlocoIpLanCliente.SetFocus
       objForm.TxtBlocoIpLanCliente.SelStart = 0
       objForm.TxtBlocoIpLanCliente.SelLength = Len(objForm.TxtBlocoIpLanCliente.Text)
       CamposPreenchidosOK = False
       Exit Function
   End If

    If Not IsValidIP(objForm.txtIpLoopbackClienteCpe.Text) Then
       MsgBox "O Campo IP Loopback Cliente CPE possui um endereço inválido.", vbExclamation
       objForm.txtIpLoopbackClienteCpe.SetFocus
       objForm.txtIpLoopbackClienteCpe.SelStart = 0
       objForm.txtIpLoopbackClienteCpe.SelLength = Len(objForm.txtIpLoopbackClienteCpe.Text)
       CamposPreenchidosOK = False
       Exit Function
   End If
   
    If Not IsValidIP(objForm.txtIpLoopbackLnsVrfCliente.Text) Then
       MsgBox "O Campo IP Loopback LNS VRF Cliente possui um endereço inválido.", vbExclamation
       objForm.txtIpLoopbackLnsVrfCliente.SetFocus
       objForm.txtIpLoopbackLnsVrfCliente.SelStart = 0
       objForm.txtIpLoopbackLnsVrfCliente.SelLength = Len(objForm.txtIpLoopbackLnsVrfCliente.Text)
       CamposPreenchidosOK = False
       Exit Function
   End If


    
    CamposPreenchidosOK = True
    
End Function
