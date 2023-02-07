VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGenOC_TratarAcessoEmbratel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public objDadosPendencia          As New clsblf_Pendencia
Public objDadosAcesso             As New clsGenAcesso
Public objMensagem                As New clsgenMensagem
Public objdadosServico            As New clsGenServico

Public FlagRecuperando            As Boolean

Public objblfOC_SelecionarOTS     As clsgenOC_SelecionarOTS
Public bFlagControlaAlteracao     As Boolean
Public objForm                    As frmgenTratarAcessoEmbratel

Public vMaskProvedor              As Variant
Public vMensagemListaMsg          As Variant

Dim oobjPendencia                 As clsDNI_DadosPendencia
Dim sTipoOperacao                 As String
Dim sStatus_Fase                  As String
Dim sSigla_Acao                   As String
Dim lServico_ID                   As Long
Dim lAcao_ID                      As Long
Dim lServico_ID_Anterior          As Long
Dim lProduto_ID                   As Long
Dim lFase_Id                      As Long
Dim sDataFinalizacao              As String
Dim lSequencia_Acao               As Long
Dim sTextoAnotacao                As String
Dim sTipoAcessoAnterior           As String
Dim sPropriedadeAcessoAnterior    As String
Dim sCaracteristicaAcessoAnterior As String
Dim sSegurancaAcessoAnterior      As String
Dim lNrLocalInsta                 As Long
Dim lContratoID                   As Long
Dim lCentroFuncID_Realoc         As Long
Dim sPermiteExecucao              As String
Dim lCentroFuncionalID            As Long
Dim lNumeroOE                     As Long
Dim lAnoOE                        As Long
Dim lNumeroITemOE                 As Long
Dim sFamilia_Produto              As String
Dim lEnderecoID                   As Long
Dim lItemServicoID                As Long
'Não resgata os dados da OTS anterior quando se altera o
'endereço de instalação. IP VPN
Dim bResgataAnterior              As Boolean

Public Property Get Familia_Produto() As String

    Familia_Produto = sFamilia_Produto

End Property

Public Property Let Familia_Produto(ByVal sNewValue As String)

    sFamilia_Produto = sNewValue
    
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

Public Property Get objPendencia() As clsDNI_DadosPendencia

Set objPendencia = oobjPendencia

End Property

Public Property Let objPendencia(ByVal vNewValue As clsDNI_DadosPendencia)

Set oobjPendencia = vNewValue

End Property

Public Property Get CentroFuncID_Realoc() As Long

CentroFuncID_Realoc = lCentroFuncID_Realoc

End Property

Public Property Let CentroFuncID_Realoc(ByVal lNewValue As Long)

lCentroFuncID_Realoc = lNewValue

End Property

Public Property Get CentroFuncionalID() As Long
CentroFuncionalID = lCentroFuncionalID

End Property

Public Property Let CentroFuncionalID(ByVal lNewValue As Long)

lCentroFuncionalID = lNewValue

End Property

Public Property Get ContratoID() As Long

ContratoID = lContratoID

End Property

Public Property Let ContratoID(ByVal lNewValue As Long)

lContratoID = lNewValue

End Property

Public Property Get EnderecoID() As Long

EnderecoID = lEnderecoID

End Property

Public Property Let EnderecoID(ByVal lNewValue As Long)

lEnderecoID = lNewValue

End Property

Public Property Get NrLocalInsta() As Long

NrLocalInsta = lNrLocalInsta

End Property

Public Property Let NrLocalInsta(ByVal lNewValue As Long)

lNrLocalInsta = lNewValue

End Property

Public Property Get PermiteExecucao() As String

PermiteExecucao = sPermiteExecucao

End Property

Public Property Let PermiteExecucao(ByVal sNewValue As String)

sPermiteExecucao = sNewValue

End Property

Public Property Get SegurancaAcessoAnterior() As String

SegurancaAcessoAnterior = sSegurancaAcessoAnterior

End Property

Public Property Let SegurancaAcessoAnterior(ByVal sNewValue As String)

sSegurancaAcessoAnterior = sNewValue

End Property

Public Property Get CaracteristicaAcessoAnterior() As String

CaracteristicaAcessoAnterior = sCaracteristicaAcessoAnterior

End Property

Public Property Let CaracteristicaAcessoAnterior(ByVal sNewValue As String)

sCaracteristicaAcessoAnterior = sNewValue

End Property

Public Property Get TipoAcessoAnterior() As String

TipoAcessoAnterior = sTipoAcessoAnterior

End Property

Public Property Let TipoAcessoAnterior(ByVal sNewValue As String)

sTipoAcessoAnterior = sNewValue

End Property

Public Property Get PropriedadeAcessoAnterior() As String

PropriedadeAcessoAnterior = sPropriedadeAcessoAnterior

End Property

Public Property Let PropriedadeAcessoAnterior(ByVal sNewValue As String)

sPropriedadeAcessoAnterior = sNewValue

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


Public Property Get DataFinalizacao() As String

DataFinalizacao = sDataFinalizacao

End Property

Public Property Let DataFinalizacao(ByVal sNewValue As String)

sDataFinalizacao = sNewValue

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
'VIPLINE IP
Public Property Get ItemServicoID() As Long

ItemServicoID = lItemServicoID

End Property

Public Property Let ItemServicoID(ByVal lNewValue As Long)

lItemServicoID = lNewValue

End Property

Public Function CarregarForm() As Boolean

    Dim objAutenticacao        As New clsgenAutenticacao
    Dim objOC_DadosSwitch As clsVIPoc_CadastrarEqptoNGN
    Dim objDadosFase           As New clsgenFase
    Dim objPendenciaDNI        As New clsDNI_DadosPendencia
    Dim objOc_TratarEqptoCompl As clsInterOC_TratarEqptoCompl
    Dim objDadosProduto        As New clsGenProduto
    Dim paPermissoes()         As Variant
    Dim objDadosCtrlCorrecao As clsGenDadosControleCorrecao
    Dim sSituacaoOTS           As String
    Dim numUltimaFaseRealizada As Long
    Dim IdOTS                  As Long
    Dim IDFase                 As Long
    Dim lSequenciaFaseE1       As Long
    Dim lSequenciaFaseRB       As Long
    Dim cColecao               As Collection
    Dim objDadosCadastrais     As Object
    Dim objDadosAcessoSOE      As Object
    Dim objDadosAnalisarOE     As clsDNI_DadosAnalisarOE
    Dim FlagInclusao           As Boolean
    Dim flUsuarioDoCF          As Boolean
    Dim objDniDadosTecnico     As clsDNI_DadosTecnicos
    Dim sAcaoAux               As String
    Dim sSiglaProduto           As String
    Dim objDadosEqptoNGN        As VIP_Equipamento_NGN
    Dim lIDFase As Long
    'Não resgata os dados da OTS anterior quando se altera o
    'endereço de instalação. IP VPN
    Dim sCNLAtual              As String
    Dim sCNLAnterior           As String
    Dim objdadosCircuito       As New clsgenCircuito
    Dim sStatusSwitch          As String
    Dim objVIPDadosEnder As VIP_Aten_Endereco
    Dim objVIPDadosContrato As VIP_Aten_Contrato
    Dim objVIPDadosAcesso   As VIP_Aten_Dados_Acesso

    Dim sTabela_Interface    As String   'JAH
    
    On Error GoTo ErrorHandler
    
    FlagRecuperando = True
    
    CarregarForm = False
    'Verfica se a OTS selecionada ainda encontra-se em andamento
    
    objdadosServico.InformaBanco frmgenMDI_SGAPlus.objBanco
    objdadosServico.InformaMensagem objMensagem
    
    bRetorno = objdadosServico.RecuperarDados(Servico_ID)
    
    sSituacaoOTS = objdadosServico.StatusServico
    
    If objMensagem.TipoMsg <> "" Then
       
       objMensagem.ExibeMensagem
       
       GoTo Finalizar
       
    End If
    
    If sSituacaoOTS = "CAN" Or sSituacaoOTS = "CR" Then
       
       objMensagem.TipoMsg = "A"
       objMensagem.Descricao = "Esta OTS foi paralisada pelo SOE!"
       objMensagem.ExibeMensagem
       
       GoTo Finalizar
       
    End If
    
    'Circuito
    Set objdadosCircuito = New clsgenCircuito
    
    objdadosCircuito.InformaMensagem objMensagem
    objdadosCircuito.InformaBanco frmgenMDI_SGAPlus.objBanco
    
    objPendenciaDNI.InformaBanco frmgenMDI_SGAPlus.objBanco
    objPendenciaDNI.InformaMensagem objMensagem
             
    objDadosPendencia.InformaBanco frmgenMDI_SGAPlus.objBanco
    objDadosPendencia.InformaMensagem objMensagem
             
    'Verifica se a Fase ainda continua Em Andamento
    objDadosFase.InformaBanco frmgenMDI_SGAPlus.objBanco
    objDadosFase.InformaMensagem objMensagem
    
    If (Not objDadosFase.RecuperarSituacaoFase(Fase_ID)) And DataFinalizacao = "" Then

       objMensagem.ExibeMensagem

       GoTo Finalizar
    Else
       
       objMensagem.TipoMsg = ""
       objMensagem.Descricao = ""
    
    End If
    
    'Obtem o número da última fase realizada
    numUltimaFaseRealizada = objDadosFase.ObtemSequenciaUltimaFaseRealizada(Servico_ID, _
                                             GbCodigoFaseProgramarCPE, Familia_Produto)
    
    objAutenticacao.IDUsuario = frmgenMDI_SGAPlus.objAutenticacao.IDUsuario
    objAutenticacao.Login = frmgenMDI_SGAPlus.objAutenticacao.Login
    objAutenticacao.InformaMensagem objMensagem
    
    Call objAutenticacao.InformaBanco(frmgenMDI_SGAPlus.objBanco)

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
    
    
    If Familia_Produto = "D" And objdadosServico.DataRealizacao = "" Then
            
        bRetorno = VerificaUltimaVersaoContratoDNI(Servico_ID, objdadosServico.VersaoContrato, objMensagem)
        
        If Not bRetorno Then
        
            GoTo Finalizar
            
        End If
        
    End If
    
    'Seta o objeto mensagem da classe de dados
    objDadosAcesso.InformaMensagem objMensagem
    objDadosAcesso.InformaBanco frmgenMDI_SGAPlus.objBanco
    
    Set objForm = New frmgenTratarAcessoEmbratel
    
    objForm.ssTabAcesso.Tab = 0
    
    objForm.LblAcao_Status = Status_Fase
    
    objDadosProduto.InformaBanco frmgenMDI_SGAPlus.objBanco
    objDadosProduto.InformaMensagem objMensagem
    
    bRetorno = objDadosProduto.RecuperarDados(Produto_ID)
    
    objForm.lblProduto.Caption = objDadosProduto.Descricao
    objForm.Caption = objDadosProduto.Sigla & " - Aprovisionamento de Acesso"
    
    CarregarComboTipoAcessoEBT objForm.CboTipoAcessoEmbratel
    
    objForm.LblEstacaoEntrega.Caption = "Configuração       Embratel"
    
    'Carrega os Cabecalho
    If Familia_Produto = "I" Then
        
       objForm.UsrDNICabecalhoAcesso.Visible = False
       
        'IP VPN
        If (Produto_ID = GbIDProdutoBIV) Then
            
            objForm.CmdConsultarOE.Visible = True
            objForm.CmdConsultarOE.Left = 2865
            objForm.UsrBLFCabecalho_Acesso.Visible = False
            objForm.usrBIPVPNCabecalho.Visible = True
            
            Call CarregaCabecalho_BIV(objForm.usrBIPVPNCabecalho, objBlfCabecalho)
            
            objForm.UsrVIPEquipComplementares.Visible = False
        
        ElseIf (Produto_ID = GbIDProdutoVIP) Then
            
           Call CarregaCabecalho_VIPLINEIP(objForm.UsrVozBandaCabecalho, objBlfCabecalho)
            
            objForm.CmdConsultarOE.Visible = False
            objForm.FrmCaracpastaEqpto.Visible = False
            objForm.FrmEqptosBLF.Visible = False
            objForm.FrameRespInfra.Visible = False
            objForm.FrameProprietarioModem.Visible = False
            objForm.FrameProvedorInterurbano.Visible = False
            objForm.CmdAcessoContingente.Visible = False
            objForm.CmdAcessoPrincipal.Visible = False
            objForm.TxtEstacaoEntregaEbtConf.Visible = False
            objForm.CmdConsultarOE.Visible = False
            objForm.CmdRealocar.Visible = True
            objForm.ssTabAcesso.TabVisible(4) = True
            objForm.UsrVIPEquipComplementares.Visible = True
            objForm.LblEstacaoEntrega.Caption = "Embratel"
            objForm.TxtEstacaoEntregaEBT.Left = 120
            objForm.CmdCadastrarEqptoNGN.Visible = True
            objForm.CmdCadastrarEqptoNGN.Enabled = True
            objForm.CmdObterEnderecoIP.Visible = True
            objForm.FrameExecutarCentroLocal.Visible = True
            objForm.FrameClassificacaoEndereco.Visible = True
            objForm.LblFrameFlagProvedorIU.Visible = False
            objForm.FrameFlagProvedorIU.Visible = False
            
            CarregarComboClassifEndereco objForm.CboClassificacaoEndereco
            
            Set objDadosEqptoNGN = New VIP_Equipamento_NGN
            
            objDadosEqptoNGN.InformaMensagem objMensagem
            objDadosEqptoNGN.InformaBanco frmgenMDI_SGAPlus.objBanco
            
            sStatusSwitch = objDadosEqptoNGN.VerificaStatusSwitch(Fase_ID)
            objForm.UsrVIPDadosSwitch.DesabilitaCampos
            
            If sStatusSwitch = "A" Then
            
                objForm.CmdObterEnderecoIP.Caption = "Aguardando dados de Switch"
            
            ElseIf sStatusSwitch = "P" Then
                
                objForm.CmdObterEnderecoIP.Caption = "Dados de Switch - OK"
                
                objForm.CmdObterEnderecoIP.Enabled = False
                
'                bRetorno = objDadosEqptoNGN.RecuperarDados(Fase_ID)
'
'                If objMensagem.TipoMsg <> "" Then
'
'                   objMensagem.ExibeMensagem
'
'                   GoTo Finalizar
'
'                End If
                
                Set objOC_DadosSwitch = New clsVIPoc_CadastrarEqptoNGN
                
                bRetorno = objOC_DadosSwitch.CarregarDadosSwitch(objDadosEqptoNGN, Fase_ID, _
                                             objForm.UsrVIPDadosSwitch)
                
            End If
            
            'Verifica se existe motivo de devolucao pra ser exibido
                        
            lIDFase = objdadosServico.ObtemIDFase(Servico_ID, GbCodigoFaseConfiguracaoNGN)
            
            Set objDadosCtrlCorrecao = New clsGenDadosControleCorrecao
            
            objDadosCtrlCorrecao.InformaBanco frmgenMDI_SGAPlus.objBanco
            objDadosCtrlCorrecao.InformaMensagem objMensagem
            
            If objDadosCtrlCorrecao.RecuperarDados(lIDFase, "D") Then
                
                objForm.CmdMotivoDevol.Visible = True
                
            End If
            
            Set objDadosCtrlCorrecao = Nothing
        
        Else
            Call CarregaCabecalho_BLF(objForm.UsrBLFCabecalho_Acesso, objBlfCabecalho)
            
            objForm.CmdConsultarOE.Visible = False
            
            objForm.UsrVIPEquipComplementares.Visible = False
        
        End If
    
       'IP VPN
       'Obtem a Velocidade da classe objBlfCabecalho e não do cabeçalho
       'Antes de excluir as linhas com comentarios realizar testes com outros produtos
       If (Produto_ID = GbIDProdutoBLD Or _
                Produto_ID = GbIDProdutoVOIP) Then
            'Carregar a combo de Velocidade
            bRetorno = CarregarComboVelocidade(objDadosAcesso, _
                                               objBlfCabecalho.Velocidade, _
                                               "VELOCIDADE_BLD", objForm.CboVelocidadeAcesso)
       
       Else
       
            'Carregar a combo de Velocidade
            bRetorno = CarregarComboVelocidade(objDadosAcesso, _
                                               objBlfCabecalho.Velocidade, _
                                               "VELOCIDADE", objForm.CboVelocidadeAcesso)
       
 
            If objMensagem.TipoMsg <> "" Then
        
               objMensagem.ExibeMensagem
        
               Unload objForm
               Set objForm = Nothing
               
               GoTo Finalizar
        
            End If
      End If

    ElseIf Familia_Produto = "D" Then
        
        objForm.cmdPendencia.Visible = False
        
        objForm.UsrDNICabecalhoAcesso.Visible = True
        
        objForm.UsrBLFCabecalho_Acesso.Visible = False
        
        objForm.CmdTrataPendencia.Left = 2190
        
        objForm.CmdTrataPendencia.Visible = True
        objForm.CmdConsultarOE.Visible = True
        objForm.CmdConsultarOE.Left = 4095
        
        'Recupera dados do Cabecalho
        Call CarregaCabecalho_Diginet(objForm.UsrDNICabecalhoAcesso, Servico_ID, _
                                      Fase_ID, objMensagem, GbCodigoFaseAcesso)
    
        'Carregar a combo de Velocidade
        bRetorno = CarregarComboVelocidade_DNI(objDadosAcesso, _
                                           objForm.UsrDNICabecalhoAcesso.Velocidade, _
                                           "VELOCIDADE_DNI", objForm.CboVelocidadeAcesso)
    
    End If
    
    If objForm.FrmEqptosBLF.Visible = True Then
        
        bRetorno = CarregarCombos(objDadosAcesso, "EQUIPAMENTO", objForm.CboTipoEquipamento_1)
        
        bRetorno = CarregarCombos(objDadosAcesso, "EQUIPAMENTO", objForm.CboTipoEquipamento_2)
        
        bRetorno = CarregarCombos(objDadosAcesso, "EQUIPAMENTO", objForm.CboTipoEquipamento_3)
        
        If objMensagem.TipoMsg <> "" Then
           
           objMensagem.ExibeMensagem
           
           Unload objForm
           
           Set objForm = Nothing
           
           GoTo Finalizar
        
        End If
    
        objForm.CboTipoEquipamento_1.ListIndex = -1
    
        objForm.CboTipoEquipamento_2.ListIndex = -1
    
        objForm.CboTipoEquipamento_3.ListIndex = -1
    
        If Not bRetorno Then
           
           Unload objForm
           Set objForm = Nothing
           
           GoTo Finalizar
        
        End If
    
    End If
    
   'Carrega a combo de Provedor Urbano
    bRetorno = CarregarCombos(objDadosAcesso, "PROVEDOR", objForm.CboProvedorUrbano)

    If objMensagem.TipoMsg <> "" Then

       objMensagem.ExibeMensagem

       Unload objForm
       Set objForm = Nothing
       
       GoTo Finalizar

    End If

    If Not bRetorno Then

       Unload objForm
       Set objForm = Nothing

       GoTo Finalizar

    End If

    objForm.InformaClasseOC Me

    If objForm.FrameProvedorInterurbano.Visible Then
        'Carrega a combo de Provedor Interurbano
        bRetorno = CarregarCombos(objDadosAcesso, "PROVEDOR", objForm.CboProvedorInterurbano)
    
        If objMensagem.TipoMsg <> "" Then
    
           objMensagem.ExibeMensagem
    
           Unload objForm
           Set objForm = Nothing
    
           GoTo Finalizar
    
        End If
    
        If Not bRetorno Then
    
           Unload objForm
           Set objForm = Nothing
    
           GoTo Finalizar
    
        End If
    
    End If
    
    'Carregar a combo de Interface de Embratel
    bRetorno = CarregarCombos(objDadosAcesso, "INTERFACEEBT", objForm.CboPadraoInterfaceEBT)

    If objMensagem.TipoMsg <> "" Then

       objMensagem.ExibeMensagem
       
       Unload objForm
       Set objForm = Nothing

       GoTo Finalizar

    End If

    'Carregar a combo de Interface Cliente
    bRetorno = CarregarCombos(objDadosAcesso, _
                "INTERFACECLI", _
                objForm.CboInterfaceClienteMenor)

    If objMensagem.TipoMsg <> "" Then

       objMensagem.ExibeMensagem

       Unload objForm
       Set objForm = Nothing

       GoTo Finalizar

    End If

    If Not bRetorno Then

       Unload objForm
       Set objForm = Nothing

       GoTo Finalizar

    End If

    'Carregar a combo de Interface de Embratel
    bRetorno = CarregarCombos(objDadosAcesso, "INTERFACEEBT", objForm.CboInterfaceEBTMenor)

    If objMensagem.TipoMsg <> "" Then

       objMensagem.ExibeMensagem

       Unload objForm
       Set objForm = Nothing
       
       GoTo Finalizar

    End If

    If Not bRetorno Then

       Unload objForm
       Set objForm = Nothing

       GoTo Finalizar

    End If

    objForm.ssTabAcesso.Tab = 0
    objForm.FlControlaAlteracao = False
    
    objForm.CboTipoAcessoEmbratel.ListIndex = -1
    objForm.CboTipoAcessoTerceiro.ListIndex = -1
    
    'Caso exista Servico Anterior e a fase ainda não foi iniciada,
    'Obtem os dados da OTS anterior
    If (Servico_ID_Anterior <> 0) Then
       
        'Não resgata os dados da OTS anterior quando se altera o
        'endereço de instalação. IP VPN
       
        bResgataAnterior = True
        
        'Local de instalação da nova OTS
        If (Produto_ID = GbIDProdutoBIV) And _
            (objdadosServico.ID_ACAO = GbIDAcaoAlteracao Or _
                objdadosServico.ID_ACAO = GbIDAcaoCorrecaoAtivacao Or _
                objdadosServico.ID_ACAO = GbIDAcaoCorrecaoAlteracao) Then
            
            'OTS atual
            bRetorno = objdadosCircuito.RecuperarDados(Servico_ID)
            
            If (Not bRetorno) Then
               objMensagem.ExibeMensagem
        
               Unload objForm
               Set objForm = Nothing
               
               GoTo Finalizar
            End If
            
            'OTS atual
            sCNLAtual = objdadosCircuito.CNL_Ponta_A
            
            'OTS anterior
            bRetorno = objdadosCircuito.RecuperarDados(Servico_ID_Anterior)

            If (Not bRetorno) Then
               
               objMensagem.ExibeMensagem
        
               Unload objForm
               Set objForm = Nothing
               
               GoTo Finalizar
            End If
            
            'OTS anterior
            sCNLAnterior = objdadosCircuito.CNL_Ponta_A
            
            If (sCNLAtual <> sCNLAnterior) Then
            
                bResgataAnterior = False
                
            End If
            
        End If
              
       objDadosAcesso.CaracteristicaAcesso = "P"
       
       If (Not objDadosAcesso.FaseExistente(Fase_ID) And _
                bResgataAnterior) Then
           
           IdOTS = Servico_ID_Anterior
           
           'Problema identificado durante os testes do BLD
           If (Sigla_Acao = "DES" Or Sigla_Acao = "CAN") Then
                
                'GbCodigoFaseAcesso - ID da fase na tabela Tipo_Fase
                IDFase = objdadosServico.ObtemIDFase_Recursivo(Servico_ID_Anterior, _
                            GbCodigoFaseAcesso)
           Else
                'GbCodigoFaseAcesso - ID da fase na tabela Tipo_Fase
                IDFase = objdadosServico.ObtemIDFase(Servico_ID_Anterior, _
                        GbCodigoFaseAcesso)
           End If

       Else
       
          IdOTS = Servico_ID
          IDFase = Fase_ID
          
       End If
           
    Else
       
       IdOTS = Servico_ID
       IDFase = Fase_ID
       
    End If
        
    'Atribui ao objeto de dados a acao e a situacao da OTS selecionada
    
    objDadosAcesso.Acao = Sigla_Acao
    objForm.LblAcao_Status = Status_Fase
    
    objDadosAcesso.CaracteristicaAcesso = "P"
        
    'Obtem dados de Local de Instalacao
    Call LimpaCamposTela
    
    objForm.ssTabAcesso.TabEnabled(3) = True

    'Caso a fase não tenha sido iniciada e seja uma OTS de Desativacao ou Cancelamento,
    'Exibe os dados da OTS original
    
    If Not objDadosAcesso.FaseExistente(Fase_ID) Then
       
       TipoOperacao = "I"     'Indica que a fase não foi iniciada ainda
       
    Else
       
       TipoOperacao = "A"     'Indica que a fase já foi iniciada anteriormente
       
    End If

    bRetorno = RecuperaDadosAcesso("P", IdOTS, IDFase, FlagInclusao)
    
    'IP VPN
    If (Produto_ID = GbIDProdutoBLF Or _
            Produto_ID = GbIDProdutoBIV Or _
            Produto_ID = GbIDProdutoJAH) Then
        
        Set objDadosCadastrais = New clsblfDadosCadastrais
        
        objDadosCadastrais.InformaBanco frmgenMDI_SGAPlus.objBanco
        objDadosCadastrais.InformaMensagem objMensagem
        
        'Recupera dados de Local de Instalação
        sTabela_Interface = ""
        
        If (lProduto_ID = GbIDProdutoJAH) Then
            
            sTabela_Interface = "ATEN_LOCAL_INSTA"
            
        End If
        
        bRetorno = objDadosCadastrais.RecuperarDados(ContratoID, NrLocalInsta, "T", sTabela_Interface)
    
        If objMensagem.TipoMsg <> "" Then
            
            GoTo ErrorHandler
        
        End If
    
    
        'IP VPN
        If objDadosProduto.Sigla = "VN" Then
        
            sSiglaProduto = GbCodigo_Produto_IPVPN
            
        End If

'''
        Set objDadosAcessoSOE = New clsblfDadosAcesso
    
        objDadosAcessoSOE.InformaBanco frmgenMDI_SGAPlus.objBanco
        objDadosAcessoSOE.InformaMensagem objMensagem
        
        'Recupera Dados de Acesso
        sTabela_Interface = ""
        
        If (lProduto_ID = GbIDProdutoJAH) Then
            
            sTabela_Interface = "ATEN_ACESS_FACIL"
            sSiglaProduto = GbCodigo_Produto_JAH
            
        End If
        
        bRetorno = objDadosAcessoSOE.RecuperarDados(ContratoID, NrLocalInsta, sSiglaProduto, sTabela_Interface)
        
        If objMensagem.TipoMsg <> "" Then
        
            GoTo ErrorHandler
            
        End If
        
        objForm.LblNumeroEstudoViabilidade.Caption = objDadosAcessoSOE.EstudoViabilidade
        
        objForm.optSegurancaAcesso(0).Value = (objDadosAcessoSOE.SegurancaAcesso = "S")
        objForm.optSegurancaAcesso(1).Value = (objDadosAcessoSOE.SegurancaAcesso = "N")
        objDadosAcesso.Cd_Veloc = objDadosAcessoSOE.CdVelocidadeAcesso
        
        If Val(objDadosAcessoSOE.CdVelocidadeAcesso) <> 0 Then
           
           Call SelecionaCombo(objForm.CboVelocidadeAcesso, _
                               CLng(objDadosAcessoSOE.CdVelocidadeAcesso))
        
        End If
        
        'objForm.CboVelocidadeAcesso.Enabled = False
        
        Set objDadosAcessoSOE = Nothing
        
    'BLD
    ElseIf (Produto_ID = GbIDProdutoBLD Or _
                Produto_ID = GbIDProdutoVOIP) Then
        
        Set objDadosCadastrais = New clsbldDadosContrato
        
        objDadosCadastrais.InformaBanco frmgenMDI_SGAPlus.objBanco
        objDadosCadastrais.InformaMensagem objMensagem
        
        'Recupera dados de Local de Instalação
        bRetorno = objDadosCadastrais.RecuperarDados(objdadosServico.NumOE, objdadosServico.AnoOE)
    
        If objMensagem.TipoMsg <> "" Then
        
            GoTo ErrorHandler
            
        End If
    
    
        'IP VPN
        If objDadosProduto.Sigla = "VN" Then
    
            sSiglaProduto = "IVN"
        
        End If

        Set objDadosAcessoSOE = New clsbldDadosAcesso
    
        objDadosAcessoSOE.InformaBanco frmgenMDI_SGAPlus.objBanco
        objDadosAcessoSOE.InformaMensagem objMensagem
        
        'Recupera Dados de Acesso
        bRetorno = objDadosAcessoSOE.RecuperarDados(objdadosServico.NumOE, objdadosServico.AnoOE)
        
        If objMensagem.TipoMsg <> "" Then
        
            GoTo ErrorHandler
            
        End If
        
        objForm.LblNumeroEstudoViabilidade.Caption = objDadosAcessoSOE.EstudoViabilidade
        
        'objForm.optSegurancaAcesso(0).Value = (objDadosAcessoSOE.SegurancaAcesso = "S")
        'objForm.optSegurancaAcesso(1).Value = (objDadosAcessoSOE.SegurancaAcesso = "N")
        objDadosAcesso.Cd_Veloc = objDadosAcessoSOE.VelocidadeAcesso
        
        If objDadosAcessoSOE.VelocidadeAcesso <> "" Then
           
           objForm.CboVelocidadeAcesso.Text = objDadosAcessoSOE.VelocidadeAcesso
        
        End If
        
        'objForm.CboVelocidadeAcesso.Enabled = False
        
        Set objDadosAcessoSOE = Nothing
        
    ElseIf (Produto_ID = GbIDProdutoWID) Or _
           (Produto_ID = GbIDProdutoVPN_RA) Or _
           (Produto_ID = GbIDProdutoVPN_DIRECT) Then
        
        Set objDadosCadastrais = New clswidDadosCadastrais
        
        objDadosCadastrais.InformaBanco frmgenMDI_SGAPlus.objBanco
        objDadosCadastrais.InformaMensagem objMensagem
        
        'Recupera dados de Local de Instalação
        bRetorno = objDadosCadastrais.RecuperarDados(ContratoID, NrLocalInsta, "C")
    
        Set objDadosAcessoSOE = New clswidDadosAcesso
    
        objDadosAcessoSOE.InformaBanco frmgenMDI_SGAPlus.objBanco
        objDadosAcessoSOE.InformaMensagem objMensagem
        
        'Recupera Dados de Acesso
        bRetorno = objDadosAcessoSOE.RecuperarDados(ContratoID, NrLocalInsta)
        
        If objMensagem.TipoMsg <> "" Then
        
            GoTo ErrorHandler
            
        End If
        
        objForm.LblNumeroEstudoViabilidade.Caption = objDadosAcessoSOE.EstudoViabilidade
        
        Set objDadosAcessoSOE = Nothing
        
    ElseIf (Produto_ID = GbIDProdutoVIP) Then
        
        Set objVIPDadosEnder = New VIP_Aten_Endereco
        Set objVIPDadosContrato = New VIP_Aten_Contrato
        Set objVIPDadosAcesso = New VIP_Aten_Dados_Acesso
        
        objVIPDadosAcesso.InformaBanco frmgenMDI_SGAPlus.objBanco
        objVIPDadosAcesso.InformaMensagem objMensagem
        
        objVIPDadosEnder.InformaBanco frmgenMDI_SGAPlus.objBanco
        objVIPDadosEnder.InformaMensagem objMensagem
    
        objVIPDadosContrato.InformaBanco frmgenMDI_SGAPlus.objBanco
        objVIPDadosContrato.InformaMensagem objMensagem
        
        bRetorno = objVIPDadosContrato.RecuperarDados(ContratoID)
        
        If objMensagem.TipoMsg <> "" Then
            
           GoTo ErrorHandler
            
        End If
        
        bRetorno = objVIPDadosEnder.RecuperarDados(EnderecoID)
        
        If objMensagem.TipoMsg <> "" Then
            
           GoTo ErrorHandler
            
        End If
        
        bRetorno = objVIPDadosAcesso.RecuperarDados(ItemServicoID)
        
        If objMensagem.TipoMsg <> "" Then
            
           GoTo ErrorHandler
            
        End If
        
        objForm.LblNumeroEstudoViabilidade.Caption = objVIPDadosAcesso.Nr_Estudo_Viab
        
        objForm.LblLogradouro.Caption = objVIPDadosEnder.Logradouro
        objForm.LblBairroEndBR = objVIPDadosEnder.Bairro
        objForm.LblCEP = objVIPDadosEnder.CEP
        objForm.LblCidadeEndBR = objVIPDadosEnder.Cidade
        objForm.LblComplementoEndBR = objVIPDadosEnder.Complemento
        objForm.LblNumeroEndBR = ""  'o ATEN Não possui este campo
        objForm.LBlUFEndBR = objVIPDadosEnder.UF
        
        objForm.LblCNPJEndBR = objVIPDadosContrato.CNPJ
        objForm.LblInscrEstadual = objVIPDadosContrato.InscricaoEstadual
        objForm.LblInscrMunicEndBR = objVIPDadosContrato.InscricaoMunicipal
        objForm.LblNomeContatoTecnico = objVIPDadosContrato.ResponsavelTecnicoNome
        objForm.LblTelefoneContatoT1 = objVIPDadosContrato.ResponsavelTecnicoTelefone
        objForm.optSegurancaAcesso(0).Value = False
        objForm.optSegurancaAcesso(1).Value = True
    
    ElseIf Familia_Produto = "D" Then
        
'        objForm.optPropriedadeAcesso(0).Enabled = False
'        objForm.optPropriedadeAcesso(1).Enabled = False
        
        Set objDniDadosTecnico = New clsDNI_DadosTecnicos
        
        objDniDadosTecnico.InformaMensagem objMensagem
        objDniDadosTecnico.InformaBanco frmgenMDI_SGAPlus.objBanco
        
        bRetorno = objDniDadosTecnico.RecuperarDados(NumeroOE, _
                                                     AnoOE, _
                                                     NumeroItemOE, _
                                                     ContratoID)
        
        If bRetorno And Sigla_Acao <> "CAN" And Sigla_Acao <> "DES" And TipoOperacao = "I" Then
            
            With objDniDadosTecnico
                
                objForm.LBLPropriedadeAcessoContingente.Caption = .Seguranca_Acesso
                
                objForm.optSegurancaAcesso(0).Value = _
                                     (.Seguranca_Acesso <> "N" And .Seguranca_Acesso <> "")
                
                objForm.optSegurancaAcesso(1).Value = _
                                     (.Seguranca_Acesso = "N" Or .Seguranca_Acesso = "")
                
                If .Seguranca_Acesso <> "N" And .Seguranca_Acesso <> "" Then
                   
                   objDadosAcesso.SegurancaAcesso = "S"
                   
                Else
                   
                   objDadosAcesso.SegurancaAcesso = "N"
                
                End If
                   
                objDadosAcesso.PropridedadeAcesso = .Propriedade_Acesso
                
                If .Seguranca_Acesso = "" Then
                   
                   objForm.optSegurancaAcesso(1).Value = True
                
                End If
                
            End With
        
        End If
        
        'Carrega os Dados Vindos do SOE
        Set objDadosAnalisarOE = New clsDNI_DadosAnalisarOE
    
        objDadosAnalisarOE.InformaMensagem objMensagem
        objDadosAnalisarOE.InformaBanco frmgenMDI_SGAPlus.objBanco
        
        sAcaoAux = objDadosAnalisarOE.ObtemAcaoOrderEntry(Servico_ID)
        
        bRetorno = objDadosAnalisarOE.ObtemDadosAnalisarOE(ContratoID, _
                                                           objDadosProduto.Sigla, _
                                                           sAcaoAux)
        
        If objMensagem.TipoMsg <> "" Then
            
           GoTo ErrorHandler
            
        End If
                
        If Produto_ID = GbIDProdutoDiginet Then
        
            objForm.LblNumeroEstudoViabilidade.Caption = objDadosAnalisarOE.NrEstudoViab
        
        ElseIf Produto_ID = GbIDProdutoEmbratelIPL Then
            
            objForm.LblNumeroEstudoViabilidade.Caption = objDadosAnalisarOE.NrEstudoViabIPL
        
        End If
    
    End If
    
    If Familia_Produto = "I" And (Produto_ID <> GbIDProdutoVIP) Then
        
        objForm.LblLogradouro.Caption = objDadosCadastrais.Logradouro
        objForm.LblBairroEndBR = objDadosCadastrais.Bairro
    
        objForm.LblCEP = objDadosCadastrais.CEP
    
        objForm.LblCidadeEndBR = objDadosCadastrais.Cidade
        objForm.LblCNPJEndBR = objDadosCadastrais.CNPJ
        objForm.LblComplementoEndBR = objDadosCadastrais.Complemento
        objForm.LblInscrEstadual = objDadosCadastrais.InscricaoEstadual
        objForm.LblInscrMunicEndBR = objDadosCadastrais.InscricaoMunicipal
        objForm.LblNomeContatoTecnico = objDadosCadastrais.ResponsavelTecnicoNome
        objForm.LblNumeroEndBR = objDadosCadastrais.Numero
        objForm.LblTelefoneContatoT1 = objDadosCadastrais.ResponsavelTecnicoTelefone
        objForm.LBlUFEndBR = objDadosCadastrais.UF
    
    ElseIf Familia_Produto = "D" Then
    
        objForm.LblLogradouro.Caption = objDadosAnalisarOE.LogradouroEndBR
        objForm.LblBairroEndBR = objDadosAnalisarOE.BairroEndBR
        objForm.LblCidadeEndBR = objDadosAnalisarOE.CidadeEndBR
        objForm.LblCNPJEndBR = objDadosAnalisarOE.CNPJEndBR
        objForm.LblComplementoEndBR = objDadosAnalisarOE.ComplementoEndBR
        objForm.LblInscrEstadual = objDadosAnalisarOE.InscrEstadualEndBR
        objForm.LblInscrMunicEndBR = objDadosAnalisarOE.InscrMunicBR
        objForm.LblNomeContatoTecnico = objDadosAnalisarOE.NomeContatoTecnico
        objForm.LblNumeroEndBR = objDadosAnalisarOE.NumeroEndBR
        objForm.LblTelefoneContatoT1 = "(" & objDadosAnalisarOE.CodAreaContatoTec1 & ") " & _
                                       objDadosAnalisarOE.TelefoneContatoTec1 & _
                                       IIf(objDadosAnalisarOE.RamalContatoTec1 = "", "", _
                                           " - " & objDadosAnalisarOE.RamalContatoTec1)
        
        objForm.LBlUFEndBR = objDadosAnalisarOE.UFEndBR
        objForm.LblCEP = objDadosAnalisarOE.CEPEndBR
    
    End If
    
    Set objDadosCadastrais = Nothing
        
    'Verifica se existe alguma fase posterior a esta já iniciada, em caso afirmativo,
    'desabilitar os campos da tela os os botões de ações ou caso o usuário não possua
    'permissão de atualização
    If Familia_Produto = "D" Then
    
        'Problema identificado por Faquino em 18/12/2003
        If UBound(paPermissoes) = 0 Then

           'Uma fase posterior a esta já foi iniciada
           Call DesabilitaCamposTela
           
           objForm.CmdTrataPendencia.Enabled = False

        ElseIf (numUltimaFaseRealizada > Sequencia_Acao) Or _
            (paPermissoes(4, 1) = 0) Or (Not flUsuarioDoCF) Then
           
           'Uma fase posterior a esta já foi iniciada
           Call DesabilitaCamposTela
           
           objForm.CmdTrataPendencia.Enabled = False
           
        End If
    
    Else
        'Problema identificado por Faquino em 18/12/2003
        If UBound(paPermissoes) = 0 Then
    
           'Uma fase posterior a esta já foi iniciada
           Call DesabilitaCamposTela
    
        ElseIf (DataFinalizacao <> "" And PermiteExecucao <> "S" And _
                Produto_ID <> GbIDProdutoVIP) Or _
               (numUltimaFaseRealizada > Sequencia_Acao And PermiteExecucao <> "S" And _
                Produto_ID = GbIDProdutoVIP) Or _
               paPermissoes(4, 1) = 0 Or Not flUsuarioDoCF Then
           
           'Uma fase posterior a esta já foi iniciada
           Call DesabilitaCamposTela
        
        End If
    
    End If

    If Produto_ID = GbIDProdutoVIP Then
        
        Set objOc_TratarEqptoCompl = New clsInterOC_TratarEqptoCompl
        
        Call objOc_TratarEqptoCompl.CarregarUsrControl(objDadosAcesso, objForm, False)
    
    End If
    
    objForm.InformaClasseOC Me
        
    objForm.optSegurancaAcesso(0).Enabled = False
    objForm.optSegurancaAcesso(1).Enabled = False

    If FlagInclusao Then
        
        objForm.optCaracAcessoEmbratel(0).Value = True
        objForm.optCaracAcessoTerceiro(0).Value = True
        objForm.optCaracAcessoEqpto(0).Value = True
        objForm.optCaracAcessoGen(0).Value = True
        objForm.FlControlaAlteracao = False
        
        Set cColecao = New Collection
        
        objDadosPendencia.ColecaoPendencia = cColecao
        
        objPendencia = objPendenciaDNI
        
        'Caso o usuário não possua permissão de atualização
        'Problema identificado por Faquino em 18/12/2003
        If UBound(paPermissoes) = 0 Then
        
            Call DesabilitaCamposTela
        
        'Caso o usuário não possua permissão de atualização
        ElseIf paPermissoes(4, 1) = 0 Or Not flUsuarioDoCF Then
           
           Call DesabilitaCamposTela
        
        End If
        
        objForm.Show
        
        GoTo Finalizar
        
    Else
        
        'Popula os campos da tela
        Call PopulaCamposTela(objDadosAcesso)
        
    End If
    
    If TipoOperacao = "I" Or (Produto_ID = GbIDProdutoVIP) Then
    
       objForm.CmdAcessoContingente.Visible = False
       objForm.CmdAcessoPrincipal.Visible = False
    
    End If
    
    'Caso seja uma OTS de desativacao ou cancelamento, desabilita os campos da tela
    If Sigla_Acao = "DES" Or Sigla_Acao = "CAN" Then

       Call DesabilitaCamposTela

       objForm.CmdSalvar.Enabled = True
       objForm.CmdLimpar.Enabled = False

    End If
        
    objForm.InformaClasseOC Me
    
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
    
    If objdadosServico.DataRealizacao <> "" Then
       
       'Circuito já ativado
       Call DesabilitaCamposTela

    End If
    
    If objForm.optPropriedadeAcesso(0).Value = True Then
    
        If (Sigla_Acao = "ALT" Or Sigla_Acao = "ATG" Or Sigla_Acao = "ALG") _
            And objdadosServico.Origem = "SGA" Then
       
'            objForm.SpdDataAtivacaoAcessoEBT.Enabled = False
            
        End If
    
    ElseIf objForm.optPropriedadeAcesso(1).Value = True Then
    
        If (Sigla_Acao = "ALT" Or Sigla_Acao = "ATG" Or Sigla_Acao = "ALG") _
           And objdadosServico.Origem = "SGA" Then
        
'            objForm.SpdDataAtivacaoAcessoTerc.Enabled = False
'            objForm.SpdDataEntregaUrbano.Enabled = False
'            objForm.SpdDataEntregaInterUrbano.Enabled = False
            
        End If
        
    End If
    
    objForm.FlControlaAlteracao = False
    
    'Executa método para popular os campos do Controle Cabecalho
    CarregarForm = True
   
Finalizar:
    
    FlagRecuperando = False
    
    Set objAutenticacao = Nothing
    Set objDadosFase = Nothing
    Set objPendenciaDNI = Nothing
    Set objDadosProduto = Nothing
    Set objdadosCircuito = Nothing
    
    Exit Function
    
ErrorHandler:


    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                           vbCrLf & "Houve um erro de execução no método: " & vbCrLf & vbCrLf & "CarregarForm do Objeto Controle " & _
                          "clsGenOC_TratarAcessoEmbratel"
       
    objMensagem.ExibeMensagem
       
    GoTo Finalizar
       
End Function

Public Function GravarDados(FlagDescarregaForm As Boolean, Optional Flag_Atualiza_Pendencia As Boolean) As Boolean

    Dim objDadosFase                As New clsgenFase
    Dim ObjEqpto                    As clsGenEqptoComplementar
    Dim objDadosPropriedadeTerc     As clsGenPropriedadeTerceiros
    Dim objWEBIP                    As clsInterOC_TrataWebIP
    Dim objDadosAtivacaoFinal       As clsblf_AtivacaoFinalizacao
    Dim objDadosAnalisarOE          As clsDNI_DadosAnalisarOE
    Dim objOc_TrataEqptoCompl       As clsInterOC_TratarEqptoCompl
    Dim sSituacaoOTS                As String
    Dim bRetorno                    As Boolean
    Dim FlagIniciouTransacao        As Boolean
    Dim numUltimaFaseRealizada      As Long
    Dim bTemAcessoRedundante        As Boolean
    Dim lIdContingencia             As Long
    Dim lSeqUltimaFaseOTS           As Long
    Dim iQtEquipamento              As Integer
    Dim lIDEquip                    As Long
    Dim iSequenciaAcaoAux           As Integer
    Dim FlagTemPendenciaEmAberto    As Boolean
    Dim FlagFaseFinalizadaOK        As Boolean
    Dim FlagTemPendenciaCliente     As Boolean
    Dim FlagManterPeriodo           As Boolean
    
    Dim lIdFaseCentroLocal          As Long
    Dim lIDCentroFuncional          As Long
    Dim sCodigoCFCentroLocal        As String
    Dim sDesignacaoSOE              As String
    Dim bRetornoCriticaOk           As Boolean
    Dim bFlagNaoExecutaCentroLocal  As Boolean
    Dim objDadosDuracaoFase         As clsgenDuracaoFase
    
    On Error GoTo ErrorHandler
    
    ReDim vMensagemListaMsg(1)
    
    FlagIniciouTransacao = False
    
    GravarDados = False
    
    bRetornoCriticaOk = False
    
    If (Sigla_Acao <> "DES" And Sigla_Acao <> "CAN") Then
    
       bRetornoCriticaOk = VerificaFinalizacaoOK(objDadosAcesso)
       
    Else
        
        bRetornoCriticaOk = True
        
    End If
    
    bFlagNaoExecutaCentroLocal = False
    
    If Produto_ID = GbIDProdutoVIP Then
    
        If bRetornoCriticaOk And objForm.optExecCentroLocal(1).Value = True Then
            
            objMensagem.TipoMsg = "Q"
            objMensagem.Descricao = "A fase Tratar Centro Local não será executada! Confirma S/N? "
            
            If objMensagem.ExibeMensagem = "S" Then
            
               bFlagNaoExecutaCentroLocal = True
               
            Else
            
               bRetornoCriticaOk = False
               
            End If
        
        End If
        
    End If
    
    bRetorno = objdadosServico.RecuperarDados(Servico_ID)
    
    If (objMensagem.TipoMsg <> "") Then
       
       GoTo ErrorHandler
       
    End If
    
    sSituacaoOTS = objdadosServico.StatusServico
    
    If (sSituacaoOTS = "CAN" Or sSituacaoOTS = "CR") Then
       
       objMensagem.TipoMsg = "A"
       objMensagem.Descricao = "Esta OTS foi paralisada pelo SOE!"
       objMensagem.ExibeMensagem
       
       GoTo Finalizar
    
    End If

    objDadosFase.InformaBanco frmgenMDI_SGAPlus.objBanco
    objDadosFase.InformaMensagem objMensagem

    If (Not objDadosFase.RecuperarSituacaoFase(Fase_ID)) And _
       DataFinalizacao = "" Then

       Set objDadosFase = Nothing

       objMensagem.ExibeMensagem

       GravarDados = True

       GoTo Finalizar
    
    Else
       
       objMensagem.TipoMsg = ""
    
    End If

    'Checa se o produto pode alocar vários CFs para a fase Tratar Centro Local
    Set objDadosDuracaoFase = New clsgenDuracaoFase
    objDadosDuracaoFase.InformaBanco frmgenMDI_SGAPlus.objBanco
    objDadosDuracaoFase.InformaMensagem objMensagem

    'Obtem o número da última fase realizada
    numUltimaFaseRealizada = objDadosFase.ObtemSequenciaUltimaFaseRealizada(Servico_ID, _
                                         GbCodigoFaseProgramarCPE, Familia_Produto)

    
    If (objMensagem.TipoMsg <> "") Then
       
       objMensagem.ExibeMensagem
       GoTo Finalizar
    
    End If
    
    If (numUltimaFaseRealizada > Sequencia_Acao) And _
            (PermiteExecucao <> "S") Then
    
       objMensagem.TipoMsg = "A"
       objMensagem.Descricao = "Os dados desta fase não poderão ser alterados! Pois, a fase seguinte já foi iniciada."
       objMensagem.ExibeMensagem
       
       GravarDados = True
       
       GoTo Finalizar
    
    End If

    If Familia_Produto = "D" And objdadosServico.DataRealizacao = "" Then
            
        bRetorno = VerificaUltimaVersaoContratoDNI(Servico_ID, objdadosServico.VersaoContrato, objMensagem)
        
        If Not bRetorno Then
        
            GoTo Finalizar
            
        End If
        
    End If

    'Carrega as propriedades da classe de dados e executa o método de Inclusao ou alteracao
    
    'Popula dados da Pasta Dados Gerais
    
    If objForm.optCaracAcessoEmbratel(0).Value = True Then
       
       objDadosAcesso.CaracteristicaAcesso = "P"
    
    Else
       
       objDadosAcesso.CaracteristicaAcesso = "C"
    
    End If
    
    If objForm.optSegurancaAcesso(0).Value = True Then
       
       objDadosAcesso.SegurancaAcesso = "S"
    
    Else
       
       objDadosAcesso.SegurancaAcesso = "N"
    
    End If
    
    objDadosAcesso.eMailUsuario = TrataPlick(objForm.TxtEMailGic.Text)
    objDadosAcesso.RIT = objForm.TxtRIT.Text
    
    If objForm.CboPadraoInterfaceEBT.ListIndex > 0 Then
       
       objDadosAcesso.PadraoInterfaceEBT_ID = objForm.CboPadraoInterfaceEBT.ItemData( _
                                              objForm.CboPadraoInterfaceEBT.ListIndex)
    Else
       
       objDadosAcesso.PadraoInterfaceEBT_ID = 0
    
    End If
    
    If objForm.SpdDataPedidoAcesso.Text <> "" Then
        
        objDadosAcesso.DataPedidoAcesso = Format(objForm.SpdDataPedidoAcesso.Text, "dd/mm/yyyy")
    
    Else
    
        objDadosAcesso.DataPedidoAcesso = ""
    
    End If
    
    'Esta regra não á válida para OTS do produto BLD
    
    If Familia_Produto = "I" And _
            (Produto_ID <> GbIDProdutoBLD And _
                Produto_ID <> GbIDProdutoVOIP) Then
    
        If objForm.CboVelocidadeAcesso.ListIndex >= 0 Then
           
           objDadosAcesso.Cd_Veloc = Format(objForm.CboVelocidadeAcesso.ItemData( _
                                     objForm.CboVelocidadeAcesso.ListIndex), "00000")
        
        Else
           
           objDadosAcesso.Cd_Veloc = ""
        
        End If
    
    ElseIf Familia_Produto = "D" Then
    
        If objForm.CboVelocidadeAcesso.ListIndex > 0 Then
        
            objDadosAcesso.Cd_Veloc = vVelocidadeDNI(objForm.CboVelocidadeAcesso.ListIndex).CodigoAlpha
    
        Else
        
            objDadosAcesso.Cd_Veloc = ""
        
        End If
        
    End If
    
    objDadosAcesso.ProprietarioModem = ""
    
    If objForm.optProprietarioModem(0).Value = True Then
        
        objDadosAcesso.ProprietarioModem = "E"
    
    Else
    
        If objForm.optProprietarioModem(1).Value = True Then
          
           objDadosAcesso.ProprietarioModem = "C"
        
        End If
    
    End If
    
    If objForm.OptResponsavelInfra(0).Value = True Then
        
        objDadosAcesso.ResponsavelInfra = "E"
    
    Else
    
        If objForm.OptResponsavelInfra(1).Value = True Then
         
           objDadosAcesso.ResponsavelInfra = "C"
    
        End If
        
    End If
    
    If objForm.OptTipoRede(0).Value = True Then
        
        objDadosAcesso.TipoRede = "E"      'Rede E1
    
    Else
    
        If objForm.OptTipoRede(1).Value = True Then
         
           objDadosAcesso.TipoRede = "T"   'Tellabs
    
        End If
        
    End If
    
    'Corrigido por Mozart Simões em 06/10/2003
    'Acerto de lançamento de código de usuário nas alterações da
    'Fase de Tratamento de Acesso.
    'objDadosAcesso.UsuarioAtualizacao_ID = GbIDAcaoAlteracao
    objDadosAcesso.UsuarioAtualizacao_ID = frmgenMDI_SGAPlus.objAutenticacao.IDUsuario
    
    objDadosAcesso.FlagExecutaCentrolocal = ""
    
    If objForm.FrameExecutarCentroLocal.Visible = True Then
       
       objDadosAcesso.FlagExecutaCentrolocal = _
                                          IIf(objForm.optExecCentroLocal(0).Value = True, _
                                                                                   "S", "N")
    End If
    
    '-- JRCRUZ ------------------------------------------------------**
    If objForm.FrameClassificacaoEndereco.Visible = True Then
       objDadosAcesso.ClassificacaoEndereco = objForm.CboClassificacaoEndereco.ListIndex
    End If
    '-- JRCRUZ ------------------------------------------------------**
    
    If objForm.optPropriedadeAcesso(0).Value = True Then
       
       objDadosAcesso.PropridedadeAcesso = "E"
       
    Else
    
       objDadosAcesso.PropridedadeAcesso = "T"
    
    End If
    
    If objDadosAcesso.PropridedadeAcesso = "E" Then
       
       Set objDadosPropriedadeEBT = New clsGenPropriedadeEmbratel
       objDadosAcesso.objPropriedadeEBT = objDadosPropriedadeEBT
       Set objDadosPropriedadeEBT = Nothing
       
       Call CarregaAtributosPropriedadesEBT(objDadosAcesso)
    
    Else
       
       Set objDadosPropriedadeTerc = New clsGenPropriedadeTerceiros
       objDadosAcesso.objPropriedadeTerceiros = objDadosPropriedadeTerc
       Set objDadosPropriedadeTerc = Nothing
       
       Call CarregaAtributosPropriedadesTerceiros(objDadosAcesso.objPropriedadeTerceiros)
    
    End If
    
    'Carrega os atributos de Equipamentos Complementares
    If Produto_ID <> GbIDProdutoVIP Then   'Se vipline IP
        
        Call CarregaAtributosEqtos(objDadosAcesso)
    
    End If
    
    'Inicia a Transacao
    bRetorno = frmgenMDI_SGAPlus.objBanco.Iniciar_Transacao()
        
    If Not bRetorno Then
        
        GoTo ErrorHandler
    
    End If
    
    FlagIniciouTransacao = True
    
    'Atualiza dados de Acesso
    If objDadosAcesso.FaseExistente(Fase_ID) Then
       
       TipoOperacao = "A"
       
       Gravar = objDadosAcesso.AtualizarDados(Fase_ID)
    
    Else
       
       TipoOperacao = "I"
       
       Gravar = objDadosAcesso.IncluirDados(Fase_ID)
       
       objDadosAcesso.iD = objDadosAcesso.ObtemIDContingencia(Fase_ID, objDadosAcesso.CaracteristicaAcesso)
    
    End If
   
    If objMensagem.TipoMsg <> "" Then
       
       GoTo ErrorHandler
    
    End If
    
    If Produto_ID = GbIDProdutoVIP Then   'Se vipline IP
    
        Set objOc_TrataEqptoCompl = New clsInterOC_TratarEqptoCompl
        
        objOc_TrataEqptoCompl.InformaMensagem objMensagem
        
        Call objOc_TrataEqptoCompl.GravarDadosEqptos(objDadosAcesso.iD, objForm)
        
        'Obtem os dados de Equipamento Complementar
        bRetorno = objDadosAcesso.RecuperarDadosEquipamentos(objDadosAcesso.iD, objMensagem)
          
        If objMensagem.TipoMsg <> "" Then
            
           GoTo ErrorHandler
         
        End If
        
        Call objForm.UsrVIPEquipComplementares.CarregarGridEqpto(objDadosAcesso.objColecaoEqpto)
        
    End If
    
    If objMensagem.TipoMsg <> "" Then
       
       GoTo ErrorHandler
    
    End If
    
    'Caso Tenha gravado Acesso OK
    If (Gravar) Then
    
        If Produto_ID <> GbIDProdutoVIP Then  'Tratamento abaixo serve para todos os
                                              'produtos exceto VIPLINE IP

            'Grava dados de Equipamentos Complementares
            iQtEquipamento = 0
            
            For Each ObjEqpto In objDadosAcesso.objColecaoEqpto
                
                iQtEquipamento = iQtEquipamento + 1
                
                ObjEqpto.InformaBanco frmgenMDI_SGAPlus.objBanco
                
                If Val(ObjEqpto.TipoEquipamento) <> 0 Then
                    
                    If TipoOperacao = "I" Then
                    
                        ObjEqpto.iD = 0
                        
                    End If
                    
                    If ObjEqpto.iD = 0 Then 'Inclusao
                       
                       bRetorno = ObjEqpto.IncluirDados("E", objDadosAcesso.iD, objMensagem)
                     
                       If bRetorno Then
                          
                          lIDEquip = ObjEqpto.ObtemID(iQtEquipamento, objDadosAcesso.iD, objMensagem)
                       
                       End If
                       
                       Select Case iQtEquipamento
                              Case 1
                                    objForm.TxtIDEquipamento1.Text = lIDEquip
                              Case 2
                                    objForm.TxtIDEquipamento2.Text = lIDEquip
                              Case 3
                                    objForm.TxtIDEquipamento3.Text = lIDEquip
                       End Select
                       
                    Else
                                       
                       bRetorno = ObjEqpto.AtualizarDados("E", objDadosAcesso.iD, ObjEqpto.iD, objMensagem)
                                       
                    End If
                    
                Else
                
                    If TipoOperacao <> "I" And ObjEqpto.iD <> 0 Then  'Atualizacao
                    
                       bRetorno = ObjEqpto.ExcluirDados(objDadosAcesso.iD, ObjEqpto.iD, objMensagem)
                    
                    End If
                    
                End If
                
                If objMensagem.TipoMsg <> "" Then
                 
                   GoTo ErrorHandler
                
                End If
                
            Next
        
        End If
        
        If objDadosAcesso.PropridedadeAcesso = "E" Then 'Propriedade Embratel
                    
            'Grava dados Propriedade Embratel
            bRetorno = GravaDadosPropriedadeEmbratel(objDadosAcesso)
            
            If Not bRetorno Then
            
                GoTo ErrorHandler
            
            End If
        
        Else
        
            'Grava dados Propriedade Terceiros
            bRetorno = GravaDadosPropriedadeTerceiros(objDadosAcesso)
            
            If Not bRetorno Then
            
                GoTo ErrorHandler
            
            End If
        
        End If
    
        'Se existia Seguranca de acesso e não mais, o acesso Contingente deve ser excluido
        If SegurancaAcessoAnterior = "S" And objForm.optSegurancaAcesso(1).Value = True Then
        
          lIdContingencia = objDadosAcesso.ObtemIDContingencia(Fase_ID, "C")
       
           bRetorno = ExcluiAcesso("C", lIdContingencia)
        
           If objMensagem.TipoMsg <> "" Then
                
              GoTo ErrorHandler
             
           End If
        
        End If
    End If
    
    'Verifica se Existe pendencia a ser incluida
'    If Produto_ID = GbIDProdutoBLF Or Produto_ID = GbIDProdutoWID Then
    If (Familia_Produto = "I") Then
        
        If (objDadosPendencia.ColecaoPendencia.Count <> 0) Then
           
           bRetorno = objDadosPendencia.AtualizarDados(Fase_ID, Servico_ID, _
                                                       Produto_ID, FlagTemPendenciaEmAberto, _
                                                       objdadosServico.DiretorRegional, _
                                                       objdadosServico.NumOE, _
                                                       objdadosServico.AnoOE, _
                                                       objdadosServico.ItemOE)
           
           If (Not bRetorno) Then
              GoTo ErrorHandler
           End If
           
        End If
    
    ElseIf (Familia_Produto = "D") Then
    
        'Verifica se Existe pendencia a ser incluida
        If (objPendencia.Tipo_Pendencia_ID <> 0 And _
                Flag_Atualiza_Pendencia) Then
            
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
       
       bRetorno = objblfOC_SelecionarOTS.GravarAnotacaoGen(False)
         
       If (objMensagem.TipoMsg <> "") Then
       
          GoTo ErrorHandler
          
       End If
    
    End If
     
    If (TipoOperacao = "I" And _
        objDadosAcesso.CaracteristicaAcesso = "P") Then 'Indica que a fase está sendo gravada pela 1a vez
       
       bRetorno = objDadosFase.IniciaFase(Fase_ID)
       
       If (objMensagem.TipoMsg <> "") Then
       
          GoTo ErrorHandler
          
       End If
       
    End If
    
    FlagFaseFinalizadaOK = False
    
    'Verifica se a fase pode ser finalizada
    
'    If (Sigla_Acao <> "DES" And Sigla_Acao <> "CAN") Then
'
'       bRetorno = VerificaFinalizacaoOK(objDadosAcesso)
'
'    Else
'
'        bRetorno = True
'
'    End If
        
    If (bRetornoCriticaOk And FlagTemPendenciaEmAberto = False) Then

       If (DataFinalizacao = "") Then
       
          'Para OTS de Cancelamento e Desativacao, não será necessário calcular o periodo
          'para a fase seguinte já que todas as fases após a fase de Alocacao, são executadas
          'em paralelo
          
          If (Sigla_Acao <> "CAN" And Sigla_Acao <> "DES") And _
                (Familia_Produto = "I") Then
                
                'gerar servico de desativação
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
                        
                    End If  'sDesignacaoAtual
                    
                End If  'objdadosServico.ID_ACAO
                'fim

                If (Produto_ID <> GbIDProdutoJAH) And (Produto_ID <> GbIDProdutoVIP) Then
                
                    'Tratar Centro Local
                    'Uma OTS de cancelamento pode não gerar Fase Tratar Centro Local Aquino IP VPN
                    'Está fora do bloco que checa várias fases pois o retorno é verificado
                    'no Else
                    lIdFaseCentroLocal = objDadosFase.ObtemIDFase(Servico_ID, _
                                GbCodigoFaseCentroLocal)
            
                    If (objMensagem.TipoMsg <> "") Then
                    
                        GoTo ErrorHandler
                        
                    End If
            
                    'Checa se o produto pode alocar vários CFs para a fase Tratar Centro Local
                    If objDadosDuracaoFase.ChecaPermissaoMultiplosCF(Produto_ID, _
                                GbCodigoFaseCentroLocal) Then
            
                        'Verifica se a fase Centro local já foi finalizada,
                        'Caso o método retorne 0, indica que a fase Centro Local ainda não foi finalizada
                        If objDadosFase.ObtemSequenciaFaseFinalizada(Servico_ID, _
                                    GbCodigoFaseCentroLocal) = 0 And _
                                    (lIdFaseCentroLocal > 0) Then
                    
                            'Carrega a tela Alocação de Centro Funcional e atualiza os CFs alocados
                            bRetorno = CarregarFormCFsDaFase(lIdFaseCentroLocal)
                              
                            If (Not bRetorno) Then
                            
                                GoTo ErrorHandler
                                
                            End If
            
                        End If 'ObtemSequenciaFaseFinalizada
                        
                    End If  'ChecaPermissaoMultiplosCF
                
                End If  '(Produto_ID <> GbIDProdutoJAH)
        
                'Caso seja uma Desativacao ou Cancelamento, a fase de Configuração será executada
                'em paralelo com a fase de Acesso, portanto a atualização da tabela de Integração,
                'será feita ao final da fase de Alocação de Centro Funcional
                
                If (Produto_ID <> GbIDProdutoVIP) Then
                    
                    If (objWEBIP Is Nothing) Then
                    
                        Set objWEBIP = New clsInterOC_TrataWebIP
                        
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
                
                End If
                
                iSequenciaAcaoAux = Sequencia_Acao
                'Verifica se fase Tratar Centro Local deve ser executada ou não.
                
                If (Produto_ID = GbIDProdutoVIP) Then
                    
                    'Caso o usuario opte por não executar a fase Tratar Centro Local
                    If objForm.optExecCentroLocal(1).Value = True And _
                       bFlagNaoExecutaCentroLocal Then
                        
                        lIdFaseCentroLocal = objDadosFase.ObtemIDFase(Servico_ID, _
                                                                     GbCodigoFaseCentroLocal)
                        
                        If lIdFaseCentroLocal <> 0 Then
                           
                           bRetorno = objDadosFase.RetirarPeriodoPrevisto(lIdFaseCentroLocal)
                        
                        End If
                        
                        'Obter a sequencia acao da fase posterior a fase Centro Local
                        iSequenciaAcaoAux = objDadosFase.ObtemSequenciadaFase(Servico_ID, _
                                                                   GbCodigoFaseCentroLocal)
                    
                    End If
                    
                End If
                
                
                'Calcula o período para a proxima fase
                bRetorno = objDadosFase.CalculaPeriodoProximaFase(iSequenciaAcaoAux, Servico_ID)
                
                If (objMensagem.TipoMsg <> "") Then
                  
                  GoTo ErrorHandler
                
                End If
            
        ElseIf (Familia_Produto = "D") Then
            
            'Calcula o período para a proxima fase
            bRetorno = objDadosFase.CalculaPeriodoProximaFase(Sequencia_Acao, _
                                                              Servico_ID, "D")
          
            If (objMensagem.TipoMsg <> "") Then
            
                GoTo ErrorHandler
                
            End If
            
        End If
          
          'Quando for uma OTS de Cancelamento ou Desativacao, verifica se
          'as fases que estão sendo executadas em paralelo já foram finalizadas
          'caso Todas as outras fases já estejam finalizadas, o sistema deverá finalizar
          'o Servico.
          
          If (Sigla_Acao = "CAN" Or Sigla_Acao = "DES") And _
                Familia_Produto = "I" Then

                bRetorno = objdadosServico.RecuperarDados(Servico_ID)
             
                If (Not bRetorno) Then
                   GoTo ErrorHandler
                End If
                
                bRetorno = VerificaFinalizacaoDesativacao(objDadosFase, _
                                                          objdadosServico, _
                                                          Servico_ID, Fase_ID, _
                                                          Sigla_Acao, _
                                                          objMensagem)
                 If (objMensagem.TipoMsg <> "") Then
                    
                    GoTo ErrorHandler
                    
                 End If
          
          End If
          
          'Preenche o periodo realizado indicando que a fase foi finalizada
          bRetorno = objDadosFase.FinalizaFase(Fase_ID)
        
          If (objMensagem.TipoMsg <> "") Then
             
             GoTo ErrorHandler
             
          End If
          
          If (Familia_Produto = "D") Then
              
              'Obtem a sequencia_acao da última fase existente para a OTS
              lSeqUltimaFaseOTS = objDadosFase.ObtemSequenciaUltimaFase(Servico_ID, _
                                                                        Familia_Produto)

              If (objMensagem.TipoMsg <> "") Then
            
                 objMensagem.ExibeMensagem
                 
                 GoTo Finalizar
            
              End If
              
              If (Sequencia_Acao = lSeqUltimaFaseOTS) Then
                 
                 Set objDadosAnalisarOE = New clsDNI_DadosAnalisarOE

                 objDadosAnalisarOE.InformaMensagem objMensagem
                 objDadosAnalisarOE.InformaBanco frmgenMDI_SGAPlus.objBanco
               
                 sDesignacaoSOE = Trim(objForm.UsrDNICabecalhoAcesso.CNLPontaA) & "/" & _
                                  Trim(objForm.UsrDNICabecalhoAcesso.CNLPontaB) & "/" & _
                                  Trim(objForm.UsrDNICabecalhoAcesso.CodigoNP) & "/"
                                  
                 If Produto_ID = GbIDProdutoDiginet Then
               
                   sDesignacaoSOE = sDesignacaoSOE & _
                                Format(objForm.UsrDNICabecalhoAcesso.SequencialDesig, "0000")
               
                 Else
                   
                   sDesignacaoSOE = sDesignacaoSOE & _
                            Format(objForm.UsrDNICabecalhoAcesso.SequencialDesig, "0000000")
               
                 End If

                 bRetorno = objDadosAnalisarOE.EncerraWorkFlow(Servico_ID, _
                                                     Servico_ID_Anterior, _
                                                     AnoOE, NumeroOE, NumeroItemOE, _
                                                     "02", GbUsuario, _
                                                     sDesignacaoSOE, _
                                                     Sigla_Acao, _
                                                     objdadosServico)
              
                Set objDadosAnalisarOE = Nothing
                
                If Sigla_Acao = "DES" Or Sigla_Acao = "CAN" Then
                
                   objdadosServico.FinalizaServico (Servico_ID)
                
                End If
                
              End If
          
          End If
          
       End If
       
       FlagFaseFinalizadaOK = True
       
    Else
       
       'Caso a fase já tenha sido finalizada e uma pendencia tenha sido incluida,
       'A fase deverá ser considerada como Em Andamento
       If (DataFinalizacao <> "") Then

          objDadosFase.ID_Servico = Servico_ID
          objDadosFase.Id_Fase = Fase_ID
                 
          'Quando for OTS de Desativacao ou Cancelamento, o periodo previsto para a fase seguinte
          'deverá ser mantido já que as fases para estas ações serão executadas em paralelo.
         
          If Sigla_Acao = "DES" Or Sigla_Acao = "CAN" Then
         
             FlagManterPeriodo = True
             
          Else
         
             FlagManterPeriodo = False
         
         End If
        
         bRetorno = objDadosFase.RetirarFinalizacaoFase(Servico_ID, Fase_ID, _
                                       Sequencia_Acao, FlagManterPeriodo, Familia_Produto)
             
         If objMensagem.TipoMsg <> "" Then
               
            Set objDadosFase = Nothing
            
            GoTo ErrorHandler
             
         End If
       
       End If
    
       DataFinalizacao = ""
       
    End If

    bRetorno = frmgenMDI_SGAPlus.objBanco.Finalizar_Transacao
    
    If Not bRetorno Then
    
        GoTo ErrorHandler
        
    End If
    
    GravarDados = True
    
    bFlagControlaAlteracao = False
    objForm.FlControlaAlteracao = False
    objForm.UsrVIPEquipComplementares.bFlagAlteracao = False
    
    If objDadosAcesso.SegurancaAcesso = "S" Then 'Seguranca de Acesso
          
          If objDadosAcesso.CaracteristicaAcesso = "P" Then 'Acesso Principal
                          
             
             If FlagFaseFinalizadaOK = True Then

                 objMensagem.Descricao = "Os Dados do Acesso Principal foram " & _
                                        "atualizados com Sucesso!"
                
                 FlagDescarregaForm = True
            
'#########################
'adeil RN-SGAD-81746VNS 07/10/2003
'inicio

            ElseIf Not FlagFaseFinalizadaOK And Not FlagTemPendenciaEmAberto Then

                objMensagem.TipoMsg = "A"
                
                objMensagem.Descricao = "Os Dados do Acesso Principal foram atualizados com Sucesso. " & vbCrLf & _
                               " Porém a fase continua em Andamento," & vbCrLf & _
                               " pois campos obrigatórios não foram preenchidos."
    
    
                 FlagDescarregaForm = False
                 
                 If Not objForm.CmdAcessoPrincipal.Visible And (Produto_ID <> GbIDProdutoVIP) Then
                 
                    objForm.CmdAcessoContingente.Visible = True
                    objForm.CmdAcessoPrincipal.Visible = True
                 
                 End If
    
    
            ElseIf Not FlagFaseFinalizadaOK And FlagTemPendenciaEmAberto Then

                objMensagem.TipoMsg = "A"

                objMensagem.Descricao = "Os Dados do Acesso Principal foram atualizados com Sucesso. " & vbCrLf & _
                               " Porém a fase continua em andamento, devido a existência de pendência!"
'fim
'########### ##############
                
                 FlagDescarregaForm = False
                 
                 If (Not objForm.CmdAcessoPrincipal.Visible) And (Produto_ID <> GbIDProdutoVIP) Then
                 
                    objForm.CmdAcessoContingente.Visible = True
                    objForm.CmdAcessoPrincipal.Visible = True
                 
                 End If
             
             End If
             
'''             objMensagem.ExibeMensagem
'''             objMensagem.TipoMsg = ""
'''             objMensagem.Descricao = ""
             
             objForm.ssTabAcesso.Tab = 0
          
          Else
             
             objMensagem.TipoMsg = "A"
             
             If (FlagFaseFinalizadaOK = True) Then
             
                FlagDescarregaForm = True
                       
                objMensagem.Descricao = "Os Dados do Acesso Contingente foram atualizados com Sucesso!"
                       
             Else
             
                objMensagem.Descricao = _
                                   vbCrLf & vbCrLf & _
                                   "Os Dados do Acesso Contingente foram atualizados " & _
                                  "com Sucesso! " & vbCrLf & _
                                  "Porém a fase continua em Andamento."
             
                FlagDescarregaForm = False
                
             End If
             
             'objMensagem.ExibeMensagem
          
          End If
          
    Else
            
        objMensagem.TipoMsg = "A"
        
        If FlagFaseFinalizadaOK = True Then
        
            objMensagem.Descricao = "Os Dados do Acesso foram atualizados com Sucesso!"
            
            FlagDescarregaForm = True
           
'#########################
'adeil RN-SGAD-81746VNS 07/10/2003
'inicio

        ElseIf Not FlagFaseFinalizadaOK And Not FlagTemPendenciaEmAberto Then

            objMensagem.TipoMsg = "A"
            
            objMensagem.Descricao = "Os Dados do Acesso Principal foram atualizados com Sucesso. " & vbCrLf & _
                           " Porém a fase continua em Andamento," & vbCrLf & _
                           " pois campos obrigatórios não foram preenchidos."

            FlagDescarregaForm = False

        ElseIf Not FlagFaseFinalizadaOK And FlagTemPendenciaEmAberto Then

            objMensagem.TipoMsg = "A"

            objMensagem.Descricao = "Os Dados do Acesso Principal foram atualizados com Sucesso. " & vbCrLf & _
                           " Porém a fase continua em andamento, devido a existência de pendência!"
                           
            FlagDescarregaForm = False
'fim
'#########################
        End If
        'objMensagem.ExibeMensagem
        
    End If
    
        'Trata ativação da OTS de Topologia pelo SAIP. IP VPN
        'IP VPN
        'Inicia a próxima Fase apenas se a OTS de Topologia estiver Ativada no SAIP.
'''        If (FlagFaseFinalizadaOK = True And _
'''                Produto_ID = GbIDProdutoBIV And _
'''                FlagFaseAguardaSAIPAtivarTopologia) Then
'''
'''                objMensagem.Descricao = objMensagem.Descricao & _
'''                    vbCrLf & vbCrLf & "Esta OTS estará disponível na fase Alocação/Configuração " & _
'''                                        "apenas após Ativação pelo SAIP da OTS de Topologia!"
'''        End If
        
    objMensagem.ExibeMensagem
    objMensagem.TipoMsg = ""
    objMensagem.Descricao = ""


Finalizar:
    Set objWEBIP = Nothing
    Set objDadosDuracaoFase = Nothing
    
    Set objDadosFase = Nothing
    Set ObjEqpto = Nothing
    Set objDadosPropriedadeTerc = Nothing
    Set objDadosAtivacaoFinal = Nothing
    Set objDadosAnalisarOE = Nothing
    
    Exit Function
    
ErrorHandler:
       
    ' Cancelar Transação

    If FlagIniciouTransacao Then
       
       bRetorno = frmgenMDI_SGAPlus.objBanco.Cancelar_Transacao
    
    End If
    
    If objMensagem.TipoMsg = "" Then
       
       objMensagem.TipoMsg = "E"
       objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                   "Houve um erro de execução no Método: " & vbCrLf & vbCrLf & "GravarDados " & _
                   "da classe clsGenOC_TratarAcessoEmbratel"
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

Private Function ExcluiAcesso(pCaracteristicaAcesso As String, plIdContingencia As Long) As Boolean

    Dim objAcesso               As clsGenAcesso
    Dim objEquipamentos         As clsGenEqptoComplementar
    Dim objDadosPropriedadeEBT  As clsGenPropriedadeEmbratel
    Dim objDadosPropriedadeTerc As clsGenPropriedadeTerceiros
    
    On Error GoTo ErrorHandler
    
    Set objAcesso = New clsGenAcesso
    
    objAcesso.InformaBanco frmgenMDI_SGAPlus.objBanco
    objAcesso.InformaMensagem objMensagem
    
    bRetorno = objAcesso.RecuperarDados(pCaracteristicaAcesso, Servico_ID, Fase_ID, objMensagem)
    
    If objMensagem.TipoMsg <> "" Then
        
       GoTo ErrorHandler
     
    End If
    
    'Exclui dados de Acesso
    If objAcesso.PropridedadeAcesso = "E" Then      'Propriedade Embratel
    
        'Obtem o tipo de Acesso
        Set objDadosPropriedadeEBT = New clsGenPropriedadeEmbratel
        
        objDadosPropriedadeEBT.InformaBanco frmgenMDI_SGAPlus.objBanco
        
        bRetorno = objDadosPropriedadeEBT.RecuperarDados(plIdContingencia, objMensagem)
        
        If objMensagem.TipoMsg <> "" Then
            
           GoTo ErrorHandler
         
        End If
        
        'Exclui o Tipo de Acesso
        bRetorno = ExcluiTipoAcessoEBT(objDadosPropriedadeEBT.TipoAcesso, plIdContingencia, objMensagem)
       
        If objMensagem.TipoMsg <> "" Then
            
           GoTo ErrorHandler
         
        End If
       
        'Exclui a Propriedade Embratel
        Retorno = objDadosPropriedadeEBT.ExcluirDados(plIdContingencia, objMensagem)
        
        If objMensagem.TipoMsg <> "" Then
            
           GoTo ErrorHandler
         
        End If
        
        Set objDadosPropriedadeEBT = Nothing
    
    Else
    
        'Obtem o tipo de Acesso
        Set objDadosPropriedadeTerc = New clsGenPropriedadeTerceiros
        
        objDadosPropriedadeTerc.InformaBanco frmgenMDI_SGAPlus.objBanco
        
        bRetorno = objDadosPropriedadeTerc.RecuperarDados(plIdContingencia, objMensagem)
        
        If objMensagem.TipoMsg <> "" Then
            
           GoTo ErrorHandler
         
        End If
    
        'Exclui o tipo de Acesso
         bRetorno = ExcluiTipoAcessoTerceiros(objDadosPropriedadeTerc.TipoAcesso, plIdContingencia, objMensagem)
         
         If objMensagem.TipoMsg <> "" Then
            
            GoTo ErrorHandler
          
         End If
         
         'Exclui propriedade Terceiro
         bRetorno = objDadosPropriedadeTerc.ExcluirDados(plIdContingencia, objMensagem)
    
         If objMensagem.TipoMsg <> "" Then
            
            GoTo ErrorHandler
          
         End If
    
    End If
       
    'Exclui os registros de Equipamentos
    
    Set objEquipamentos = New clsGenEqptoComplementar
    
    objEquipamentos.InformaBanco frmgenMDI_SGAPlus.objBanco
    
    bRetorno = objEquipamentos.ExcluirDados(plIdContingencia, 0, objMensagem)
    
    Set objEquipamentos = Nothing
    
    If objMensagem.TipoMsg <> "" Then
        
       GoTo ErrorHandler
     
    End If
    
    bRetorno = objDadosAcesso.ExcluirDados(plIdContingencia)
   
    If objMensagem.TipoMsg <> "" Then
       
       GoTo ErrorHandler
    
    End If

Finalizar:

     On Error GoTo 0
     
     Exit Function
     
ErrorHandler:
       
    If objMensagem.TipoMsg = "" Then
       
       objMensagem.TipoMsg = "E"
       objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                   "Houve um erro de execução no Método: " & vbCrLf & vbCrLf & "ExcluiAcesso " & _
                   "da classe clsGenOC_TratarAcessoEmbratel"
    End If
    
    GoTo Finalizar

End Function

Public Property Get Sigla_Acao() As String

Sigla_Acao = sSigla_Acao

End Property

Public Property Let Sigla_Acao(ByVal sNewValue As String)

sSigla_Acao = sNewValue

End Property

Public Property Get Status_Fase() As String

Status_Fase = sStatus_Fase

End Property

Public Property Let Status_Fase(ByVal vNewValue As String)

sStatus_Fase = vNewValue

End Property
Private Sub DesabilitaCamposTela()

    objForm.CmdTrataPendencia.Enabled = False
    objForm.optProprietarioModem(0).Enabled = False
    objForm.optProprietarioModem(1).Enabled = False
    objForm.SpdDataPedidoAcesso.Enabled = False
    objForm.CboPadraoInterfaceEBT.Enabled = False
    objForm.OptResponsavelInfra(0).Enabled = False
    objForm.OptResponsavelInfra(1).Enabled = False
    
    objForm.OptTipoRede(0).Enabled = False
    objForm.OptTipoRede(1).Enabled = False
    
    objForm.CboTipoAcessoEmbratel.Enabled = False
    objForm.TxtDesignacaoAcessoEBT.Enabled = False
    objForm.TxtIdentificadorAcessoEBT.Enabled = False
    objForm.TxtEstacaoEntregaEbtConf.Enabled = False
    objForm.TxtEstacaoEntregaEBT.Enabled = False
    objForm.TxtEstacaoCliente.Enabled = False
    
    objForm.SpdDataAtivacaoAcessoEBT.Enabled = False
    objForm.TxtDIDFila_Radio.Enabled = False
    objForm.TxtDIDArmacao_Radio.Enabled = False
    objForm.TxtDIDRegua_Radio.Enabled = False
    objForm.TxtDIDPosicao_Radio.Enabled = False
    objForm.TxtNumeroRepeticoes.Enabled = False
    objForm.TxtDIDFila_Fibra.Enabled = False
    objForm.TxtDIDArmacao_Fibra.Enabled = False
    objForm.TxtDidRegua_Fibra.Enabled = False
    objForm.TxtDIDPosicao_Fibra.Enabled = False
    objForm.TxtNumeroCaboFibra.Enabled = False
    objForm.TxtNumeroParFibra.Enabled = False
    objForm.TxtPredioEBT_ADE.Enabled = False
    objForm.TxtNumeroCabo_ADE.Enabled = False
    objForm.TxtNumeroPar_ADE.Enabled = False
    objForm.TxtPredioCabo_Cabo.Enabled = False
    objForm.TxtNumeroParEBT_Cabo.Enabled = False
    objForm.TxtNumeroCaboEBT_Cabo.Enabled = False
    objForm.TxtNumeroParCliente_Cabo.Enabled = False
    objForm.TxtNumeroCaboCliente_Cabo.Enabled = False
    
    objForm.optSegurancaAcesso(0).Enabled = False
    objForm.optSegurancaAcesso(1).Enabled = False
    
    objForm.optPropriedadeAcesso(1).Enabled = False
    objForm.optPropriedadeAcesso(0).Enabled = False
    
    objForm.optCaracAcessoEmbratel(1).Enabled = False
    objForm.optCaracAcessoTerceiro(1).Enabled = False
    
    objForm.CboTipoAcessoEmbratel.Enabled = False
    objForm.CboTipoAcessoTerceiro.Enabled = False
   
    objForm.optCaracAcessoGen(1).Enabled = False

    objForm.TxtNumeroOtsTx_EBT.Enabled = False
    objForm.TxtItemOtsTxEBT.Enabled = False

    objForm.CboTipoAcessoTerceiro.Enabled = False
    objForm.TxtEstacaoEntregaTercConf.Enabled = False
    objForm.TxtEstacaoEntregaTerc.Enabled = False
    objForm.TxtDesignacaoAcessoTerceiro.Enabled = False
    objForm.SpdNumeroAcessoUrbano.Enabled = False
    objForm.TxtCNLPontaA_Urbano.Enabled = False
    objForm.SpdDataEntregaUrbano.Enabled = False
    objForm.TxtNumeroOTSTXTerc.Enabled = False
    objForm.TxtItemOTSTxTerc.Enabled = False
    objForm.TxtIDentificadorAcessoUrbano.Enabled = False
    objForm.SpdDataAtivacaoAcessoTerc.Enabled = False
    
    objForm.optTipoTecnologia(0).Enabled = False
    objForm.optTipoTecnologia(1).Enabled = False
    objForm.optTipoTecnologia(2).Enabled = False
    
    objForm.CboProvedorUrbano.Enabled = False
    objForm.optExisteAcessoIU(0).Enabled = False
    objForm.optExisteAcessoIU(1).Enabled = False
    objForm.CboProvedorInterurbano.Enabled = False
    objForm.TxtIdentificadorAcessoEBTInterUrbano.Enabled = False
    objForm.SpdNumeroAcessoInterurbano.Enabled = False
    objForm.TxtCNLPontaA.Enabled = False
    objForm.TxtCNLPontaB.Enabled = False
    objForm.SpdDataEntregaInterUrbano.Enabled = False
    objForm.CboInterfaceEBTMenor.Enabled = False
    objForm.CboInterfaceClienteMenor.Enabled = False
    objForm.TxtNumeroCaboAcessoMenor.Enabled = False
    objForm.TxtNumeroParAcessoMenor.Enabled = False
    objForm.TxtBlocoDGMenor.Enabled = False
    objForm.TxtParDGMenor.Enabled = False
    objForm.TxtBlocoTransdataMenor.Enabled = False
    objForm.TxtParTransdataMenor.Enabled = False
    objForm.TxtDIDArmacao.Enabled = False
    objForm.TxtDIDFila.Enabled = False
    objForm.TxtDIDPosicao.Enabled = False
    objForm.TxtDIDRegua.Enabled = False
    objForm.TxtDominio.Enabled = False
    objForm.txtNo.Enabled = False
    
    objForm.CboPorta_Rede.Enabled = False
    objForm.CboSlot_Rede.Enabled = False
    objForm.CboTSI_Rede.Enabled = False
    objForm.CboTSF_Rede.Enabled = False
    
    objForm.TxtDIDArmacao_Rede.Enabled = False
    objForm.TxtDIDFila_Rede.Enabled = False
    objForm.TxtDIDPosicao_Rede.Enabled = False
    objForm.TxtDIDRegua_Rede.Enabled = False
    
    objForm.CboTipoEquipamento_1.Enabled = False
    objForm.TxtFabricanteEquipamento_1.Enabled = False
    objForm.optLocalInstalacao_1(0).Enabled = False
    objForm.optLocalInstalacao_1(1).Enabled = False
    
    objForm.CboTipoEquipamento_2.Enabled = False
    objForm.TxtFabricanteEquipamento_2.Enabled = False
    objForm.optLocalInstalacao_2(0).Enabled = False
    objForm.optLocalInstalacao_2(1).Enabled = False
    
    objForm.CboTipoEquipamento_3.Enabled = False
    objForm.TxtFabricanteEquipamento_3.Enabled = False
    objForm.optLocalInstalacao_3(0).Enabled = False
    objForm.optLocalInstalacao_3(1).Enabled = False

    objForm.TxtEMailGic.Enabled = False
    objForm.TxtRIT.Enabled = False
    objForm.CboVelocidadeAcesso.Enabled = False
    
'    objForm.CmdPendencia.Enabled = False
    objForm.CmdSalvar.Enabled = False
    objForm.CmdLimpar.Enabled = False
    objForm.CmdRealocar.Enabled = False
    
    objForm.UsrVIPEquipComplementares.DesabilitaBotoes
    
'    If objForm.CmdCadastrarEqptoNGN.Visible Then
'
'        objForm.CmdCadastrarEqptoNGN.Enabled = False
'
'    End If

End Sub

Public Property Get Acao_ID() As Long

Acao_ID = lAcao_ID

End Property

Public Property Let Acao_ID(ByVal lNewValue As Long)

lAcao_ID = lNewValue

End Property

Public Property Get Produto_ID() As Long

Produto_ID = lProduto_ID

End Property

Public Property Let Produto_ID(ByVal lNewValue As Long)

lProduto_ID = lNewValue

End Property

Public Function VerificaFinalizacaoOK(objDadosAcesso As clsGenAcesso) As Boolean
    
    Dim bRetorno        As Boolean
    Dim iDigitoVeloc    As Integer
    Dim sLetraVeloc     As String
    Dim bRetornoDadosOK As Boolean
    Dim sRetAux         As String
    Dim iAux            As Integer
    
    iAux = 0
    
    bRetornoDadosOK = False
    
    VerificaFinalizacaoOK = False
    'Verifica o preenchimento dos campos da Pasta Dados Gerais
    
    bRetorno = True
    
'    If objForm.CboTipoEquipamento_1.ListIndex = -1 And _
'       objForm.CboTipoEquipamento_2.ListIndex = -1 And _
'       objForm.CboTipoEquipamento_3.ListIndex = -1 Then
'
'        objMensagem.Descricao = "Tipo equipamento obrigatório."
'        bRetorno = False
'        GoTo Finalizar
'
'    End If
    
    iAux = iAux + 1
    ReDim vMensagemListaMsg(iAux)
    vMensagemListaMsg(0) = "Os campos abaixo foram preenchidos incorretamente ou não foram preenchidos:"
    iAux = iAux + 1
    
    If objForm.CboPadraoInterfaceEBT.ListIndex <= 0 Then
        
        bRetorno = False
        
        iAux = iAux + 1
        
        ReDim Preserve vMensagemListaMsg(iAux)
        vMensagemListaMsg(iAux - 1) = "Padrão de Interface Embratel " & _
        IIf(objForm.optCaracAcessoGen(0).Value = True, " do Acesso Principal", _
                                                       " do Acesso Contingente") & _
                                      " é obrigatório."
       
    End If
    
    If Produto_ID <> GbIDProdutoVIP Then
        If objForm.OptResponsavelInfra(0).Value = False And _
           objForm.OptResponsavelInfra(1).Value = False And objForm.optPropriedadeAcesso(0).Value = True Then
        
            bRetorno = False
            
            iAux = iAux + 1
            
            ReDim Preserve vMensagemListaMsg(iAux)
            vMensagemListaMsg(iAux - 1) = "Responsabilidade da infraEstrutura " & _
            IIf(objForm.optCaracAcessoGen(0).Value = True, " do Acesso Principal", _
                                                           " do Acesso Contingente") & _
                                          " é obrigatório."
           
        End If
    End If
'===================================================================
'= Autor: Carlos Soares                                            =
'= Data: 26/10/2004                                                =
'= Motivo: A pedido o tipo de rede não é mais obrigatório          =
'===================================================================
    'If objForm.OptTipoRede(0).Value = False And _
    '   objForm.OptTipoRede(1).Value = False Then
    '
    '    bRetorno = False
    '
    '    iAux = iAux + 1
    '
    '    ReDim Preserve vMensagemListaMsg(iAux)
    '    vMensagemListaMsg(iAux - 1) = "Tipo de Rede Determinística " & _
    '    IIf(objForm.optCaracAcessoGen(0).Value = True, " do Acesso Principal", _
    '                                                   " do Acesso Contingente") & _
    '                                  " é obrigatório."
    'End If
'===================================================================
'===================================================================
    
'    If objForm.TxtRamalGic.Text = "" Then
'
'        bRetorno = False
'
'        iAux = iAux + 1
'
'        ReDim Preserve vMensagemListaMsg(iAux)
'        vMensagemListaMsg(iAux - 1) = "Ramal " & _
'        IIf(objForm.optCaracAcessoGen(0).Value = True, " do Acesso Principal", _
'                                                       " do Acesso Contingente") & _
'                                      " é obrigatório."
'    End If
    
    If objForm.TxtEMailGic.Text = "" Then
    
        bRetorno = False
        
        iAux = iAux + 1
        
        ReDim Preserve vMensagemListaMsg(iAux)
        vMensagemListaMsg(iAux - 1) = "E-mail " & _
        IIf(objForm.optCaracAcessoGen(0).Value = True, " do Acesso Principal", _
                                                       " do Acesso Contingente") & _
                                                       " é obrigatório."
                                                       
    End If
     
    If objForm.TxtRIT.Text = "" Then
    
        bRetorno = False
       
        iAux = iAux + 1
        
        ReDim Preserve vMensagemListaMsg(iAux)
        vMensagemListaMsg(iAux - 1) = "RIT " & _
        IIf(objForm.optCaracAcessoGen(0).Value = True, " do Acesso Principal", _
                                                       " do Acesso Contingente") & _
                                                       " é obrigatório."
       
    End If
     
    If Trim(objForm.CboVelocidadeAcesso.Text) = "" Then
    
        bRetorno = False
        
        iAux = iAux + 1
        
        ReDim Preserve vMensagemListaMsg(iAux)
        vMensagemListaMsg(iAux - 1) = "Velocidade do Acesso " & _
        IIf(objForm.optCaracAcessoGen(0).Value = True, " do Acesso Principal", _
                                                       " do Acesso Contingente") & _
                                                       " é obrigatório."
    End If
    
    If objForm.optPropriedadeAcesso(0).Value = True Then  'Propriedade Embratel
    
       If Trim(objForm.SpdDataAtivacaoAcessoEBT.Text) = "" Then
          
            bRetorno = False
            
            iAux = iAux + 1
            
            ReDim Preserve vMensagemListaMsg(iAux)
            vMensagemListaMsg(iAux - 1) = "Data de Ativação " & _
            IIf(objForm.optCaracAcessoGen(0).Value = True, " do Acesso Principal", _
                                                           " do Acesso Contingente") & _
                                                           " é obrigatório."
       End If
        
       If Trim(objForm.TxtEstacaoEntregaEbtConf) = "" And _
          objForm.TxtEstacaoEntregaEbtConf.Visible = True Then
       
            bRetorno = False
            
            iAux = iAux + 1
            
            ReDim Preserve vMensagemListaMsg(iAux)
            vMensagemListaMsg(iAux - 1) = "Estação de entrega Configuração " & _
            IIf(objForm.optCaracAcessoGen(0).Value = True, " do Acesso Principal", _
                                                           " do Acesso Contingente") & _
                                                           " é obrigatório."
       
       End If
       
       If Trim(objForm.TxtEstacaoEntregaEBT) = "" Then
       
            bRetorno = False
       
            iAux = iAux + 1
            
            ReDim Preserve vMensagemListaMsg(iAux)
            vMensagemListaMsg(iAux - 1) = "Estação de entrega Embratel " & _
            IIf(objForm.optCaracAcessoGen(0).Value = True, " do Acesso Principal", _
                                                           " do Acesso Contingente") & _
                                                           " é obrigatório."
       
       End If
       
       If Trim(objForm.TxtEstacaoCliente.Text) = "" Then
       
            bRetorno = False
            
            iAux = iAux + 1
            
            ReDim Preserve vMensagemListaMsg(iAux)
            vMensagemListaMsg(iAux - 1) = "Estação do Cliente " & _
            IIf(objForm.optCaracAcessoGen(0).Value = True, " do Acesso Principal", _
                                                           " do Acesso Contingente") & _
                                                           " é obrigatório."
       
       End If
       
       If Trim(objForm.CboVelocidadeAcesso.Text) <> "" Then
       
          Call obtemDigitoVelocidade(objForm.CboVelocidadeAcesso.Text, iDigitoVeloc, sLetraVeloc)
         
          'Se a Velocidade do Circuito for >= 2M e não for cabo interno, numero ots obrigatorio
        
          If sLetraVeloc = "M" And iDigitoVeloc >= 2 And _
             objForm.CboTipoAcessoEmbratel.ListIndex <> 3 Then
        
             If Trim(objForm.TxtNumeroOtsTx_EBT.Text) = "" Then
           
                bRetorno = False
             
                iAux = iAux + 1
              
                ReDim Preserve vMensagemListaMsg(iAux)
              
                vMensagemListaMsg(iAux - 1) = "Número OTS TX " & _
                IIf(objForm.optCaracAcessoGen(0).Value = True, " do Acesso Principal", _
                                                              " do Acesso Contingente") & _
                                                              " é obrigatório."
              
             End If
           
             If Val(objForm.TxtItemOtsTxEBT.Text) = 0 Then
           
                bRetorno = False
              
                iAux = iAux + 1
              
                ReDim Preserve vMensagemListaMsg(iAux)
             
                vMensagemListaMsg(iAux - 1) = "Item OTS TX " & _
                IIf(objForm.optCaracAcessoGen(0).Value = True, " do Acesso Principal", _
                                                            " do Acesso Contingente") & _
                                                            " é obrigatório."
              
             End If
          
          End If
          
       End If
        
'       If Trim(objForm.TxtDesignacaoAcessoEBT.Text) = "" Then
'
'            bRetorno = False
'
'            iAux = iAux + 1
'
'            ReDim Preserve vMensagemListaMsg(iAux)
'
'            vMensagemListaMsg(iAux - 1) = "Designação do Acesso Embratel " & _
'            IIf(objForm.optCaracAcessoGen(0).Value = True, " do Acesso Principal", _
'                                                           " do Acesso Contingente") & _
'                                                           " é obrigatório."
'
'       End If

        'Se tipo de acesso for Radio
       If objForm.CboTipoAcessoEmbratel.ListIndex = 0 Then

          If Trim(objForm.TxtDIDFila_Radio.Text) = "" Then
             
            bRetorno = False
            
            iAux = iAux + 1
            
            ReDim Preserve vMensagemListaMsg(iAux)
            
            vMensagemListaMsg(iAux - 1) = "DID Fila " & _
            IIf(objForm.optCaracAcessoGen(0).Value = True, " do Acesso Principal", _
                                                           " do Acesso Contingente") & _
                                                           " é obrigatório."

          End If

          If Val(objForm.TxtDIDArmacao_Radio.Text) = 0 Then
             
             bRetorno = False
            
             iAux = iAux + 1
            
             ReDim Preserve vMensagemListaMsg(iAux)
            
             vMensagemListaMsg(iAux - 1) = "DID armação " & _
             IIf(objForm.optCaracAcessoGen(0).Value = True, " do Acesso Principal", _
                                                           " do Acesso Contingente") & _
                                                           " é obrigatório."

          End If

          If Val(objForm.TxtDIDRegua_Radio.Text) = 0 Then
             
            bRetorno = False

            iAux = iAux + 1
            
            ReDim Preserve vMensagemListaMsg(iAux)
            
            vMensagemListaMsg(iAux - 1) = "DID régua " & _
            IIf(objForm.optCaracAcessoGen(0).Value = True, " do Acesso Principal", _
                                                           " do Acesso Contingente") & _
                                                           " é obrigatório."

          End If

          If Val(objForm.TxtDIDPosicao_Radio.Text) = 0 Then
             
            bRetorno = False
            
            iAux = iAux + 1
            
            ReDim Preserve vMensagemListaMsg(iAux)
            
            vMensagemListaMsg(iAux - 1) = "DID posição " & _
            IIf(objForm.optCaracAcessoGen(0).Value = True, " do Acesso Principal", _
                                                           " do Acesso Contingente") & _
                                                           " é obrigatório."

          End If


'          If Val(objForm.TxtNumeroRepeticoes.Text) = 0 Then
'
'            objMensagem.Descricao = "Nº repetições obrigatório."
'            bRetorno = False
'            GoTo Finalizar
'
'          End If

       End If
    
       'Se tipo de acesso for Fibra
       If objForm.CboTipoAcessoEmbratel.ListIndex = 1 Then
    
          If Trim(objForm.TxtDIDFila_Fibra.Text) = "" Then
             
                bRetorno = False
                    
                iAux = iAux + 1
                
                ReDim Preserve vMensagemListaMsg(iAux)
                
                vMensagemListaMsg(iAux - 1) = "DID Fila " & _
                IIf(objForm.optCaracAcessoGen(0).Value = True, " do Acesso Principal", _
                                                               " do Acesso Contingente") & _
                                                               " é obrigatório."

          End If

          If Val(objForm.TxtDIDArmacao_Fibra.Text) = 0 Then
             
             bRetorno = False

             iAux = iAux + 1
            
             ReDim Preserve vMensagemListaMsg(iAux)
            
             vMensagemListaMsg(iAux - 1) = "DID armação " & _
             IIf(objForm.optCaracAcessoGen(0).Value = True, " do Acesso Principal", _
                                                           " do Acesso Contingente") & _
                                                           " é obrigatório."

          End If

          If Val(objForm.TxtDidRegua_Fibra.Text) = 0 Then
             
             bRetorno = False

             iAux = iAux + 1
             
             ReDim Preserve vMensagemListaMsg(iAux)
            
             vMensagemListaMsg(iAux - 1) = "DID régua " & _
             IIf(objForm.optCaracAcessoGen(0).Value = True, " do Acesso Principal", _
                                                           " do Acesso Contingente") & _
                                                           " é obrigatório."

          End If

          If Val(objForm.TxtDIDPosicao_Fibra.Text) = 0 Then
             
             bRetorno = False
            
             iAux = iAux + 1
            
             ReDim Preserve vMensagemListaMsg(iAux)
            
             vMensagemListaMsg(iAux - 1) = "DID posição " & _
             IIf(objForm.optCaracAcessoGen(0).Value = True, " do Acesso Principal", _
                                                           " do Acesso Contingente") & _
                                                           " é obrigatório."

          End If

          If Trim(objForm.TxtNumeroCaboFibra.Text) = "" Then
             
             bRetorno = False

             iAux = iAux + 1
            
             ReDim Preserve vMensagemListaMsg(iAux)
            
             vMensagemListaMsg(iAux - 1) = "Número Cabo " & _
             IIf(objForm.optCaracAcessoGen(0).Value = True, " do Acesso Principal", _
                                                           " do Acesso Contingente") & _
                                                           " é obrigatório."
          End If

          If Trim(objForm.TxtNumeroParFibra.Text) = "" Then
             
             bRetorno = False
                
             iAux = iAux + 1
            
             ReDim Preserve vMensagemListaMsg(iAux)
            
             vMensagemListaMsg(iAux - 1) = "Número PAR " & _
             IIf(objForm.optCaracAcessoGen(0).Value = True, " do Acesso Principal", _
                                                           " do Acesso Contingente") & _
                                                           " é obrigatório."
    
          End If
       
       End If
    
       'Se tipo de acesso for ADE
       If objForm.CboTipoAcessoEmbratel.ListIndex = 2 Then
       
          If objForm.optProprietarioModem(0).Value = False And _
             objForm.optProprietarioModem(1).Value = False Then
          
             bRetorno = False
            
             iAux = iAux + 1
            
             ReDim Preserve vMensagemListaMsg(iAux)
            
             vMensagemListaMsg(iAux - 1) = "Proprietário do Modem " & _
             IIf(objForm.optCaracAcessoGen(0).Value = True, " do Acesso Principal", _
                                                           " do Acesso Contingente") & _
                                                           " é obrigatório."
          
          End If
          
          If Trim(objForm.TxtPredioEBT_ADE.Text) = "" Then
             
             bRetorno = False
            
             iAux = iAux + 1
            
             ReDim Preserve vMensagemListaMsg(iAux)
            
             vMensagemListaMsg(iAux - 1) = "Prédio " & _
             IIf(objForm.optCaracAcessoGen(0).Value = True, " do Acesso Principal", _
                                                           " do Acesso Contingente") & _
                                                           " é obrigatório."

          End If
       
          If Trim(objForm.TxtNumeroCabo_ADE.Text) = "" Then

             bRetorno = False
            
             iAux = iAux + 1
            
             ReDim Preserve vMensagemListaMsg(iAux)
            
             vMensagemListaMsg(iAux - 1) = "Número do Cabo " & _
             IIf(objForm.optCaracAcessoGen(0).Value = True, " do Acesso Principal", _
                                                           " do Acesso Contingente") & _
                                                           " é obrigatório."

          End If
       
          If Trim(objForm.TxtNumeroPar_ADE.Text) = "" Then
             
             bRetorno = False
            
             iAux = iAux + 1
            
             ReDim Preserve vMensagemListaMsg(iAux)
            
             vMensagemListaMsg(iAux - 1) = "Número do PAR " & _
             IIf(objForm.optCaracAcessoGen(0).Value = True, " do Acesso Principal", _
                                                           " do Acesso Contingente") & _
                                                           " é obrigatório."

          End If
       
       End If
    
       'Se tipo de acesso for Cabo Interno
       If objForm.CboTipoAcessoEmbratel.ListIndex = 3 Then
    
          If objForm.optProprietarioModem(0).Value = False And _
             objForm.optProprietarioModem(1).Value = False Then
          
             bRetorno = False
            
             iAux = iAux + 1
            
             ReDim Preserve vMensagemListaMsg(iAux)
            
             vMensagemListaMsg(iAux - 1) = "Proprietário do Modem " & _
             IIf(objForm.optCaracAcessoGen(0).Value = True, " do Acesso Principal", _
                                                           " do Acesso Contingente") & _
                                                           " é obrigatório."
          
          End If
    
          If Trim(objForm.TxtPredioCabo_Cabo.Text) = "" Then
             
             bRetorno = False
            
             iAux = iAux + 1
            
             ReDim Preserve vMensagemListaMsg(iAux)
            
             vMensagemListaMsg(iAux - 1) = "Prédio " & _
             IIf(objForm.optCaracAcessoGen(0).Value = True, " do Acesso Principal", _
                                                           " do Acesso Contingente") & _
                                                           " é obrigatório."

          End If
    
          If Trim(objForm.TxtNumeroParEBT_Cabo.Text) = "" Then

             bRetorno = False
            
             iAux = iAux + 1
            
             ReDim Preserve vMensagemListaMsg(iAux)
            
             vMensagemListaMsg(iAux - 1) = "Número do PAR " & _
             IIf(objForm.optCaracAcessoGen(0).Value = True, " do Acesso Principal", _
                                                           " do Acesso Contingente") & _
                                                           " é obrigatório."

          End If
    
          If Trim(objForm.TxtNumeroCaboEBT_Cabo.Text) = "" Then
             
             bRetorno = False
            
             iAux = iAux + 1
            
             ReDim Preserve vMensagemListaMsg(iAux)
            
             vMensagemListaMsg(iAux - 1) = "Número do Cabo " & _
             IIf(objForm.optCaracAcessoGen(0).Value = True, " do Acesso Principal", _
                                                           " do Acesso Contingente") & _
                                                           " é obrigatório."

          End If
    
       End If
    Else
       
       If objForm.optTipoTecnologia(0).Value = 0 And _
          objForm.optTipoTecnologia(1).Value = 0 And _
          objForm.optTipoTecnologia(2).Value = 0 Then
          
            bRetorno = False
            
            iAux = iAux + 1
            
            ReDim Preserve vMensagemListaMsg(iAux)
            
            vMensagemListaMsg(iAux - 1) = "Tecnologia do acesso " & _
            IIf(objForm.optCaracAcessoGen(0).Value = True, " do Acesso Principal", _
                                                           " do Acesso Contingente") & _
                                                           " é obrigatório."
       End If
        
       'Propriedade Acesso Terceiro
        If objForm.optProprietarioModem(0).Value = False And _
           objForm.optProprietarioModem(1).Value = False And _
           objForm.CboTipoAcessoTerceiro.ListIndex <> 2 Then  'É obrigatório quando o
                                                              'tipo de acesso for <>
                                                              'de Rede Determinística
           bRetorno = False
          
           iAux = iAux + 1
          
           ReDim Preserve vMensagemListaMsg(iAux)
          
           vMensagemListaMsg(iAux - 1) = "Proprietário do Modem " & _
           IIf(objForm.optCaracAcessoGen(0).Value = True, " do Acesso Principal", _
                                                         " do Acesso Contingente") & _
                                                         " é obrigatório."
        
        End If
       
       If objForm.SpdDataAtivacaoAcessoTerc.Text = "" Then
          
            bRetorno = False
            
            iAux = iAux + 1
            
            ReDim Preserve vMensagemListaMsg(iAux)
            
            vMensagemListaMsg(iAux - 1) = "Data de Ativação " & _
            IIf(objForm.optCaracAcessoGen(0).Value = True, " do Acesso Principal", _
                                                           " do Acesso Contingente") & _
                                                           " é obrigatório."
       
       End If
        
'       If Trim(objForm.TxtDesignacaoAcessoTerceiro.Text) = "" Then
'
'            bRetorno = False
'
'            iAux = iAux + 1
'
'            ReDim Preserve vMensagemListaMsg(iAux)
'
'            vMensagemListaMsg(iAux - 1) = "Designação de acesso " & _
'            IIf(objForm.optCaracAcessoGen(0).Value = True, " do Acesso Principal", _
'                                                           " do Acesso Contingente") & _
'                                                           " é obrigatório."
'
'       End If
        
       If Trim(objForm.TxtEstacaoEntregaTercConf.Text) = "" Then
            
            bRetorno = False
            
            iAux = iAux + 1
            
            ReDim Preserve vMensagemListaMsg(iAux)
            
            vMensagemListaMsg(iAux - 1) = "Estação de Entrega Configuração" & _
            IIf(objForm.optCaracAcessoGen(0).Value = True, " do Acesso Principal", _
                                                           " do Acesso Contingente") & _
                                                           " é obrigatório."
       
       End If
       
       If Trim(objForm.TxtEstacaoEntregaTerc.Text) = "" Then
       
            bRetorno = False
       
            iAux = iAux + 1
            
            ReDim Preserve vMensagemListaMsg(iAux)
            
            vMensagemListaMsg(iAux - 1) = "Estação de Entrega Embratel " & _
            IIf(objForm.optCaracAcessoGen(0).Value = True, " do Acesso Principal", _
                                                           " do Acesso Contingente") & _
                                                           " é obrigatório."
       
       End If
       
       If objForm.CboProvedorUrbano.ListIndex <= 0 Then
       
            bRetorno = False
       
            iAux = iAux + 1
            
            ReDim Preserve vMensagemListaMsg(iAux)
            
            vMensagemListaMsg(iAux - 1) = "Provedor urbano " & _
            IIf(objForm.optCaracAcessoGen(0).Value = True, " do Acesso Principal", _
                                                           " do Acesso Contingente") & _
                                                           " é obrigatório."
       
       End If
                
       'Designacao de Provedor Urbano
       If Trim(objForm.SpdNumeroAcessoUrbano.Text) = "" Then
       
            bRetorno = False
       
            iAux = iAux + 1
            
            ReDim Preserve vMensagemListaMsg(iAux)
            
            vMensagemListaMsg(iAux - 1) = "Número do Acesso Urbano " & _
            IIf(objForm.optCaracAcessoGen(0).Value = True, " do Acesso Principal", _
                                                           " do Acesso Contingente") & _
                                                           " é obrigatório."
       
       End If
       
       
       If Trim(objForm.TxtCNLPontaA_Urbano.Text) = "" Then
       
            bRetorno = False
            
            iAux = iAux + 1
            
            ReDim Preserve vMensagemListaMsg(iAux)
            
            vMensagemListaMsg(iAux - 1) = "CNL Ponta A do Provedor Urbano " & _
            IIf(objForm.optCaracAcessoGen(0).Value = True, " do Acesso Principal", _
                                                           " do Acesso Contingente") & _
                                                           " é obrigatório."
       
       End If
    
       
       If Trim(objForm.SpdDataEntregaUrbano.Text) = "" Then
       
            bRetorno = False
            
            iAux = iAux + 1
            
            ReDim Preserve vMensagemListaMsg(iAux)
            
            vMensagemListaMsg(iAux - 1) = "Data de entrega do Provedor urbano " & _
            IIf(objForm.optCaracAcessoGen(0).Value = True, " do Acesso Principal", _
                                                           " do Acesso Contingente") & _
                                                           " é obrigatório."
       
       End If
       
       If objForm.optExisteAcessoIU(0).Value = True Then
       
          If objForm.CboProvedorInterurbano.ListIndex <= 0 Then
           
             bRetorno = False
            
            iAux = iAux + 1
            
            ReDim Preserve vMensagemListaMsg(iAux)
            
            vMensagemListaMsg(iAux - 1) = "Provedor interurbano " & _
            IIf(objForm.optCaracAcessoGen(0).Value = True, " do Acesso Principal", _
                                                           " do Acesso Contingente") & _
                                                           " é obrigatório."
            
          End If
          
          'Numero do acesso provedor interurbano
          If Trim(objForm.SpdNumeroAcessoInterurbano.Text) = "" Then
        
            objMensagem.Descricao = "Nº acesso interurbano obrigatório."
            bRetorno = False
            
            iAux = iAux + 1
            
            ReDim Preserve vMensagemListaMsg(iAux)
            
            vMensagemListaMsg(iAux - 1) = "Número do acesso interurbano " & _
            IIf(objForm.optCaracAcessoGen(0).Value = True, " do Acesso Principal", _
                                                           " do Acesso Contingente") & _
                                                           " é obrigatório."
        
          End If
          
          If Trim(objForm.TxtCNLPontaA.Text) = "" Then
        
            objMensagem.Descricao = "CNL ponta A interurbano obrigatório."
            bRetorno = False
        
            iAux = iAux + 1
            
            ReDim Preserve vMensagemListaMsg(iAux)
            
            vMensagemListaMsg(iAux - 1) = "CNL Ponta A interurbano " & _
            IIf(objForm.optCaracAcessoGen(0).Value = True, " do Acesso Principal", _
                                                           " do Acesso Contingente") & _
                                                           " é obrigatório."
        
          End If
          
          If Trim(objForm.TxtCNLPontaB.Text) = "" Then
        
            objMensagem.Descricao = "CNL ponta B interurbano obrigatório."
            bRetorno = False
            
            iAux = iAux + 1
            
            ReDim Preserve vMensagemListaMsg(iAux)
            
            vMensagemListaMsg(iAux - 1) = "CNL Ponta B interurbano " & _
            IIf(objForm.optCaracAcessoGen(0).Value = True, " do Acesso Principal", _
                                                           " do Acesso Contingente") & _
                                                           " é obrigatório."
        
          End If
          
          If Trim(objForm.SpdDataEntregaInterUrbano.Text) = "" Then
           
            objMensagem.Descricao = "Data entrega interurbano obrigatório."
            bRetorno = False
        
            iAux = iAux + 1
            
            ReDim Preserve vMensagemListaMsg(iAux)
            
            vMensagemListaMsg(iAux - 1) = "Data de entrega do Provedor interurbano " & _
            IIf(objForm.optCaracAcessoGen(0).Value = True, " do Acesso Principal", _
                                                           " do Acesso Contingente") & _
                                                           " é obrigatório."
        
          End If
         
          
       End If
       
'       If objForm.CboTipoAcessoTerceiro.ListIndex = 2 Then
'
'          If Val(objForm.TxtNumeroOTSTXTerc.Text) = 0 Then
'
'            objMensagem.Descricao = "Nº OTS TX obrigatório."
'
'            bRetorno = False
'
'            GoTo Finalizar
'
'          End If
'
'          If Val(objForm.TxtItemOTSTxTerc.Text) = 0 Then
'
'            objMensagem.Descricao = "Item OTS TX obrigatório."
'            bRetorno = False
'            GoTo Finalizar
'
'          End If
'
'       End If
        
       'Tipo de acesso < 2M
       If objForm.CboTipoAcessoTerceiro.ListIndex = 0 Then
'
'          If Trim(objForm.TxtNumeroCaboAcessoMenor.Text) = "" Then
'
'                objMensagem.Descricao = "Nº cabo obrigatório."
'                bRetorno = False
'                GoTo Finalizar
'
'          End If
'
'          If Trim(objForm.TxtNumeroParAcessoMenor.Text) = "" Then
'
'                objMensagem.Descricao = "Nº par obrigatório."
'                bRetorno = False
'                GoTo Finalizar
'
'          End If
       
          If objForm.CboInterfaceClienteMenor.ListIndex <= 0 Then
             
            objMensagem.Descricao = "Interface Cliente obrigatório."
            bRetorno = False
            iAux = iAux + 1
            
            ReDim Preserve vMensagemListaMsg(iAux)
            
            vMensagemListaMsg(iAux - 1) = "Interface cliente " & _
            IIf(objForm.optCaracAcessoGen(0).Value = True, " do Acesso Principal", _
                                                           " do Acesso Contingente") & _
                                                           " é obrigatório."
            
          End If

          If objForm.CboInterfaceEBTMenor.ListIndex <= 0 Then
             
             bRetorno = False
            
             iAux = iAux + 1
            
             ReDim Preserve vMensagemListaMsg(iAux)
            
             vMensagemListaMsg(iAux - 1) = "Interface Embratel " & _
             IIf(objForm.optCaracAcessoGen(0).Value = True, " do Acesso Principal", _
                                                           " do Acesso Contingente") & _
                                                           " é obrigatório."
            
          End If

'          If Trim(objForm.TxtBlocoDGMenor.Text) = "" Then
'
'                objMensagem.Descricao = "Bloco DG obrigatório."
'                bRetorno = False
'                GoTo Finalizar
'
'          End If
'
'          If Trim(objForm.TxtParDGMenor.Text) = "" Then
'
'                objMensagem.Descricao = "Par DG obrigatório."
'                bRetorno = False
'                GoTo Finalizar
'
'          End If

'          If Trim(objForm.TxtBlocoTransdataMenor.Text) = "" Then
'
'                objMensagem.Descricao = "Bloco transdata obrigatório."
'                bRetorno = False
'                GoTo Finalizar
'
'          End If
'
'          If Trim(objForm.TxtParTransdataMenor.Text) = "" Then
'
'                objMensagem.Descricao = "Par transdata obrigatório."
'                bRetorno = False
'                GoTo Finalizar
'
'          End If
       
       End If
    
       'Tipo de acesso > 2M
       If objForm.CboTipoAcessoTerceiro.ListIndex = 1 Then
    
          If Trim(objForm.TxtDIDFila.Text) = "" Then
            
            bRetorno = False
            iAux = iAux + 1
            
            ReDim Preserve vMensagemListaMsg(iAux)
            
            vMensagemListaMsg(iAux - 1) = "DID Fila" & _
            IIf(objForm.optCaracAcessoGen(0).Value = True, " do Acesso Principal", _
                                                           " do Acesso Contingente") & _
                                                           " é obrigatório."
            
          End If
    
          If Val(objForm.TxtDIDArmacao.Text) = 0 Then
            
             bRetorno = False
             
             iAux = iAux + 1
             
             ReDim Preserve vMensagemListaMsg(iAux)
             
             vMensagemListaMsg(iAux - 1) = "DID armação " & _
             IIf(objForm.optCaracAcessoGen(0).Value = True, " do Acesso Principal", _
                                                            " do Acesso Contingente") & _
                                                            " é obrigatório."
            
          End If
    
          If Val(objForm.TxtDIDRegua.Text) = 0 Then
            
             bRetorno = False
            
             iAux = iAux + 1
            
             ReDim Preserve vMensagemListaMsg(iAux)
            
             vMensagemListaMsg(iAux - 1) = "DID régua" & _
             IIf(objForm.optCaracAcessoGen(0).Value = True, " do Acesso Principal", _
                                                            " do Acesso Contingente") & _
                                                            " é obrigatório."
            
          End If
    
          If Val(objForm.TxtDIDPosicao.Text) = 0 Then
             
             bRetorno = False
            
             iAux = iAux + 1
            
             ReDim Preserve vMensagemListaMsg(iAux)
            
             vMensagemListaMsg(iAux - 1) = "DID posição" & _
             IIf(objForm.optCaracAcessoGen(0).Value = True, " do Acesso Principal", _
                                                            " do Acesso Contingente") & _
                                                            " é obrigatório."
            
          End If
    
       End If
    
       'Tipo de acesso Rede Deterministica
       If objForm.CboTipoAcessoTerceiro.ListIndex = 2 Then
    
          If Val(objForm.TxtDominio.Text) = 0 Then
            
             objMensagem.Descricao = "Domínio obrigatório."
             bRetorno = False
         
             iAux = iAux + 1
            
             ReDim Preserve vMensagemListaMsg(iAux)
            
             vMensagemListaMsg(iAux - 1) = "Domínio" & _
             IIf(objForm.optCaracAcessoGen(0).Value = True, " do Acesso Principal", _
                                                           " do Acesso Contingente") & _
                                                           " é obrigatório."
            
          End If
       
          If Val(objForm.txtNo.Text) = 0 Then
            
             objMensagem.Descricao = "Nó obrigatório."
             bRetorno = False
             
             iAux = iAux + 1
            
             ReDim Preserve vMensagemListaMsg(iAux)
            
             vMensagemListaMsg(iAux - 1) = "Número do Nó" & _
             IIf(objForm.optCaracAcessoGen(0).Value = True, " do Acesso Principal", _
                                                           " do Acesso Contingente") & _
                                                           " é obrigatório."
            
          End If
       
          If Trim(objForm.CboSlot_Rede.Text) = "" Then
            
             bRetorno = False
            
             iAux = iAux + 1
            
             ReDim Preserve vMensagemListaMsg(iAux)
            
             vMensagemListaMsg(iAux - 1) = "SLOT " & _
             IIf(objForm.optCaracAcessoGen(0).Value = True, " do Acesso Principal", _
                                                           " do Acesso Contingente") & _
                                                           " é obrigatório."
            
          End If
       
          If Trim(objForm.CboPorta_Rede.Text) = "" Then
            
             bRetorno = False
            
             iAux = iAux + 1
            
             ReDim Preserve vMensagemListaMsg(iAux)
            
             vMensagemListaMsg(iAux - 1) = "Porta" & _
             IIf(objForm.optCaracAcessoGen(0).Value = True, " do Acesso Principal", _
                                                           " do Acesso Contingente") & _
                                                           " é obrigatório."
          End If
          
          If Val(objForm.CboTSI_Rede.Text) = 0 Then
            
             bRetorno = False
                
             iAux = iAux + 1
            
             ReDim Preserve vMensagemListaMsg(iAux)
            
             vMensagemListaMsg(iAux - 1) = "Time slot inicial" & _
             IIf(objForm.optCaracAcessoGen(0).Value = True, " do Acesso Principal", _
                                                           " do Acesso Contingente") & _
                                                           " é obrigatório."
            
          End If
       
          If Val(objForm.CboTSF_Rede.Text) = 0 Then
             
             bRetorno = False
            
             iAux = iAux + 1
            
             ReDim Preserve vMensagemListaMsg(iAux)
            
             vMensagemListaMsg(iAux - 1) = "Time slot final" & _
             IIf(objForm.optCaracAcessoGen(0).Value = True, " do Acesso Principal", _
                                                           " do Acesso Contingente") & _
                                                           " é obrigatório."
          End If
       
          If Trim(objForm.TxtDIDFila_Rede.Text) = "" Then
            
             bRetorno = False
            
             iAux = iAux + 1
            
             ReDim Preserve vMensagemListaMsg(iAux)
            
             vMensagemListaMsg(iAux - 1) = "DID Fila" & _
             IIf(objForm.optCaracAcessoGen(0).Value = True, " do Acesso Principal", _
                                                           " do Acesso Contingente") & _
                                                           " é obrigatório."
            
          End If
    
          If Val(objForm.TxtDIDArmacao_Rede.Text) = 0 Then
            
             bRetorno = False
             iAux = iAux + 1
            
             ReDim Preserve vMensagemListaMsg(iAux)
            
             vMensagemListaMsg(iAux - 1) = "DID armação" & _
             IIf(objForm.optCaracAcessoGen(0).Value = True, " do Acesso Principal", _
                                                           " do Acesso Contingente") & _
                                                           " é obrigatório."
            
          End If
    
          If Val(objForm.TxtDIDRegua_Rede.Text) = 0 Then
            
             bRetorno = False
            
             iAux = iAux + 1
            
             ReDim Preserve vMensagemListaMsg(iAux)
            
             vMensagemListaMsg(iAux - 1) = "DID Régua" & _
             IIf(objForm.optCaracAcessoGen(0).Value = True, " do Acesso Principal", _
                                                           " do Acesso Contingente") & _
                                                           " é obrigatório."
            
          End If
    
          If Val(objForm.TxtDIDPosicao_Rede.Text) = 0 Then
             
             bRetorno = False
             iAux = iAux + 1
            
             ReDim Preserve vMensagemListaMsg(iAux)
            
             vMensagemListaMsg(iAux - 1) = "DID Posição" & _
             IIf(objForm.optCaracAcessoGen(0).Value = True, " do Acesso Principal", _
                                                           " do Acesso Contingente") & _
                                                           " é obrigatório."
            
          End If
       
       End If
    
    End If
    

    If objForm.Visible Then
        If objForm.CmdObterEnderecoIP.Caption <> "Dados de Switch - OK" Then
        
            bRetorno = False
            
            iAux = iAux + 1
            
            ReDim Preserve vMensagemListaMsg(iAux)
            
            vMensagemListaMsg(iAux - 1) = "Os dados de Switch são obrigatórios para a finalização da fase!"
            
        End If
    
    End If
    
    If objForm.optSegurancaAcesso(0).Value = True Then   'Seguranca de Acesso igual a Sim
           
           If objForm.optCaracAcessoEmbratel(1).Value = True Then  'Acesso Contingente
              
              'Caso os dados do Acesso Contingente estejam OK, verifica se os dados do
              'Acesso Principal foi gravado 100%
              
'              If bRetorno Then
              
                 sRetAux = VerificaDadosOK("P")
                 
                 If sRetAux = "O" Then
                 
                    bRetornoDadosOK = True
                 
                 ElseIf sRetAux = "N" Then
                    
                    bRetornoDadosOK = False
                 
                 End If
                 
'              End If
              
           Else
           
              'Caso os dados do Acesso Principal estejam OK, verifica se os dados do
              'Acesso Contingente foi gravado 100%
              
'              If bRetorno Then
              
                  sRetAux = VerificaDadosOK("C")

                 If sRetAux = "O" Then
                 
                    bRetornoDadosOK = True
                 
                 ElseIf sRetAux = "N" Then
                    
                    bRetornoDadosOK = False
                 
                 ElseIf sRetAux = "I" Then
                    
                    bRetornoDadosOK = False
                 
                 End If

           End If
    Else
        
        bRetornoDadosOK = True
    
    End If
    
                    
    If bRetorno Then
    
       bRetorno = bRetornoDadosOK
       
    End If
    
Finalizar:
    
    If objMensagem.TipoMsg <> "" Then
    
        objMensagem.ExibeMensagem
            
    Else
    
    If Not bRetorno Then
        
        If sRetAux <> "I" Then
           
           If objForm.optSegurancaAcesso(0).Value = True Then   'Seguranca de Acesso igual a Sim
              
              If objForm.optCaracAcessoEmbratel(1).Value = True Then  'Acesso Contingente
                 
                 Call ExibeListaMensagem
                 
              End If
              
           Else
           
              Call ExibeListaMensagem
           
           End If
        
        End If
        
    End If
    
    End If
    
    VerificaFinalizacaoOK = bRetorno
    
    Exit Function
       
End Function

Public Function VerificaDadosOK(pCaracteristicaAcesso As String) As String

    Dim bRetorno                As Boolean
    Dim iDigitoVeloc            As Integer
    Dim sLetraVeloc             As String
    
    Dim objAcesso               As clsGenAcesso
    Dim objDadosPropriedadeEBT  As clsGenPropriedadeEmbratel
    Dim objDadosPropriedadeTerc As clsGenPropriedadeTerceiros
    Dim lContingenciaAcesso_ID  As Long
    Dim sCaracteristicaAcesso   As String
    
    Dim objDadosTipoRadio       As clsGenTipoAcessoRadio
    Dim objDadosTipoFibra       As clsGenTipoAcessoFibra
    Dim objdadosTipoADE         As clsGenTipoAcessoADE
    Dim objDadosTipoCabo        As clsGenTipoAcessoCabo
    Dim objDadosVelocidade      As clsGenVelocidade
              
    Dim objdadosTipoAcessoMaior As clsGenTipoAcessoMaior2M
    Dim objdadosTipoAcessoMenor As clsGenTipoAcessoMenor2M
    Dim objdadosTipoAcessoRede  As clsGenTipoAcessoRedeDeterm
    '#BLD
    Dim sTabela                 As String
    
    Dim bRetornoOK              As Boolean
    Dim iAux                    As Integer
    
    On Error GoTo ErrorHandler

    bRetorno = True
    
    iAux = UBound(vMensagemListaMsg, 1)
    iAux = iAux + 1
    
    VerificaDadosOK = "N"       'N - Não está OK. Existem dados não preenchidos ou
                                'preenchidos incorretamente

    Set objAcesso = New clsGenAcesso
    
    objAcesso.InformaMensagem objMensagem
    objAcesso.InformaBanco frmgenMDI_SGAPlus.objBanco
    
    sCaracteristicaAcesso = pCaracteristicaAcesso
    
    If Not objAcesso.RecuperarDados(sCaracteristicaAcesso, Servico_ID, Fase_ID, objMensagem) Then
    
       bRetorno = False
       
       If objMensagem.TipoMsg <> "" Then
           
          objMensagem.ExibeMensagem
          
          GoTo Finalizar
        
       End If
        
'       objMensagem.Descricao = "Os dados do Acesso Contingente são de preenchimento " & _
                               "obrigatório para a finalização da fase! "
       
'       iAux = iAux + 1
    
'       ReDim Preserve vMensagemListaMsg(iAux)
    
'       vMensagemListaMsg(iAux - 1) = "Os dados do Acesso Contingente são de " & _
                                     "preenchimento obrigatório para "
       
'       iAux = iAux + 1
    
'       ReDim Preserve vMensagemListaMsg(iAux)
    
'       vMensagemListaMsg(iAux - 1) = "a finalização da fase!"
       
       VerificaDadosOK = "I"  'Dados inexistentes
       
       GoTo Finalizar
        
    Else
    
        If objAcesso.PadraoInterfaceEBT_ID = 0 Then
        
           bRetorno = False
           
           iAux = iAux + 1
        
           ReDim Preserve vMensagemListaMsg(iAux)
        
           vMensagemListaMsg(iAux - 1) = "Padrão de Interface Embratel" & _
           IIf(sCaracteristicaAcesso = "P", " do Acesso Principal", _
                                                       " do Acesso Contingente") & _
                                                       " é obrigatório."
           
        End If
    
        If objAcesso.ResponsavelInfra = "" And objAcesso.PropridedadeAcesso = "E" Then
        
           bRetorno = False
           
           iAux = iAux + 1
        
           ReDim Preserve vMensagemListaMsg(iAux)
        
           vMensagemListaMsg(iAux - 1) = "Responsabilidade Infra-estrutura" & _
           IIf(sCaracteristicaAcesso = "P", " do Acesso Principal", _
                                                       " do Acesso Contingente") & _
                                                       " é obrigatório."
           
        End If
        
        lContingenciaAcesso_ID = objAcesso.iD
        
        'Obtem os dados de Equipamento Complementar
        bRetornoOK = objAcesso.RecuperarDadosEquipamentos(lContingenciaAcesso_ID, objMensagem)
          
        If objMensagem.TipoMsg <> "" Then
            
           GoTo ErrorHandler
         
        End If
          
        'Caso o Acesso já esteja Cadastrado, Recuperar os dados Conforme a Propriedade do
        'Acesso
        If objAcesso.PropridedadeAcesso = "E" Then   'Propriedade Acesso = Embratel
        
           Set objDadosPropriedadeEBT = New clsGenPropriedadeEmbratel
           
           objDadosPropriedadeEBT.InformaBanco frmgenMDI_SGAPlus.objBanco
           
           bRetornoOK = objDadosPropriedadeEBT.RecuperarDados(lContingenciaAcesso_ID, objMensagem)
            
           If objMensagem.TipoMsg <> "" Then
              
              GoTo ErrorHandler
           
           End If
           
           If objDadosPropriedadeEBT.EstacaoEntregaAcesso = "" Then
           
              bRetorno = False
              
              iAux = iAux + 1
        
              ReDim Preserve vMensagemListaMsg(iAux)
        
              vMensagemListaMsg(iAux - 1) = "Estação de entrega Embratel" & _
              IIf(sCaracteristicaAcesso = "P", " do Acesso Principal", _
                                                          " do Acesso Contingente") & _
                                                          " é obrigatório."
              
           End If
           
           If objDadosPropriedadeEBT.EstacaoEntregaAcessoConfiguracao = "" Then
              
              bRetorno = False
              
              iAux = iAux + 1
        
              ReDim Preserve vMensagemListaMsg(iAux)
        
              vMensagemListaMsg(iAux - 1) = "Estação de Entrega Configuração" & _
              IIf(sCaracteristicaAcesso = "P", " do Acesso Principal", _
                                                          " do Acesso Contingente") & _
                                                          " é obrigatório."
              
           End If
           
           If objDadosPropriedadeEBT.estacaoCliente = "" Then
              
              bRetorno = False
              
              iAux = iAux + 1
        
              ReDim Preserve vMensagemListaMsg(iAux)
        
              vMensagemListaMsg(iAux - 1) = "Estação do Cliente" & _
              IIf(sCaracteristicaAcesso = "P", " do Acesso Principal", _
                                                          " do Acesso Contingente") & _
                                                          " é obrigatório."
              
           End If
           
           If objDadosPropriedadeEBT.DataAtivacaoAcesso = "" Then
           
              bRetorno = False
              iAux = iAux + 1
        
              ReDim Preserve vMensagemListaMsg(iAux)
        
              vMensagemListaMsg(iAux - 1) = "Data de Ativação " & _
              IIf(sCaracteristicaAcesso = "P", " do Acesso Principal", _
                                                          " do Acesso Contingente") & _
                                                          " é obrigatório."
              
           End If
           
           Set objDadosVelocidade = New clsGenVelocidade
           
           objDadosVelocidade.InformaMensagem objMensagem
           objDadosVelocidade.InformaBanco frmgenMDI_SGAPlus.objBanco
                       
           '#BLD
           sTabela = ""
           
           If (Produto_ID = GbIDProdutoBLD Or _
                    Produto_ID = GbIDProdutoVOIP) Then

               sTabela = "TAB_VELOCIDADE"
           
           End If
           
           bRetorno = objDadosVelocidade.RecuperarDados(objDadosAcesso.Cd_Veloc, _
                                        sTabela)
           
           If bRetorno Then
           
              Call obtemDigitoVelocidade(objDadosVelocidade.Nm_Veloc, iDigitoVeloc, sLetraVeloc)
           
           Else
              
              objMensagem.TipoMsg = "A"
              objMensagem.Descricao = "Foi encontrado um erro de inconsistência " & _
                                      "de dados referente a Velocidade do Acesso"
              GoTo Finalizar
           
           End If
           
           'Se a Velocidade do Circuito for >= 2M e não for cabo interno, numero ots obrigatorio
           If sLetraVeloc = "M" And iDigitoVeloc >= 2 And _
              objDadosPropriedadeEBT.TipoAcesso <> "C" Then
                
                If Trim(objDadosPropriedadeEBT.NumeroOTSTX) = "" Then
                   
                   bRetorno = False
                     
                   iAux = iAux + 1
            
                   ReDim Preserve vMensagemListaMsg(iAux)
            
                   vMensagemListaMsg(iAux - 1) = "Número OTS TX" & _
                   IIf(sCaracteristicaAcesso = "P", " do Acesso Principal", _
                                                              " do Acesso Contingente") & _
                                                              " é obrigatório."
                     
                End If
           
           
                If Val(objDadosPropriedadeEBT.ItemOTSTX) = 0 Then
                  
                   bRetorno = False
                     
                   iAux = iAux + 1
            
                   ReDim Preserve vMensagemListaMsg(iAux)
            
                   vMensagemListaMsg(iAux - 1) = "Item OTS TX" & _
                   IIf(sCaracteristicaAcesso = "P", " do Acesso Principal", _
                                                              " do Acesso Contingente") & _
                                                              " é obrigatório."
                     
                End If
           
           End If
           
           'Carrega a propriedade Objeto Propriedade Embratel
           If objDadosPropriedadeEBT.TipoAcesso = "R" Then
               
              Set objDadosTipoRadio = New clsGenTipoAcessoRadio
              
              objDadosTipoRadio.InformaBanco frmgenMDI_SGAPlus.objBanco
           
              bRetornoOK = objDadosTipoRadio.RecuperarDados(lContingenciaAcesso_ID, objMensagem)
              
              If objMensagem.TipoMsg <> "" Then
                 
                 GoTo ErrorHandler
                 
              End If

              If objDadosTipoRadio.DIDFila = "" Then
                 
                  bRetorno = False
                  
                  iAux = iAux + 1
            
                  ReDim Preserve vMensagemListaMsg(iAux)
            
                  vMensagemListaMsg(iAux - 1) = "DID Fila" & _
                  IIf(sCaracteristicaAcesso = "P", " do Acesso Principal", _
                                                              " do Acesso Contingente") & _
                                                              " é obrigatório."
    
               End If
    
               If objDadosTipoRadio.DIDArmacao = 0 Then
                 
                  bRetorno = False
    
                  iAux = iAux + 1
            
                  ReDim Preserve vMensagemListaMsg(iAux)
            
                  vMensagemListaMsg(iAux - 1) = "DID armação" & _
                  IIf(sCaracteristicaAcesso = "P", " do Acesso Principal", _
                                                              " do Acesso Contingente") & _
                                                              " é obrigatório."
    
               End If
      
               If objDadosTipoRadio.DIDRegua = 0 Then
                 
                  bRetorno = False
                  
                  iAux = iAux + 1
            
                  ReDim Preserve vMensagemListaMsg(iAux)
            
                  vMensagemListaMsg(iAux - 1) = "DID Régua" & _
                  IIf(sCaracteristicaAcesso = "P", " do Acesso Principal", _
                                                              " do Acesso Contingente") & _
                                                              " é obrigatório."
    
               End If
    
               If objDadosTipoRadio.DIDPosicao = 0 Then
                 
                  bRetorno = False
                  
                  iAux = iAux + 1
            
                  ReDim Preserve vMensagemListaMsg(iAux)
            
                  vMensagemListaMsg(iAux - 1) = "DID posição" & _
                  IIf(sCaracteristicaAcesso = "P", " do Acesso Principal", _
                                                              " do Acesso Contingente") & _
                                                              " é obrigatório."
    
               End If
     
'               If objDadosTipoRadio.NumeroRepeticoes = 0 Then
'
'                  objMensagem.Descricao = "Nº repetições obrigatório."
'                  bRetorno = False
'                  GoTo Finalizar
'
'               End If
    
               Set objDadosTipoRadio = Nothing
               
           ElseIf objDadosPropriedadeEBT.TipoAcesso = "F" Then  'Fibra Otica
              
              Set objDadosTipoFibra = New clsGenTipoAcessoFibra
              
              objDadosTipoFibra.InformaBanco frmgenMDI_SGAPlus.objBanco
           
              bRetornoOK = objDadosTipoFibra.RecuperarDados(lContingenciaAcesso_ID, objMensagem)
              
              If objMensagem.TipoMsg <> "" Then
                 
                 GoTo ErrorHandler
                 
              End If
              
              If objDadosTipoFibra.DIDFila = "" Then
              
                 bRetorno = False
                 
                 iAux = iAux + 1
            
                 ReDim Preserve vMensagemListaMsg(iAux)
           
                 vMensagemListaMsg(iAux - 1) = "DID Fila" & _
                 IIf(sCaracteristicaAcesso = "P", " do Acesso Principal", _
                                                             " do Acesso Contingente") & _
                                                             " é obrigatório."
                 
              End If
              
              If objDadosTipoFibra.DIDArmacao = 0 Then
              
                 bRetorno = False
                 
                 iAux = iAux + 1
            
                 ReDim Preserve vMensagemListaMsg(iAux)
           
                 vMensagemListaMsg(iAux - 1) = "DID armação" & _
                 IIf(sCaracteristicaAcesso = "P", " do Acesso Principal", _
                                                             " do Acesso Contingente") & _
                                                             " é obrigatório."
                 
              End If
              
              If objDadosTipoFibra.DIDPosicao = 0 Then
              
                 bRetorno = False
                 
                 iAux = iAux + 1
            
                 ReDim Preserve vMensagemListaMsg(iAux)
           
                 vMensagemListaMsg(iAux - 1) = "DID posição" & _
                 IIf(sCaracteristicaAcesso = "P", " do Acesso Principal", _
                                                             " do Acesso Contingente") & _
                                                             " é obrigatório."
                 
              End If
              
              If objDadosTipoFibra.DIDRegua = 0 Then
              
                 bRetorno = False
                 
                 iAux = iAux + 1
            
                 ReDim Preserve vMensagemListaMsg(iAux)
           
                 vMensagemListaMsg(iAux - 1) = "DID Régua" & _
                 IIf(sCaracteristicaAcesso = "P", " do Acesso Principal", _
                                                             " do Acesso Contingente") & _
                                                             " é obrigatório."
                 
              End If
              
              If objDadosTipoFibra.NumeroCabo = "" Then
              
                 bRetorno = False
                 
                 iAux = iAux + 1
            
                 ReDim Preserve vMensagemListaMsg(iAux)
           
                 vMensagemListaMsg(iAux - 1) = "Número do Cabo" & _
                 IIf(sCaracteristicaAcesso = "P", " do Acesso Principal", _
                                                             " do Acesso Contingente") & _
                                                             " é obrigatório."
                 
              End If
              
              If objDadosTipoFibra.NumeroPar = "" Then
              
                 bRetorno = False
                 
                 iAux = iAux + 1
            
                 ReDim Preserve vMensagemListaMsg(iAux)
           
                 vMensagemListaMsg(iAux - 1) = "Número PAR" & _
                 IIf(sCaracteristicaAcesso = "P", " do Acesso Principal", _
                                                             " do Acesso Contingente") & _
                                                             " é obrigatório."
                 
              End If
              
              Set objDadosTipoFibra = Nothing
           
           ElseIf objDadosPropriedadeEBT.TipoAcesso = "A" Then  'ADE
              
              If objAcesso.ProprietarioModem = "" Then
                
                 bRetorno = False
                  
                 iAux = iAux + 1
            
                 ReDim Preserve vMensagemListaMsg(iAux)
           
                 vMensagemListaMsg(iAux - 1) = "Proprietário Modem" & _
                 IIf(sCaracteristicaAcesso = "P", " do Acesso Principal", _
                                                             " do Acesso Contingente") & _
                                                             " é obrigatório."
                  
              End If
                
              Set objdadosTipoADE = New clsGenTipoAcessoADE
              
              objdadosTipoADE.InformaBanco frmgenMDI_SGAPlus.objBanco
           
              bRetornoOK = objdadosTipoADE.RecuperarDados(lContingenciaAcesso_ID, objMensagem)
              
              If objMensagem.TipoMsg <> "" Then
                 
                 GoTo ErrorHandler
                 
              End If
              
              If objdadosTipoADE.PredioEBT = "" Then
              
                 bRetorno = False
                 
                 iAux = iAux + 1
            
                 ReDim Preserve vMensagemListaMsg(iAux)
           
                 vMensagemListaMsg(iAux - 1) = "Prédio Embratel" & _
                 IIf(sCaracteristicaAcesso = "P", " do Acesso Principal", _
                                                             " do Acesso Contingente") & _
                                                             " é obrigatório."
                 
              End If
              
              If objdadosTipoADE.NumeroParEBT = "" Then
              
                 bRetorno = False
                 
                 iAux = iAux + 1
            
                 ReDim Preserve vMensagemListaMsg(iAux)
           
                 vMensagemListaMsg(iAux - 1) = "Número PAR Embratel" & _
                 IIf(sCaracteristicaAcesso = "P", " do Acesso Principal", _
                                                             " do Acesso Contingente") & _
                                                             " é obrigatório."
                 
              End If
              
              If objdadosTipoADE.NumeroCaboEBT = 0 Then
              
                 bRetorno = False
                 
                 iAux = iAux + 1
            
                 ReDim Preserve vMensagemListaMsg(iAux)
           
                 vMensagemListaMsg(iAux - 1) = "Número Cabo Embratel" & _
                 IIf(sCaracteristicaAcesso = "P", " do Acesso Principal", _
                                                             " do Acesso Contingente") & _
                                                             " é obrigatório."
                 
              End If
              
              Set objdadosTipoADE = Nothing
           
           ElseIf objDadosPropriedadeEBT.TipoAcesso = "C" Then  'Cabo Interno
              
              If objAcesso.ProprietarioModem = "" Then
                
                 bRetorno = False
                 
                 iAux = iAux + 1
            
                 ReDim Preserve vMensagemListaMsg(iAux)
           
                 vMensagemListaMsg(iAux - 1) = "Proprietário Modem" & _
                 IIf(sCaracteristicaAcesso = "P", " do Acesso Principal", _
                                                             " do Acesso Contingente") & _
                                                             " é obrigatório."
                  
              End If
              
              Set objDadosTipoCabo = New clsGenTipoAcessoCabo
              
              objDadosTipoCabo.InformaBanco frmgenMDI_SGAPlus.objBanco
           
              bRetornoOK = objDadosTipoCabo.RecuperarDados(lContingenciaAcesso_ID, objMensagem)
              
              If objMensagem.TipoMsg <> "" Then
                 
                 GoTo ErrorHandler
                 
              End If
              
              objDadosPropriedadeEBT.objTipoAcesso = objDadosTipoCabo
           
'              If Trim(objDadosTipoCabo.NumeroCaboCliente) = "" Then
'
'                 objMensagem.Descricao = "Nº cabo cliente obrigatório."
'                 bRetorno = False
'                 GoTo Finalizar
'
'              End If
'
'              If Trim(objDadosTipoCabo.NumeroParCliente) = "" Then
'
'                 objMensagem.Descricao = "Nº par cliente obrigatório."
'                 bRetorno = False
'                 GoTo Finalizar
'
'              End If
            
              If Trim(objDadosTipoCabo.NumeroCaboEBT) = "" Then
                 bRetorno = False
               
                 iAux = iAux + 1
            
                 ReDim Preserve vMensagemListaMsg(iAux)
           
                 vMensagemListaMsg(iAux - 1) = "Número Cabo Embratel" & _
                 IIf(sCaracteristicaAcesso = "P", " do Acesso Principal", _
                                                             " do Acesso Contingente") & _
                                                             " é obrigatório."
               
              End If
                       
              If Trim(objDadosTipoCabo.NumeroParEBT) = "" Then
            
                 bRetorno = False
               
                 iAux = iAux + 1
            
                 ReDim Preserve vMensagemListaMsg(iAux)
           
                 vMensagemListaMsg(iAux - 1) = "Número PAR Embratel" & _
                 IIf(sCaracteristicaAcesso = "P", " do Acesso Principal", _
                                                             " do Acesso Contingente") & _
                                                             " é obrigatório."
               
              End If
            
              If Trim(objDadosTipoCabo.Predio) = "" Then
            
                 bRetorno = False
                 
                 iAux = iAux + 1
            
                 ReDim Preserve vMensagemListaMsg(iAux)
           
                 vMensagemListaMsg(iAux - 1) = "Prédio" & _
                 IIf(sCaracteristicaAcesso = "P", " do Acesso Principal", _
                                                             " do Acesso Contingente") & _
                                                             " é obrigatório."
               
              End If
            
              Set objDadosTipoCabo = Nothing
           End If
        
           Set objDadosPropriedadeEBT = Nothing
        
        Else
           
           'Propriedade Acesso = Terceiros
           Set objDadosPropriedadeTerc = New clsGenPropriedadeTerceiros
           
           objDadosPropriedadeTerc.InformaBanco frmgenMDI_SGAPlus.objBanco
           
           bRetornoOK = objDadosPropriedadeTerc.RecuperarDados(lContingenciaAcesso_ID, objMensagem)
            
           If objMensagem.TipoMsg <> "" Then
              
              GoTo ErrorHandler
           
           End If
           
           If objDadosPropriedadeTerc.EstacaoEntregaAcessoConfiguracao = "" Then
             
              bRetorno = False
              
              iAux = iAux + 1
         
              ReDim Preserve vMensagemListaMsg(iAux)
        
              vMensagemListaMsg(iAux - 1) = "Estação de entrega Configuração " & _
              IIf(sCaracteristicaAcesso = "P", " do Acesso Principal", _
                                                          " do Acesso Contingente") & _
                                                          " é obrigatório."
           End If
           
           If objDadosPropriedadeTerc.EstacaoEntregaAcesso = "" Then
             
              bRetorno = False
              
              iAux = iAux + 1
         
              ReDim Preserve vMensagemListaMsg(iAux)
        
              vMensagemListaMsg(iAux - 1) = "Estação de entrega Embratel " & _
              IIf(sCaracteristicaAcesso = "P", " do Acesso Principal", _
                                                          " do Acesso Contingente") & _
                                                          " é obrigatório."
           End If
           
           If objDadosPropriedadeTerc.DataAtivacaoAcesso = "" Then
             
              bRetorno = False
              
              iAux = iAux + 1
         
              ReDim Preserve vMensagemListaMsg(iAux)
        
              vMensagemListaMsg(iAux - 1) = "Data de Ativação do acesso " & _
              IIf(sCaracteristicaAcesso = "P", " do Acesso Principal", _
                                                          " do Acesso Contingente") & _
                                                          " é obrigatório."
           End If
           
'           If objDadosPropriedadeTerc.TipoAcesso = "R" Then
'
'               If Val(objDadosPropriedadeTerc.NumeroOTSTx) = 0 Then
'
'                  objMensagem.Descricao = "Nº OTS TX obrigatório."
'
'                  bRetorno = False
'
'                  GoTo Finalizar
'
'               End If
'
'
'               If Val(objDadosPropriedadeTerc.ItemOTSTx) = 0 Then
'
'                  objMensagem.Descricao = "Item OTS TX obrigatório."
'
'                  bRetorno = False
'
'                  GoTo Finalizar
'
'               End If
'
'           End If
'
          'Carrega a propriedade Objeto Propriedade Terceiros
           If objDadosPropriedadeTerc.TipoAcesso = "<" Then   'Tipo de Acesso Menor 2M
              
              If objAcesso.ProprietarioModem = "" Then
                
                 bRetorno = False
                 
                 iAux = iAux + 1
            
                 ReDim Preserve vMensagemListaMsg(iAux)
           
                 vMensagemListaMsg(iAux - 1) = "Proprietário Modem" & _
                 IIf(sCaracteristicaAcesso = "P", " do Acesso Principal", _
                                                             " do Acesso Contingente") & _
                                                             " é obrigatório."
              End If
               
              Set objdadosTipoAcessoMenor = New clsGenTipoAcessoMenor2M
              
              objdadosTipoAcessoMenor.InformaBanco frmgenMDI_SGAPlus.objBanco
           
              bRetornoOK = objdadosTipoAcessoMenor.RecuperarDados(lContingenciaAcesso_ID, objMensagem)
              
              If objMensagem.TipoMsg <> "" Then
                 
                 GoTo ErrorHandler
                 
              End If
                
'              If objdadosTipoAcessoMenor.NumeroCabo = "" Then
'
'                 objMensagem.Descricao = "Nº cabo obrigatório."
'                 bRetorno = False
'                 GoTo Finalizar
'
'              End If
'
'              If objdadosTipoAcessoMenor.NumeroPar = "" Then
'
'                 objMensagem.Descricao = "Nº par obrigatório."
'                 bRetorno = False
'                 GoTo Finalizar
'
'              End If
              
              If objdadosTipoAcessoMenor.InterfaceClienteID = 0 Then
              
                 bRetorno = False
                 
                 iAux = iAux + 1
            
                 ReDim Preserve vMensagemListaMsg(iAux)
           
                 vMensagemListaMsg(iAux - 1) = "Interface cliente" & _
                 IIf(sCaracteristicaAcesso = "P", " do Acesso Principal", _
                                                             " do Acesso Contingente") & _
                                                             " é obrigatório."
                 
              End If
              
              If objdadosTipoAcessoMenor.InterfaceEBTID = 0 Then
              
                 bRetorno = False
                 iAux = iAux + 1
            
                 ReDim Preserve vMensagemListaMsg(iAux)
           
                 vMensagemListaMsg(iAux - 1) = "Interface Embratel" & _
                 IIf(sCaracteristicaAcesso = "P", " do Acesso Principal", _
                                                             " do Acesso Contingente") & _
                                                             " é obrigatório."
                 
              End If
              
'              If objdadosTipoAcessoMenor.BlocoDG = "" Then
'
'                 objMensagem.Descricao = "Bloco DG obrigatório."
'                 bRetorno = False
'                 GoTo Finalizar
'
'              End If
'
'              If objdadosTipoAcessoMenor.ParDG = "" Then
'
'                 objMensagem.Descricao = "Par DG obrigatório."
'                 bRetorno = False
'                 GoTo Finalizar
'
'              End If
              
'              If objdadosTipoAcessoMenor.BlocoTransData = "" Then
'
'                 objMensagem.Descricao = "Bloco transdata obrigatório."
'                 bRetorno = False
'                 GoTo Finalizar
'
'              End If
'
'              If objdadosTipoAcessoMenor.ParTransData = "" Then
'
'                 objMensagem.Descricao = "Par transdata obrigatório."
'                 bRetorno = False
'                 GoTo Finalizar
'
'              End If
              
              Set objdadosTipoAcessoMenor = Nothing

           ElseIf objDadosPropriedadeTerc.TipoAcesso = ">" Then  'Tipo de Acesso Maior 2M
              
              If objAcesso.ProprietarioModem = "" Then
                
                 bRetorno = False
                  
                 iAux = iAux + 1
            
                 ReDim Preserve vMensagemListaMsg(iAux)
           
                 vMensagemListaMsg(iAux - 1) = "Proprietário Modem" & _
                 IIf(sCaracteristicaAcesso = "P", " do Acesso Principal", _
                                                             " do Acesso Contingente") & _
                                                             " é obrigatório."
                  
              End If
              
              Set objdadosTipoAcessoMaior = New clsGenTipoAcessoMaior2M
              
              objdadosTipoAcessoMaior.InformaBanco frmgenMDI_SGAPlus.objBanco
           
              bRetornoOK = objdadosTipoAcessoMaior.RecuperarDados(lContingenciaAcesso_ID, objMensagem)
              
              If objMensagem.TipoMsg <> "" Then
                 
                 GoTo ErrorHandler
                 
              End If
              
              If objdadosTipoAcessoMaior.DIDFila = "" Then
              
                 bRetorno = False
                 iAux = iAux + 1
            
                 ReDim Preserve vMensagemListaMsg(iAux)
           
                 vMensagemListaMsg(iAux - 1) = "DID Fila" & _
                 IIf(sCaracteristicaAcesso = "P", " do Acesso Principal", _
                                                             " do Acesso Contingente") & _
                                                             " é obrigatório."
                 
              End If
              
              If objdadosTipoAcessoMaior.DIDArmacao = 0 Then
              
                 bRetorno = False
                 
                 iAux = iAux + 1
            
                 ReDim Preserve vMensagemListaMsg(iAux)
           
                 vMensagemListaMsg(iAux - 1) = "DID armação" & _
                 IIf(sCaracteristicaAcesso = "P", " do Acesso Principal", _
                                                             " do Acesso Contingente") & _
                                                             " é obrigatório."
                 
              End If
              
              If objdadosTipoAcessoMaior.DIDRegua = 0 Then
              
                 bRetorno = False
                 iAux = iAux + 1
            
                 ReDim Preserve vMensagemListaMsg(iAux)
           
                 vMensagemListaMsg(iAux - 1) = "DID Régua" & _
                 IIf(sCaracteristicaAcesso = "P", " do Acesso Principal", _
                                                             " do Acesso Contingente") & _
                                                             " é obrigatório."
                 
              End If
              
              If objdadosTipoAcessoMaior.DIDPosicao = 0 Then
              
                 bRetorno = False
                 
                 iAux = iAux + 1
            
                 ReDim Preserve vMensagemListaMsg(iAux)
           
                 vMensagemListaMsg(iAux - 1) = "DID Posição" & _
                 IIf(sCaracteristicaAcesso = "P", " do Acesso Principal", _
                                                             " do Acesso Contingente") & _
                                                             " é obrigatório."
                 
              End If
                            
              Set objdadosTipoAcessoMaior = Nothing
           
           ElseIf objDadosPropriedadeTerc.TipoAcesso = "R" Then  'Tipo de Acesso Rede Determ
              
              Set objdadosTipoAcessoRede = New clsGenTipoAcessoRedeDeterm
           
              objdadosTipoAcessoRede.InformaBanco frmgenMDI_SGAPlus.objBanco
           
              bRetornoOK = objdadosTipoAcessoRede.RecuperarDados(lContingenciaAcesso_ID, objMensagem)
              
              If objMensagem.TipoMsg <> "" Then
                 
                 GoTo ErrorHandler
                 
              End If
              
              If objdadosTipoAcessoRede.Dominio = 0 Then
              
                 bRetorno = False
                 iAux = iAux + 1
            
                 ReDim Preserve vMensagemListaMsg(iAux)
           
                 vMensagemListaMsg(iAux - 1) = "Domínio" & _
                 IIf(sCaracteristicaAcesso = "P", " do Acesso Principal", _
                                                             " do Acesso Contingente") & _
                                                             " é obrigatório."
                 
              End If
              
              If objdadosTipoAcessoRede.NO = 0 Then
              
                 bRetorno = False
                 iAux = iAux + 1
            
                 ReDim Preserve vMensagemListaMsg(iAux)
           
                 vMensagemListaMsg(iAux - 1) = "Número NÓ" & _
                 IIf(sCaracteristicaAcesso = "P", " do Acesso Principal", _
                                                             " do Acesso Contingente") & _
                                                             " é obrigatório."
                 
              End If
              
              If objdadosTipoAcessoRede.Slot = "" Then
              
                 bRetorno = False
                 iAux = iAux + 1
            
                 ReDim Preserve vMensagemListaMsg(iAux)
           
                 vMensagemListaMsg(iAux - 1) = "SLOT" & _
                 IIf(sCaracteristicaAcesso = "P", " do Acesso Principal", _
                                                             " do Acesso Contingente") & _
                                                             " é obrigatório."
                 
              End If
              
              If objdadosTipoAcessoRede.Porta = "" Then
              
                 bRetorno = False
                 iAux = iAux + 1
            
                 ReDim Preserve vMensagemListaMsg(iAux)
           
                 vMensagemListaMsg(iAux - 1) = "Porta" & _
                 IIf(sCaracteristicaAcesso = "P", " do Acesso Principal", _
                                                             " do Acesso Contingente") & _
                                                             " é obrigatório."
                 
              End If
              
              If objdadosTipoAcessoRede.TSI = 0 Then
              
                 bRetorno = False
                 iAux = iAux + 1
            
                 ReDim Preserve vMensagemListaMsg(iAux)
           
                 vMensagemListaMsg(iAux - 1) = "Time Slot inicial" & _
                 IIf(sCaracteristicaAcesso = "P", " do Acesso Principal", _
                                                             " do Acesso Contingente") & _
                                                             " é obrigatório."
                 
              End If
              
              If objdadosTipoAcessoRede.TSF = 0 Then
              
                 bRetorno = False
                 
                 iAux = iAux + 1
            
                 ReDim Preserve vMensagemListaMsg(iAux)
           
                 vMensagemListaMsg(iAux - 1) = "Time Slot final" & _
                 IIf(sCaracteristicaAcesso = "P", " do Acesso Principal", _
                                                             " do Acesso Contingente") & _
                                                             " é obrigatório."
                 
              End If
              
              If objdadosTipoAcessoRede.DIDFila = "" Then
              
                 bRetorno = False
                 
                 iAux = iAux + 1
            
                 ReDim Preserve vMensagemListaMsg(iAux)
           
                 vMensagemListaMsg(iAux - 1) = "DID Fila" & _
                 IIf(sCaracteristicaAcesso = "P", " do Acesso Principal", _
                                                             " do Acesso Contingente") & _
                                                             " é obrigatório."
                 
              End If
              
              If objdadosTipoAcessoRede.DIDArmacao = 0 Then
              
                 bRetorno = False
                 
                 iAux = iAux + 1
            
                 ReDim Preserve vMensagemListaMsg(iAux)
           
                 vMensagemListaMsg(iAux - 1) = "DID armação" & _
                 IIf(sCaracteristicaAcesso = "P", " do Acesso Principal", _
                                                             " do Acesso Contingente") & _
                                                             " é obrigatório."
                 
              End If
              
              If objdadosTipoAcessoRede.DIDPosicao = 0 Then
              
                 bRetorno = False
                 
                 iAux = iAux + 1
            
                 ReDim Preserve vMensagemListaMsg(iAux)
           
                 vMensagemListaMsg(iAux - 1) = "DID posição" & _
                 IIf(sCaracteristicaAcesso = "P", " do Acesso Principal", _
                                                             " do Acesso Contingente") & _
                                                             " é obrigatório."
                 
              End If
              
              If objdadosTipoAcessoRede.DIDRegua = 0 Then
              
                 bRetorno = False
                 iAux = iAux + 1
            
                 ReDim Preserve vMensagemListaMsg(iAux)
           
                 vMensagemListaMsg(iAux - 1) = "DID Régua" & _
                 IIf(sCaracteristicaAcesso = "P", " do Acesso Principal", _
                                                             " do Acesso Contingente") & _
                                                             " é obrigatório."
                 
              End If
              
           End If
        
        End If
    
    End If

   Set objAcesso = Nothing
   Set objdadosTipoAcessoRede = Nothing
   Set objDadosPropriedadeTerc = Nothing
   
   If bRetorno Then
      
      VerificaDadosOK = "O"  'O - Dados OK
      
   End If
   
Finalizar:
    
    Exit Function
    
ErrorHandler:

       If objMensagem.TipoMsg = "" Then
          
          objMensagem.TipoMsg = "E"
          objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                             vbCrLf & "Houve um erro de execução no método: " & vbCrLf & vbCrLf & "VerificaDadosOK do Objeto Controle " & _
                             "clsGenOC_TratarAcessoEmbratel"
       
       End If
       
       objMensagem.ExibeMensagem
       
       GoTo Finalizar

End Function


Private Function PopularDadosEmbratel(objPropriedadeEBT As clsGenPropriedadeEmbratel)

    objForm.TxtDesignacaoAcessoEBT.Text = objPropriedadeEBT.DesignacaoAcesso
    objForm.TxtIdentificadorAcessoEBT.Text = objPropriedadeEBT.IdentificadorAcesso
    objForm.TxtEstacaoEntregaEbtConf.Text = objPropriedadeEBT.EstacaoEntregaAcessoConfiguracao
    objForm.TxtEstacaoEntregaEBT.Text = objPropriedadeEBT.EstacaoEntregaAcesso
    
    objForm.TxtEstacaoCliente.Text = objPropriedadeEBT.estacaoCliente
    
    objForm.TxtNumeroOtsTx_EBT.Text = objPropriedadeEBT.NumeroOTSTX
    objForm.TxtItemOtsTxEBT.Text = objPropriedadeEBT.ItemOTSTX
    
    objForm.SpdDataAtivacaoAcessoEBT.Row = 1
    objForm.SpdDataAtivacaoAcessoEBT.Col = 1
    objForm.SpdDataAtivacaoAcessoEBT.Text = objPropriedadeEBT.DataAtivacaoAcesso
    
    If objPropriedadeEBT.TipoAcesso = "R" Then    'Tipo Acesso = Radio
        
        objForm.TxtDIDFila_Radio.Text = objPropriedadeEBT.objTipoAcesso.DIDFila
        objForm.TxtDIDArmacao_Radio.Text = objPropriedadeEBT.objTipoAcesso.DIDArmacao
        objForm.TxtDIDRegua_Radio.Text = objPropriedadeEBT.objTipoAcesso.DIDRegua
        objForm.TxtDIDPosicao_Radio.Text = objPropriedadeEBT.objTipoAcesso.DIDPosicao
        objForm.TxtNumeroRepeticoes = objPropriedadeEBT.objTipoAcesso.NumeroRepeticoes
    
    ElseIf objPropriedadeEBT.TipoAcesso = "F" Then  'Tipo Acesso = Fibra Otica
        
        objForm.TxtDIDFila_Fibra.Text = objPropriedadeEBT.objTipoAcesso.DIDFila
        objForm.TxtDIDArmacao_Fibra.Text = objPropriedadeEBT.objTipoAcesso.DIDArmacao
        objForm.TxtDidRegua_Fibra.Text = objPropriedadeEBT.objTipoAcesso.DIDRegua
        objForm.TxtDIDPosicao_Fibra.Text = objPropriedadeEBT.objTipoAcesso.DIDPosicao
        objForm.TxtNumeroCaboFibra.Text = objPropriedadeEBT.objTipoAcesso.NumeroCabo
        objForm.TxtNumeroParFibra.Text = objPropriedadeEBT.objTipoAcesso.NumeroPar
    
    ElseIf objPropriedadeEBT.TipoAcesso = "A" Then  'Tipo Acesso = ADE
        
        objForm.TxtPredioEBT_ADE.Text = objPropriedadeEBT.objTipoAcesso.PredioEBT
        objForm.TxtNumeroCabo_ADE.Text = objPropriedadeEBT.objTipoAcesso.NumeroCaboEBT
        objForm.TxtNumeroPar_ADE.Text = objPropriedadeEBT.objTipoAcesso.NumeroParEBT
    
    ElseIf objPropriedadeEBT.TipoAcesso = "C" Then  'Tipo Acesso = Cabo
        
        objForm.TxtPredioCabo_Cabo.Text = objPropriedadeEBT.objTipoAcesso.Predio
        objForm.TxtNumeroParEBT_Cabo.Text = objPropriedadeEBT.objTipoAcesso.NumeroParEBT
        objForm.TxtNumeroCaboEBT_Cabo.Text = objPropriedadeEBT.objTipoAcesso.NumeroCaboEBT
        objForm.TxtNumeroParCliente_Cabo.Text = objPropriedadeEBT.objTipoAcesso.NumeroParCliente
        objForm.TxtNumeroCaboCliente_Cabo.Text = objPropriedadeEBT.objTipoAcesso.NumeroCaboCliente
    
    End If

    
End Function


Private Function PopularEquipamentos(objDadosAcesso As clsGenAcesso) As Boolean

    Dim ObjEqpto    As clsGenEqptoComplementar
    Dim iCount      As Integer
    
    iCount = 0
    objForm.ssTabAcesso.TabEnabled(3) = True
    objForm.TxtIDEquipamento1 = ""
    objForm.CboTipoEquipamento_1.ListIndex = -1
    objForm.TxtFabricanteEquipamento_1 = ""
    objForm.optLocalInstalacao_1(0).Value = False
    objForm.optLocalInstalacao_1(1).Value = False
    
    objForm.CboTipoEquipamento_2.ListIndex = -1
    objForm.TxtFabricanteEquipamento_2 = ""
    objForm.optLocalInstalacao_2(0).Value = False
    objForm.optLocalInstalacao_2(1).Value = False
    
    objForm.CboTipoEquipamento_3.ListIndex = -1
    objForm.TxtFabricanteEquipamento_3 = ""
    objForm.optLocalInstalacao_3(0).Value = False
    objForm.optLocalInstalacao_3(1).Value = False
    
    objForm.TxtIDEquipamento1 = ""
    objForm.TxtIDEquipamento2 = ""
    objForm.TxtIDEquipamento3 = ""
    
    'Popula os campos de Equipamento Complementar
    For Each ObjEqpto In objDadosAcesso.objColecaoEqpto
    
        iCount = iCount + 1
        
        Select Case iCount
            Case 1
                'Equipamento 1
                objForm.TxtIDEquipamento1 = Str(ObjEqpto.iD)
                
                If Val(ObjEqpto.TipoEquipamento) <> 0 Then
                   Call SelecionaCombo(objForm.CboTipoEquipamento_1, CLng(ObjEqpto.TipoEquipamento))
                End If
                
                objForm.TxtFabricanteEquipamento_1 = ObjEqpto.Fabricante
                
                If ObjEqpto.LocalInstalacao = "E" Then
                   objForm.optLocalInstalacao_1(0).Value = True
                Else
                   objForm.optLocalInstalacao_1(1).Value = True
                End If
                
            Case 2
                'Equipamento 2
                objForm.TxtIDEquipamento2 = ObjEqpto.iD
                
                If Val(ObjEqpto.TipoEquipamento) <> 0 Then
                    Call SelecionaCombo(objForm.CboTipoEquipamento_2, CLng(ObjEqpto.TipoEquipamento))
                End If
                
                objForm.TxtFabricanteEquipamento_2 = ObjEqpto.Fabricante
                
                If ObjEqpto.LocalInstalacao = "E" Then
                   objForm.optLocalInstalacao_2(0).Value = True
                Else
                   objForm.optLocalInstalacao_2(1).Value = True
                End If
                
            Case 3
                'Equipamento 3
                objForm.TxtIDEquipamento3 = ObjEqpto.iD
                
                If Val(ObjEqpto.TipoEquipamento) <> 0 Then
                    Call SelecionaCombo(objForm.CboTipoEquipamento_3, CLng(ObjEqpto.TipoEquipamento))
                End If
                
                objForm.TxtFabricanteEquipamento_3 = ObjEqpto.Fabricante
                
                If ObjEqpto.LocalInstalacao = "E" Then
                   objForm.optLocalInstalacao_3(0).Value = True
                Else
                   objForm.optLocalInstalacao_3(1).Value = True
                End If
                
        End Select
    
    Next

End Function
Private Function PopularDadosTerceiros(objPropriedadeTerc As clsGenPropriedadeTerceiros)

    objForm.TxtEstacaoEntregaTercConf.Text = objPropriedadeTerc.EstacaoEntregaAcessoConfiguracao
    objForm.TxtEstacaoEntregaTerc.Text = objPropriedadeTerc.EstacaoEntregaAcesso
    objForm.TxtDesignacaoAcessoTerceiro.Text = objPropriedadeTerc.DesignacaoAcesso
    objForm.SpdNumeroAcessoUrbano.Text = objPropriedadeTerc.NumeroAcessoUrbano
    objForm.TxtCNLPontaA_Urbano.Text = objPropriedadeTerc.CNL_Ponta_A_Urbano
    
    objForm.SpdDataEntregaUrbano.Row = 1
    objForm.SpdDataEntregaUrbano.Col = 1
    objForm.SpdDataEntregaUrbano.Text = objPropriedadeTerc.DataEntregaUrbano
   
   If objPropriedadeTerc.TecnologiaAcesso = "L" Then
      
      objForm.optTipoTecnologia(0).Value = True
    
    ElseIf objPropriedadeTerc.TecnologiaAcesso = "R" Then
    
      objForm.optTipoTecnologia(1).Value = True
      
    ElseIf objPropriedadeTerc.TecnologiaAcesso = "F" Then
    
      objForm.optTipoTecnologia(2).Value = True
     
    Else
      
        objForm.optTipoTecnologia(0).Value = False
        objForm.optTipoTecnologia(1).Value = False
        objForm.optTipoTecnologia(2).Value = False
    
    End If
    
    If objPropriedadeTerc.DataAtivacaoAcesso <> "" Then
       
       objForm.SpdDataAtivacaoAcessoTerc.Row = 1
       objForm.SpdDataAtivacaoAcessoTerc.Col = 1
       objForm.SpdDataAtivacaoAcessoTerc.Text = Format( _
                                    objPropriedadeTerc.DataAtivacaoAcesso, "dd/mm/yyyy")
    
    End If
    
    If objPropriedadeTerc.DataEntregaInterurbano <> "" Then
       
        objForm.SpdDataEntregaInterUrbano.Row = 1
        objForm.SpdDataEntregaInterUrbano.Col = 1
        objForm.SpdDataEntregaInterUrbano.Text = Format( _
                                    objPropriedadeTerc.DataEntregaInterurbano, "dd/mm/yyyy")
    
    End If

    objForm.TxtNumeroOTSTXTerc.Text = objPropriedadeTerc.NumeroOTSTX
    objForm.TxtItemOTSTxTerc.Text = IIf(objPropriedadeTerc.ItemOTSTX = 0, "", objPropriedadeTerc.ItemOTSTX)
    
    objForm.TxtNumeroOTSTXTerc.Enabled = False
    objForm.TxtItemOTSTxTerc.Enabled = False
    
    objForm.TxtIDentificadorAcessoUrbano.Text = objPropriedadeTerc.IdentificadorAcessoUrbano

    If objPropriedadeTerc.CodigoProvedorUrbano <> "" Then
       
       Call SelecionaCombo(objForm.CboProvedorUrbano, CLng(objPropriedadeTerc.CodigoProvedorUrbano))
    
    End If
    
    If Val(objPropriedadeTerc.CodigoProvedorInterUrbano) <> 0 Then
       
       objForm.optExisteAcessoIU(0).Value = True
       Call SelecionaCombo(objForm.CboProvedorInterurbano, CLng(objPropriedadeTerc.CodigoProvedorInterUrbano))
       objForm.TxtIdentificadorAcessoEBTInterUrbano.Text = objPropriedadeTerc.IdentificadorAcessoInterUrbano
    
       objForm.SpdNumeroAcessoInterurbano.Text = objPropriedadeTerc.NumeroAcessoInterurbano
       objForm.TxtCNLPontaA.Text = objPropriedadeTerc.CNLPontaA
       objForm.TxtCNLPontaB.Text = objPropriedadeTerc.CNLPontaB
    
    Else
    
       objForm.optExisteAcessoIU(1).Value = True
    
    End If
    
    If objPropriedadeTerc.TipoAcesso = "<" Then     'Tipo Acesso = Menor 2M
        
        If objPropriedadeTerc.objTipoAcesso.InterfaceEBTID <> 0 Then
           
           Call SelecionaCombo(objForm.CboInterfaceEBTMenor, _
                               CLng(objPropriedadeTerc.objTipoAcesso.InterfaceEBTID))
        
        End If
         
        If objPropriedadeTerc.objTipoAcesso.InterfaceClienteID <> 0 Then
           
           Call SelecionaCombo(objForm.CboInterfaceClienteMenor, _
                               CLng(objPropriedadeTerc.objTipoAcesso.InterfaceClienteID))
        
        End If
         
        objForm.TxtNumeroCaboAcessoMenor.Text = objPropriedadeTerc.objTipoAcesso.NumeroCabo
        objForm.TxtNumeroParAcessoMenor.Text = objPropriedadeTerc.objTipoAcesso.NumeroPar
        objForm.TxtBlocoDGMenor.Text = objPropriedadeTerc.objTipoAcesso.BlocoDG
        objForm.TxtParDGMenor.Text = objPropriedadeTerc.objTipoAcesso.ParDG
        objForm.TxtBlocoTransdataMenor.Text = objPropriedadeTerc.objTipoAcesso.BlocoTransData
        objForm.TxtParTransdataMenor.Text = objPropriedadeTerc.objTipoAcesso.ParTransData
        
    ElseIf objPropriedadeTerc.TipoAcesso = ">" Then  'Tipo Acesso = Maior 2M
        
'        objForm.TxtNumeroOTSTXTerc.Enabled = True
'        objForm.TxtItemOTSTxTerc.Enabled = True
        
        With objPropriedadeTerc.objTipoAcesso
            
            objForm.TxtDIDArmacao.Text = IIf(.DIDArmacao = 0, "", .DIDArmacao)
            objForm.TxtDIDFila.Text = .DIDFila
            objForm.TxtDIDPosicao.Text = IIf(.DIDPosicao = 0, "", .DIDPosicao)
            objForm.TxtDIDRegua.Text = IIf(.DIDRegua = 0, "", .DIDRegua)
        
        End With
        
    ElseIf objPropriedadeTerc.TipoAcesso = "R" Then  'Tipo Acesso = Rede Deterministica
        
'        objForm.TxtNumeroOTSTXTerc.Enabled = True
'        objForm.TxtItemOTSTxTerc.Enabled = True
        
        objForm.TxtDominio.Text = objPropriedadeTerc.objTipoAcesso.Dominio
        objForm.txtNo.Text = objPropriedadeTerc.objTipoAcesso.NO
        
        Call SelecionaComboTexto(objForm.CboPorta_Rede, objPropriedadeTerc.objTipoAcesso.Porta)
        Call SelecionaComboTexto(objForm.CboSlot_Rede, objPropriedadeTerc.objTipoAcesso.Slot)
        Call SelecionaComboTexto(objForm.CboTSI_Rede, _
                                 Format(objPropriedadeTerc.objTipoAcesso.TSI, "00"))
        Call SelecionaComboTexto(objForm.CboTSF_Rede, _
                                 Format(objPropriedadeTerc.objTipoAcesso.TSF, "00"))
        
        objForm.TxtDIDArmacao_Rede.Text = objPropriedadeTerc.objTipoAcesso.DIDArmacao
        objForm.TxtDIDFila_Rede.Text = objPropriedadeTerc.objTipoAcesso.DIDFila
        objForm.TxtDIDPosicao_Rede.Text = objPropriedadeTerc.objTipoAcesso.DIDPosicao
        objForm.TxtDIDRegua_Rede.Text = objPropriedadeTerc.objTipoAcesso.DIDRegua
    
    End If

End Function


Public Function LimpaDadosEmbratel()

    Dim iCount      As Integer
    
    objForm.TxtEstacaoEntregaEbtConf.Text = ""
    objForm.TxtEstacaoCliente.Text = ""
    objForm.TxtEstacaoEntregaEBT.Text = ""
    objForm.TxtNumeroOtsTx_EBT.Text = ""

    'Objetos do tipo de acesso Radio
    objForm.TxtDIDFila_Radio.Text = ""
    objForm.TxtDIDArmacao_Radio.Text = ""
    objForm.TxtDIDRegua_Radio.Text = ""
    objForm.TxtDIDPosicao_Radio.Text = ""
    objForm.TxtNumeroRepeticoes.Text = ""
    
    'Objetos do tipo de acesso fibra
    objForm.TxtDIDFila_Fibra.Text = ""
    objForm.TxtDIDArmacao_Fibra.Text = ""
    objForm.TxtDidRegua_Fibra.Text = ""
    objForm.TxtDIDPosicao_Fibra.Text = ""
    objForm.TxtNumeroCaboFibra.Text = ""
    objForm.TxtNumeroParFibra.Text = ""
    
    objForm.TxtPredioEBT_ADE.Text = ""
    objForm.TxtNumeroCabo_ADE.Text = ""
    objForm.TxtNumeroPar_ADE.Text = ""
    
    objForm.TxtPredioCabo_Cabo.Text = ""
    objForm.TxtNumeroParEBT_Cabo.Text = ""
    objForm.TxtNumeroCaboEBT_Cabo.Text = ""
    objForm.TxtNumeroParCliente_Cabo.Text = ""
    objForm.TxtNumeroCaboCliente_Cabo.Text = ""
    
    objForm.TxtIDEquipamento1 = ""
    objForm.CboTipoEquipamento_1.ListIndex = -1
    objForm.TxtFabricanteEquipamento_1 = ""
    objForm.optLocalInstalacao_1(0).Value = False
    objForm.optLocalInstalacao_1(1).Value = False
    
    objForm.CboTipoEquipamento_2.ListIndex = -1
    objForm.TxtFabricanteEquipamento_2 = ""
    objForm.optLocalInstalacao_2(0).Value = False
    objForm.optLocalInstalacao_2(1).Value = False
    
    objForm.CboTipoEquipamento_3.ListIndex = -1
    objForm.TxtFabricanteEquipamento_3 = ""
    objForm.optLocalInstalacao_3(0).Value = False
    objForm.optLocalInstalacao_3(1).Value = False
    
    objForm.TxtIDEquipamento1 = ""
    objForm.TxtIDEquipamento2 = ""
    objForm.TxtIDEquipamento3 = ""
    
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

Public Sub CarregarComboClassifEndereco(cboLista As ComboBox)
    
    cboLista.Clear
    cboLista.AddItem ""
    cboLista.ItemData(0) = 0
    
    cboLista.AddItem "PPC"
    cboLista.ItemData(1) = 1
    
    cboLista.AddItem "POE"
    cboLista.ItemData(1) = 2
    
    cboLista.AddItem "CLIENTE"
    cboLista.ItemData(1) = 3
    
    cboLista.AddItem "ATENDIDO POR PPC OUTRO PRÉDIO"
    cboLista.ItemData(1) = 4
    
End Sub

Public Sub CarregarComboTipoAcessoEBT(cboLista As ComboBox)
    
    cboLista.Clear
    
    cboLista.AddItem "Rádio"
    cboLista.ItemData(0) = 0
    
    cboLista.AddItem "Fibra Ótica"
    cboLista.ItemData(1) = 1
    
    cboLista.AddItem "ADE"
    cboLista.ItemData(2) = 2
    
    If Produto_ID <> GbIDProdutoVIP Then
        
        cboLista.AddItem "Cabo Interno"
        cboLista.ItemData(1) = 3
    
    End If
    
End Sub

Public Function CarregarComboVelocidade(objDadosAcesso As clsGenAcesso, _
                                        psVelocidade As String, _
                                        RefCombo As String, _
                                        cboLista As ComboBox) As Boolean

On Error GoTo ErrorHandler

    Dim bRetorno            As Boolean
    Dim iCont               As Integer
    Dim aLista()            As Variant
    Dim iDigitoVelocCCTO    As Integer
    Dim sLetraVelocCCTO     As String
    Dim iDigitoVelocTab     As Integer
    Dim sLetraVelocTab      As String
    Dim IncluiNaCombo       As Boolean
    Dim iAux                As Integer
    Dim objDadosProduto     As New clsGenProduto
    Dim sSiglaProduto       As String
    
    CarregarComboVelocidade = False
    
    iAux = 0
        
    objDadosProduto.InformaBanco frmgenMDI_SGAPlus.objBanco
    objDadosProduto.InformaMensagem objMensagem
    
    bRetorno = objDadosProduto.RecuperarDados(Produto_ID)
    
    If bRetorno Then
    
        sSiglaProduto = objDadosProduto.Sigla
        
    End If
    
    If sSiglaProduto = "WD" Then
    
        sSiglaProduto = "WNN"
        
    ElseIf sSiglaProduto = "BF" Then
    
        sSiglaProduto = "BFN"
        
        
    ElseIf sSiglaProduto = "VN" Then
        'IP VPN
        'Com base no documento de interface SOE obter a sigla do IP VPN
    
        sSiglaProduto = GbCodigo_Produto_IPVPN
        
    End If
    
    'Limpa a combo
    bRetorno = objDadosAcesso.MontarListasAcesso(RefCombo, aLista, sSiglaProduto)
    
    If Not bRetorno Then

        GoTo ErrorHandler

    End If
    
    'Ler o array para preencher a combo
    Call obtemDigitoVelocidade(psVelocidade, iDigitoVelocCCTO, sLetraVelocCCTO)
      
    cboLista.Clear
    cboLista.AddItem ""
    cboLista.ItemData(0) = 0
    
    For iCont = 1 To UBound(aLista, 2)

        IncluiNaCombo = False
        
        If Not IsEmpty(aLista(0, iCont)) And IsNumeric(aLista(0, iCont)) Then
            
            If Trim(psVelocidade) <> "" Then
            
                Call obtemDigitoVelocidade(CStr(aLista(1, iCont)), iDigitoVelocTab, sLetraVelocTab)
                
                If sLetraVelocCCTO = "M" And sLetraVelocTab = "M" Then
                    
                   If iDigitoVelocTab >= iDigitoVelocCCTO Then
                     
                      IncluiNaCombo = True
                     
                   Else
                     
                      IncluiNaCombo = False
                  
                   End If
                    
                Else
                   
                   If sLetraVelocCCTO = "K" Then
                         
                      If sLetraVelocTab = "M" Then
                            
                         IncluiNaCombo = True
                      
                      Else
                      
                         If iDigitoVelocTab >= iDigitoVelocCCTO Then
                         
                            IncluiNaCombo = True
                       
                         Else
                       
                            IncluiNaCombo = False
                    
                         End If
                      
                      End If
                      
                   End If
                    
                End If
            
            Else
            
                IncluiNaCombo = True
            
            End If
            
            If InStr(1, aLista(1, iCont), "-") > 0 Then
            
                IncluiNaCombo = False
                
            End If
            
            If IncluiNaCombo Then
               
               iAux = iAux + 1
               cboLista.AddItem aLista(1, iCont)
               cboLista.ItemData(iAux) = CInt(aLista(0, iCont))
                
            Else
                If sSiglaProduto = "VLI" And sLetraVelocCCTO = "" Then
                   iAux = iAux + 1
                   cboLista.AddItem aLista(1, iCont)
                   cboLista.ItemData(iAux) = CInt(aLista(0, iCont))
                End If
                
            End If
            
        End If

    Next
        
    CarregarComboVelocidade = True
       
Finalizar:
    
    On Error GoTo 0
    
    Exit Function
        
ErrorHandler:

    If objMensagem.TipoMsg = "" Then
       
       objMensagem.TipoMsg = "E"
       objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & _
                   vbCrLf & vbCrLf & _
                   "Houve um erro de execução no Método: " & vbCrLf & vbCrLf & _
                   "CarregarComboVelocidade " & _
                   "da classe clsGenOC_TratarAcessoEmbratel"
    
    End If
    
    GoTo Finalizar

End Function

Public Function CarregarComboVelocidade_DNI(objDadosAcesso As clsGenAcesso, _
                                            psVelocidade As String, _
                                            RefCombo As String, _
                                            cboLista As ComboBox) As Boolean

On Error GoTo ErrorHandler

    Dim bRetorno            As Boolean
    Dim iCont               As Integer
    Dim aLista()            As Variant
    Dim iDigitoVelocCCTO    As Integer
    Dim sLetraVelocCCTO     As String
    Dim iDigitoVelocTab     As Integer
    Dim sLetraVelocTab      As String
    Dim IncluiNaCombo       As Boolean
    Dim iAux                As Integer
    Dim objDadosProduto     As New clsGenProduto
    Dim sSiglaProduto       As String
    
    CarregarComboVelocidade_DNI = False
    
    objDadosProduto.InformaBanco frmgenMDI_SGAPlus.objBanco
    objDadosProduto.InformaMensagem objMensagem
    
    bRetorno = objDadosProduto.RecuperarDados(Produto_ID)
    
    If bRetorno Then
    
        sSiglaProduto = objDadosProduto.Sigla
        
    End If
    
    'Limpa a combo
    bRetorno = objDadosAcesso.MontarListasAcesso(RefCombo, aLista, sSiglaProduto)
    
    If Not bRetorno Then

        GoTo ErrorHandler

    End If
    
    ReDim vVelocidadeDNI(1) As StrGen
    
    iAux = 1

    vVelocidadeDNI(0).CodigoAlpha = ""
    vVelocidadeDNI(0).Descricao = ""
    
    'Ler o array para preencher a combo
    Call obtemDigitoVelocidade(psVelocidade, iDigitoVelocCCTO, sLetraVelocCCTO)
      
    cboLista.Clear
    cboLista.AddItem ""
    cboLista.ItemData(0) = 0
    
    For iCont = 1 To UBound(aLista, 2)

        IncluiNaCombo = False
        
        If Not IsEmpty(aLista(0, iCont)) Then
            
            Call obtemDigitoVelocidade(Mid(CStr(aLista(1, iCont)), 1, 6), iDigitoVelocTab, sLetraVelocTab)
            
            If sLetraVelocCCTO = "M" And sLetraVelocTab = "M" Then
                
               If iDigitoVelocTab >= iDigitoVelocCCTO Then
                 
                  IncluiNaCombo = True
                 
               Else
                 
                  IncluiNaCombo = False
              
               End If
                
            Else
               
               If sLetraVelocCCTO = "K" Then
                     
                  If sLetraVelocTab = "M" Then
                        
                     IncluiNaCombo = True
                  
                  Else
                  
                     If iDigitoVelocTab >= iDigitoVelocCCTO Then
                     
                        IncluiNaCombo = True
                   
                     Else
                   
                        IncluiNaCombo = False
                
                     End If
                  
                  End If
                  
               End If
                
            End If
            
            If InStr(1, aLista(1, iCont), "-") > 0 Then
            
                IncluiNaCombo = False
                
            End If
            
            If IncluiNaCombo Then
               
               iAux = iAux + 1
               Call obtemDigitoVelocidade(Mid(CStr(aLista(0, iCont)), 1, 6), iDigitoVelocTab, sLetraVelocTab)
               
               cboLista.AddItem aLista(1, iCont)
               
               ReDim Preserve vVelocidadeDNI(iAux) As StrGen
               
               vVelocidadeDNI(iAux - 1).CodigoAlpha = aLista(0, iCont)
               vVelocidadeDNI(iAux - 1).Descricao = aLista(1, iCont)

            End If
            
        End If

    Next
        
    CarregarComboVelocidade_DNI = True
       
Finalizar:
    
    On Error GoTo 0
    
    Exit Function
        
ErrorHandler:

    If objMensagem.TipoMsg = "" Then
       
       objMensagem.TipoMsg = "E"
       objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & _
                   vbCrLf & vbCrLf & _
                   "Houve um erro de execução no Método: " & vbCrLf & vbCrLf & _
                   "CarregarComboVelocidade " & _
                   "da classe clsGenOC_TratarAcessoEmbratel"
    
    End If
    
    GoTo Finalizar

End Function

Private Function CarregaAtributosPropriedadesEBT(objDadosAcesso As clsGenAcesso)

    Dim objDadosTipoRadio      As clsGenTipoAcessoRadio
    Dim objDadosTipoFibra      As clsGenTipoAcessoFibra
    Dim objdadosTipoADE        As clsGenTipoAcessoADE
    Dim objDadosTipoCabo       As clsGenTipoAcessoCabo
    Dim objDadosPropriedadeEBT As New clsGenPropriedadeEmbratel
    
    'Carrega Propriedades da Pasta Propriedade Embratel
    
    With objDadosPropriedadeEBT
         
         .EstacaoEntregaAcessoConfiguracao = TrataPlick(Trim(objForm.TxtEstacaoEntregaEbtConf.Text))
         .EstacaoEntregaAcesso = TrataPlick(Trim(objForm.TxtEstacaoEntregaEBT.Text))
         .estacaoCliente = TrataPlick(Trim(objForm.TxtEstacaoCliente.Text))
         .NumeroOTSTX = TrataPlick(objForm.TxtNumeroOtsTx_EBT.Text)
         .ItemOTSTX = Val(objForm.TxtItemOtsTxEBT.Text)
         
         If objForm.SpdDataAtivacaoAcessoEBT.Text <> "" Then
           
            .DataAtivacaoAcesso = Format(objForm.SpdDataAtivacaoAcessoEBT.Text, "dd/mm/yyyy")
        
         Else
            
            .DataAtivacaoAcesso = ""
         
         End If
        
         .DesignacaoAcesso = TrataPlick(objForm.TxtDesignacaoAcessoEBT.Text)
         .IdentificadorAcesso = TrataPlick(objForm.TxtIdentificadorAcessoEBT.Text)
         
         Select Case objForm.CboTipoAcessoEmbratel.ListIndex
                
                Case 0       'Tipo de Acesso Radio
                    
                    .TipoAcesso = "R"
                    
                    Set objDadosTipoRadio = New clsGenTipoAcessoRadio
                    
                    objDadosTipoRadio.DIDFila = "" & TrataPlick(objForm.TxtDIDFila_Radio.Text)
                    objDadosTipoRadio.DIDArmacao = Val(objForm.TxtDIDArmacao_Radio.Text)
                    objDadosTipoRadio.DIDRegua = Val(objForm.TxtDIDRegua_Radio.Text)
                    objDadosTipoRadio.DIDPosicao = Val(objForm.TxtDIDPosicao_Radio.Text)
                    objDadosTipoRadio.NumeroRepeticoes = Val(objForm.TxtNumeroRepeticoes.Text)
                    
                    .objTipoAcesso = objDadosTipoRadio
                    Set objDadosTipoRadio = Nothing
                    
                Case 1      'Tipo de Acesso Fibra Otica
                     .TipoAcesso = "F"
                     
                     Set objDadosTipoFibra = New clsGenTipoAcessoFibra
                     
                     objDadosTipoFibra.DIDFila = "" & TrataPlick(Trim(objForm.TxtDIDFila_Fibra.Text))
                     objDadosTipoFibra.DIDArmacao = Val(objForm.TxtDIDArmacao_Fibra.Text)
                     objDadosTipoFibra.DIDRegua = Val(objForm.TxtDidRegua_Fibra.Text)
                     objDadosTipoFibra.DIDPosicao = Val(objForm.TxtDIDPosicao_Fibra.Text)
                     objDadosTipoFibra.NumeroCabo = "" & TrataPlick(Trim(objForm.TxtNumeroCaboFibra.Text))
                     objDadosTipoFibra.NumeroPar = "" & TrataPlick(Trim(objForm.TxtNumeroParFibra.Text))
                     
                     .objTipoAcesso = objDadosTipoFibra
                     
                     Set objDadosTipoFibra = Nothing
                     
                Case 2   'Tipo de Acesso ADE
                     .TipoAcesso = "A"
                     
                     Set objdadosTipoADE = New clsGenTipoAcessoADE
                     
                     objdadosTipoADE.PredioEBT = "" & TrataPlick(Trim(objForm.TxtPredioEBT_ADE.Text))
                     objdadosTipoADE.NumeroCaboEBT = Val(objForm.TxtNumeroCabo_ADE.Text)
                     objdadosTipoADE.NumeroParEBT = "" & TrataPlick(Trim(objForm.TxtNumeroPar_ADE.Text))
                     
                     .objTipoAcesso = objdadosTipoADE
                     Set objdadosTipoADE = Nothing
                     
                Case 3     'Tipo de Acesso Cabo Interno
                     
                     .TipoAcesso = "C"
                     
                     Set objDadosTipoCabo = New clsGenTipoAcessoCabo
                     
                     objDadosTipoCabo.Predio = "" & TrataPlick(Trim(objForm.TxtPredioCabo_Cabo.Text))
                     objDadosTipoCabo.NumeroCaboEBT = "" & TrataPlick(Trim(objForm.TxtNumeroCaboEBT_Cabo.Text))
                     objDadosTipoCabo.NumeroParEBT = "" & TrataPlick(Trim(objForm.TxtNumeroParEBT_Cabo.Text))
                     objDadosTipoCabo.NumeroCaboCliente = "" & TrataPlick(Trim(objForm.TxtNumeroCaboCliente_Cabo.Text))
                     objDadosTipoCabo.NumeroParCliente = "" & TrataPlick(Trim(objForm.TxtNumeroParCliente_Cabo.Text))
                     
                     .objTipoAcesso = objDadosTipoCabo
                     Set objDadosTipoCabo = Nothing
                     
         End Select
    
    End With
    
    objDadosAcesso.objPropriedadeEBT = objDadosPropriedadeEBT
    
    Set objDadosPropriedadeEBT = Nothing
    
End Function

Private Function CarregaAtributosPropriedadesTerceiros(objPropriedadeTerceiros As clsGenPropriedadeTerceiros) As Boolean

    Dim objdadosTipoAcessoMenor As clsGenTipoAcessoMenor2M
    Dim objdadosTipoAcessoMaior As clsGenTipoAcessoMaior2M
    Dim objdadosTipoAcessoRede  As clsGenTipoAcessoRedeDeterm
    
    With objPropriedadeTerceiros

         .CNLPontaA = TrataPlick(objForm.TxtCNLPontaA.Text)
         .CNLPontaB = TrataPlick(objForm.TxtCNLPontaB.Text)

         If objForm.CboProvedorInterurbano.ListIndex >= 0 Then
         
            .CodigoProvedorInterUrbano = objForm.CboProvedorInterurbano.ItemData(objForm.CboProvedorInterurbano.ListIndex)
         
         Else
            
            .CodigoProvedorInterUrbano = "0"
         
         End If
         
         If objForm.CboProvedorUrbano.ListIndex >= 0 Then
            
            .CodigoProvedorUrbano = objForm.CboProvedorUrbano.ItemData(objForm.CboProvedorUrbano.ListIndex)
         
         Else
            
            .CodigoProvedorUrbano = "0"
         
         End If
         
         If Trim(objForm.SpdDataEntregaInterUrbano.Text) <> "" Then
            
            .DataEntregaInterurbano = Format(objForm.SpdDataEntregaInterUrbano.Text, "dd/mm/yyyy")
         
         Else
         
            .DataEntregaInterurbano = ""
            
         End If
         
         .EstacaoEntregaAcessoConfiguracao = TrataPlick(objForm.TxtEstacaoEntregaTercConf.Text)
         .EstacaoEntregaAcesso = TrataPlick(objForm.TxtEstacaoEntregaTerc.Text)
         .NumeroAcessoUrbano = TrataPlick(objForm.SpdNumeroAcessoUrbano.Text)
         .CNL_Ponta_A_Urbano = TrataPlick(objForm.TxtCNLPontaA_Urbano.Text)
         .DataEntregaUrbano = Format(objForm.SpdDataEntregaUrbano.Text, "dd/mm/yyyy")
    
         .NumeroAcessoInterurbano = TrataPlick(objForm.SpdNumeroAcessoInterurbano.Text)
         .NumeroOTSTX = TrataPlick(objForm.TxtNumeroOTSTXTerc.Text)
         .ItemOTSTX = Val(objForm.TxtItemOTSTxTerc.Text)
         .DesignacaoAcesso = TrataPlick(objForm.TxtDesignacaoAcessoTerceiro.Text)
'         .DesignacaoAcesso = ""
         .IdentificadorAcessoUrbano = TrataPlick(objForm.TxtIDentificadorAcessoUrbano.Text)
         .IdentificadorAcessoInterUrbano = TrataPlick(objForm.TxtIdentificadorAcessoEBTInterUrbano.Text)
         
         If objForm.SpdDataAtivacaoAcessoTerc.Text <> "" Then
            
            .DataAtivacaoAcesso = Format(objForm.SpdDataAtivacaoAcessoTerc.Text, "dd/mm/yyyy")
         
         Else
            
            .DataAtivacaoAcesso = ""
         
         End If
                     
         If objForm.optTipoTecnologia(0).Value = True Then
         
            .TecnologiaAcesso = "L"    'LP
            
         ElseIf objForm.optTipoTecnologia(1).Value = True Then
         
            .TecnologiaAcesso = "R"    'Radio
         
         ElseIf objForm.optTipoTecnologia(2).Value = True Then
         
            .TecnologiaAcesso = "F"    'Fibra
         
         End If
         
         Select Case objForm.CboTipoAcessoTerceiro.ListIndex

                Case 0  'Tipo de Acesso Menor 2M

                     Set objdadosTipoAcessoMenor = New clsGenTipoAcessoMenor2M
                     .objTipoAcesso = objdadosTipoAcessoMenor
                     Set objdadosTipoAcessoMenor = Nothing

                     .TipoAcesso = "<"
                     
                     If objForm.CboInterfaceEBTMenor.ListIndex >= 0 Then
                        
                        .objTipoAcesso.InterfaceEBTID = _
                                                 objForm.CboInterfaceEBTMenor.ItemData( _
                                                 objForm.CboInterfaceEBTMenor.ListIndex)
                     Else
                        
                        .objTipoAcesso.InterfaceEBTID = 0
                     
                     End If
                     
                     
                     If objForm.CboInterfaceClienteMenor.ListIndex >= 0 Then
                        .objTipoAcesso.InterfaceClienteID = _
                                                 objForm.CboInterfaceClienteMenor.ItemData( _
                                                 objForm.CboInterfaceClienteMenor.ListIndex)
                     Else
                        
                        .objTipoAcesso.InterfaceClienteID = 0
                     
                     End If
                     
                     .objTipoAcesso.NumeroCabo = TrataPlick(objForm.TxtNumeroCaboAcessoMenor.Text)
                     .objTipoAcesso.NumeroPar = TrataPlick(objForm.TxtNumeroParAcessoMenor.Text)
                     .objTipoAcesso.BlocoDG = TrataPlick(objForm.TxtBlocoDGMenor.Text)
                     .objTipoAcesso.BlocoTransData = TrataPlick(objForm.TxtBlocoTransdataMenor.Text)
                     .objTipoAcesso.ParDG = TrataPlick(objForm.TxtParDGMenor.Text)
                     .objTipoAcesso.ParTransData = TrataPlick(objForm.TxtParTransdataMenor.Text)
                     
                
                Case 1  'Tipo de Acesso Maior 2M

                     Set objdadosTipoAcessoMaior = New clsGenTipoAcessoMaior2M
                     .objTipoAcesso = objdadosTipoAcessoMaior
                     Set objdadosTipoAcessoMenor = Nothing
                     
                     .TipoAcesso = ">"
                     .objTipoAcesso.DIDArmacao = Val(objForm.TxtDIDArmacao.Text)
                     .objTipoAcesso.DIDFila = TrataPlick(objForm.TxtDIDFila.Text)
                     .objTipoAcesso.DIDPosicao = Val(objForm.TxtDIDPosicao.Text)
                     .objTipoAcesso.DIDRegua = Val(objForm.TxtDIDRegua.Text)
                
                Case 2  'Tipo de Acesso Rede Deterministica

                     Set objdadosTipoAcessoRede = New clsGenTipoAcessoRedeDeterm
                     .objTipoAcesso = objdadosTipoAcessoRede
                     Set objdadosTipoAcessoRede = Nothing
                     
                     .TipoAcesso = "R"
                     
                     .objTipoAcesso.Dominio = Val(objForm.TxtDominio.Text)
                     .objTipoAcesso.NO = Val(objForm.txtNo.Text)
                     
                     .objTipoAcesso.Porta = objForm.CboPorta_Rede.Text
                     .objTipoAcesso.Slot = objForm.CboSlot_Rede.Text
                     .objTipoAcesso.TSI = Val(objForm.CboTSI_Rede.Text)
                     .objTipoAcesso.TSF = Val(objForm.CboTSF_Rede.Text)

                     .objTipoAcesso.DIDArmacao = Val(objForm.TxtDIDArmacao_Rede.Text)
                     .objTipoAcesso.DIDFila = TrataPlick(objForm.TxtDIDFila_Rede.Text)
                     .objTipoAcesso.DIDPosicao = Val(objForm.TxtDIDPosicao_Rede.Text)
                     .objTipoAcesso.DIDRegua = Val(objForm.TxtDIDRegua_Rede.Text)
         
         End Select

    End With

End Function

Public Function CarregaAtributosEqtos(objDadosAcesso As clsGenAcesso) As Boolean
    
    Dim ObjEqpto               As clsGenEqptoComplementar
    Dim cColecao               As New Collection
    Dim FlagTemEqpto           As Boolean

     FlagTemEqpto = False
     
     Set ObjEqpto = New clsGenEqptoComplementar
     
     'Carrega propriedades do Equipamento 1
     If objForm.optLocalInstalacao_1(0).Value = True Then
        ObjEqpto.LocalInstalacao = "E"

     Else

         If objForm.optLocalInstalacao_1(1).Value = True Then
            ObjEqpto.LocalInstalacao = "C"

         Else
            ObjEqpto.LocalInstalacao = ""

         End If

     End If

     If objForm.CboTipoEquipamento_1.ListIndex > 0 Then
        ObjEqpto.TipoEquipamento = objForm.CboTipoEquipamento_1.ItemData(objForm.CboTipoEquipamento_1.ListIndex)

     Else
        ObjEqpto.TipoEquipamento = ""

     End If

     ObjEqpto.iD = Val(objForm.TxtIDEquipamento1.Text)
     ObjEqpto.Fabricante = Trim(TrataPlick(objForm.TxtFabricanteEquipamento_1.Text))

     cColecao.Add ObjEqpto
     Set ObjEqpto = New clsGenEqptoComplementar

     'Carrega propriedades do Equipamento 2
     If objForm.optLocalInstalacao_2(0).Value = True Then
        
        ObjEqpto.LocalInstalacao = "E"

     Else

         If objForm.optLocalInstalacao_2(1).Value = True Then
            ObjEqpto.LocalInstalacao = "C"

         Else
            ObjEqpto.LocalInstalacao = ""

         End If

     End If

     If objForm.CboTipoEquipamento_2.ListIndex > 0 Then
        ObjEqpto.TipoEquipamento = objForm.CboTipoEquipamento_2.ItemData(objForm.CboTipoEquipamento_2.ListIndex)

     Else
        ObjEqpto.TipoEquipamento = ""

     End If

     ObjEqpto.iD = Val(objForm.TxtIDEquipamento2.Text)
     
     ObjEqpto.Fabricante = TrataPlick(Trim(objForm.TxtFabricanteEquipamento_2.Text))

     cColecao.Add ObjEqpto
     Set ObjEqpto = New clsGenEqptoComplementar

     'Carrega propriedades do Equipamento 3
     If objForm.optLocalInstalacao_3(0).Value = True Then
        ObjEqpto.LocalInstalacao = "E"

     Else

         If objForm.optLocalInstalacao_3(1).Value = True Then
            ObjEqpto.LocalInstalacao = "C"

         Else
            ObjEqpto.LocalInstalacao = ""

         End If

     End If

     If objForm.CboTipoEquipamento_3.ListIndex > 0 Then
        ObjEqpto.TipoEquipamento = objForm.CboTipoEquipamento_3.ItemData(objForm.CboTipoEquipamento_3.ListIndex)

     Else
        ObjEqpto.TipoEquipamento = ""

     End If

     ObjEqpto.iD = Val(objForm.TxtIDEquipamento3.Text)
     
     ObjEqpto.Fabricante = TrataPlick(Trim(objForm.TxtFabricanteEquipamento_3.Text))

     
     cColecao.Add ObjEqpto
     
     objDadosAcesso.objColecaoEqpto = cColecao
    
     Set cColecao = Nothing

End Function

Public Sub SelecionaComboTexto(cbo As ComboBox, pRef As String)

    Dim i As Integer
    
    cbo.ListIndex = -1
    
    If pRef = "" Then Exit Sub
    
    For i = 0 To cbo.ListCount
    
        If cbo.List(i) = pRef Then
           
           cbo.ListIndex = i
           
           Exit Sub
        
        End If
    
    Next i
    
End Sub

Public Sub SelecionaComboTexto_DNI(cbo As ComboBox, pRef As String)

    Dim i As Integer
    
    If pRef = "" Then Exit Sub
    
    For i = 0 To UBound(vVelocidadeDNI)
    
        If Trim(vVelocidadeDNI(i).CodigoAlpha) = Trim(pRef) Then
           
           cbo.ListIndex = i
           
           Exit Sub
        
        End If
    
    Next i
    
End Sub

Public Function GravaDadosPropriedadeEmbratel(objDadosAcesso As clsGenAcesso) As Boolean

    Dim bRetorno              As Boolean
    Dim lContingenciaAcesso   As Long
    Dim FlJaexiste            As Boolean
    Dim objPropriedadeTerc    As clsGenPropriedadeTerceiros

On Error GoTo ErrorHandler
            
    lContingenciaAcesso = objDadosAcesso.iD
            
    GravaDadosPropriedadeEmbratel = False
    
    objDadosAcesso.objPropriedadeEBT.InformaBanco frmgenMDI_SGAPlus.objBanco
    
    FlJaexiste = objDadosAcesso.objPropriedadeEBT.FaseExistente(lContingenciaAcesso)
    
    If objMensagem.TipoMsg <> "" Then

       GoTo ErrorHandler
      
    End If
    
    'Caso o Tipo de Acesso seja o mesmo, realiza somente uma Atualiza
    'Caso contrario, exclui o tipo anterior e incluir o novo
    
    If objDadosAcesso.objPropriedadeEBT.TipoAcesso <> TipoAcessoAnterior Then
        
        'Exclui o tipo de Acesso anterior
        bRetorno = ExcluiTipoAcessoEBT(TipoAcessoAnterior, lContingenciaAcesso, objMensagem)
    
    End If
    
    If Not FlJaexiste Then  'Inclusao
       
       bRetorno = objDadosAcesso.objPropriedadeEBT.IncluirDados(lContingenciaAcesso, objMensagem)
    
       If (objMensagem.TipoMsg <> "") Then
     
          GoTo ErrorHandler
           
       End If
       
       'Caso o usuário tenha alterado a propriedade de Acesso para a mesma caracteristica
       'de acesso, exclui os dados da propriedade anterior
       
       If PropriedadeAcessoAnterior <> "" And _
          CaracteristicaAcessoAnterior = objDadosAcesso.CaracteristicaAcesso Then
          
          Set objPropriedadeTerc = New clsGenPropriedadeTerceiros
        
          'Exclui dados de Propriedade Terceiros
          
         bRetorno = ExcluiTipoAcessoTerceiros(TipoAcessoAnterior, lContingenciaAcesso, objMensagem)
          
         If objMensagem.TipoMsg <> "" Then
       
             GoTo ErrorHandler
             
         End If
         
         objPropriedadeTerc.InformaBanco frmgenMDI_SGAPlus.objBanco
            
          bRetorno = objPropriedadeTerc.ExcluirDados(lContingenciaAcesso, objMensagem)
          
          If objMensagem.TipoMsg <> "" Then
       
             GoTo ErrorHandler
             
         End If
       
         TipoAcessoAnterior = ""
       
       End If
       
    Else
       
       bRetorno = objDadosAcesso.objPropriedadeEBT.AtualizarDados(lContingenciaAcesso, objMensagem)
    
    End If
    
    If Not bRetorno Then
    
        GoTo ErrorHandler
    
    End If
    
    'Grava dados de Tipo de Acesso
    objDadosAcesso.objPropriedadeEBT.objTipoAcesso.InformaBanco frmgenMDI_SGAPlus.objBanco
       
    If Not objDadosAcesso.objPropriedadeEBT.objTipoAcesso.VerificaExistencia(lContingenciaAcesso, objMensagem) Then
    
       bRetorno = objDadosAcesso.objPropriedadeEBT.objTipoAcesso.IncluirDados( _
                                                        lContingenciaAcesso, objMensagem)
    Else
     
       bRetorno = objDadosAcesso.objPropriedadeEBT.objTipoAcesso.AtualizarDados( _
                                                        lContingenciaAcesso, objMensagem)
     
    End If
    
    If objMensagem.TipoMsg <> "" Then
    
        GoTo ErrorHandler
    
    End If

    GravaDadosPropriedadeEmbratel = True

Finalizar:
    
    On Error GoTo 0
    
    Exit Function
    
ErrorHandler:
       
    If objMensagem.TipoMsg = "" Then
       
       objMensagem.TipoMsg = "E"
       objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                   "Houve um erro de execução no Método: " & vbCrLf & vbCrLf & "GravarDadosPropriedadeEmbratel " & _
                   "da classe clsGenOC_TratarAcessoEmbratel"
    
    End If
       
    GoTo Finalizar
    
End Function

Private Function GravaDadosPropriedadeTerceiros(objDadosAcesso As clsGenAcesso) As Boolean

    Dim bRetorno              As Boolean
    Dim lContingenciaAcesso   As Long
    Dim FlJaexiste            As Boolean
    Dim objPropriedadeEBT     As clsGenPropriedadeEmbratel

On Error GoTo ErrorHandler
            
    lContingenciaAcesso = objDadosAcesso.iD
            
    GravaDadosPropriedadeTerceiros = False
    
    objDadosAcesso.objPropriedadeTerceiros.InformaBanco frmgenMDI_SGAPlus.objBanco
    
    'Verifica se já existe dados para a fase
    FlJaexiste = objDadosAcesso.objPropriedadeTerceiros.FaseExistente(lContingenciaAcesso, objMensagem)
    
    If objMensagem.TipoMsg <> "" Then
    
        GoTo ErrorHandler
    
    End If
    
    If objDadosAcesso.objPropriedadeTerceiros.TipoAcesso <> TipoAcessoAnterior Then
        
       ' Exclui o tipo de Acesso anterior
       bRetorno = ExcluiTipoAcessoTerceiros(TipoAcessoAnterior, lContingenciaAcesso, _
                                              objMensagem)
     
    End If
    
    If Not FlJaexiste Then  'Inclusao
       
       bRetorno = objDadosAcesso.objPropriedadeTerceiros.IncluirDados(lContingenciaAcesso, objMensagem)
       
       If objMensagem.TipoMsg <> "" Then
     
          GoTo ErrorHandler
           
       End If
    
       'Caso o usuário tenha alterado a propriedade de Acesso para a mesma caracteristica
       'de acesso, exclui os dados da propriedade anterior
       
       If PropriedadeAcessoAnterior <> "" And _
          CaracteristicaAcessoAnterior = objDadosAcesso.CaracteristicaAcesso Then
          
          Set objPropriedadeEBT = New clsGenPropriedadeEmbratel
        
          'Exclui dados de Propriedade EBT
          
         bRetorno = ExcluiTipoAcessoEBT(TipoAcessoAnterior, lContingenciaAcesso, objMensagem)
          
         If objMensagem.TipoMsg <> "" Then
       
             GoTo ErrorHandler
             
         End If
         
         objPropriedadeEBT.InformaBanco frmgenMDI_SGAPlus.objBanco
            
          bRetorno = objPropriedadeEBT.ExcluirDados(lContingenciaAcesso, objMensagem)
          
          If objMensagem.TipoMsg <> "" Then
       
             GoTo ErrorHandler
             
         End If
       
         TipoAcessoAnterior = ""
       
       End If
    
    Else
       
       bRetorno = objDadosAcesso.objPropriedadeTerceiros.AtualizarDados(lContingenciaAcesso, objMensagem)
    
    End If
    
    If objMensagem.TipoMsg <> "" Then
    
        GoTo ErrorHandler
    
    End If
    
    'Grava dados de Tipo de Acesso
    objDadosAcesso.objPropriedadeTerceiros.objTipoAcesso.InformaBanco frmgenMDI_SGAPlus.objBanco
       
    'Caso o Tipo de Acesso seja o mesmo, realiza somente uma Atualizacao
    'Caso contrario, exclui o tipo anterior e incluir o novo
       
    If Not objDadosAcesso.objPropriedadeTerceiros.objTipoAcesso.VerificaExistencia( _
                                                   lContingenciaAcesso, objMensagem) Then
    
        bRetorno = objDadosAcesso.objPropriedadeTerceiros.objTipoAcesso.IncluirDados( _
                                                       lContingenciaAcesso, objMensagem)
    
    Else
     
        bRetorno = objDadosAcesso.objPropriedadeTerceiros.objTipoAcesso.AtualizarDados( _
                                                       lContingenciaAcesso, objMensagem)
    
    End If
       
    If objMensagem.TipoMsg <> "" Then
    
        GoTo ErrorHandler
    
    End If

    GravaDadosPropriedadeTerceiros = True

Finalizar:
    
    On Error GoTo 0
    
    Exit Function
    
ErrorHandler:
       
    If objMensagem.TipoMsg = "" Then
       
       objMensagem.TipoMsg = "E"
       objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                   "Houve um erro de execução no Método: " & vbCrLf & vbCrLf & "GravaDadosPropriedadeTerceiros " & _
                   "da classe clsGenOC_TratarAcessoEmbratel"
    End If
       
    GoTo Finalizar
    
End Function


Public Function RecuperaDadosAcesso(pCaracteristicaAcesso As String, pServico_ID As Long, pFase_ID As Long, Optional pFlagInclusao As Boolean) As Boolean
        
    Dim bRetorno As Boolean
    
    Dim objDadosPropriedadeEBT  As clsGenPropriedadeEmbratel
    Dim objDadosPropriedadeTerc As clsGenPropriedadeTerceiros
    Dim lContingenciaAcesso_ID  As Long
    Dim sCaracteristicaAcesso   As String
    
    Dim objDadosTipoRadio       As clsGenTipoAcessoRadio
    Dim objDadosTipoFibra       As clsGenTipoAcessoFibra
    Dim objdadosTipoADE         As clsGenTipoAcessoADE
    Dim objDadosTipoCabo        As clsGenTipoAcessoCabo
    
    On Error GoTo ErrorHandler

    RecuperaDadosAcesso = False
    
    objDadosAcesso.InformaMensagem objMensagem
    objDadosAcesso.InformaBanco frmgenMDI_SGAPlus.objBanco
    
    'Não resgata os dados da OTS anterior quando se altera o
    'endereço de instalação. IP VPN
    If (Sigla_Acao <> "ATV" And bResgataAnterior) Then
    
        If TipoOperacao = "I" Then
        
           pServico_ID = Servico_ID_Anterior
           
           'GbCodigoFaseAcesso - ID da fase na tabela Tipo_Fase
           pFase_ID = objdadosServico.ObtemIDFase(Servico_ID_Anterior, GbCodigoFaseAcesso)
    
           If objMensagem.TipoMsg <> "" Then
                  
              GoTo ErrorHandler
        
           End If
            
        End If
        
    End If
    
    sCaracteristicaAcesso = pCaracteristicaAcesso
    
    If Not objDadosAcesso.RecuperarDados(sCaracteristicaAcesso, pServico_ID, pFase_ID, objMensagem) Then
    
       pFlagInclusao = True
       
       If objMensagem.TipoMsg <> "" Then
           
          objMensagem.ExibeMensagem
          
          GoTo Finalizar
        
       End If
       
    Else
        
        lContingenciaAcesso_ID = objDadosAcesso.iD
        CaracteristicaAcessoAnterior = objDadosAcesso.CaracteristicaAcesso
        PropriedadeAcessoAnterior = objDadosAcesso.PropridedadeAcesso
        SegurancaAcessoAnterior = objDadosAcesso.SegurancaAcesso
        ParalelismoAnterior = objDadosAcesso.Paralelismo
        
        'Obtem os dados de Equipamento Complementar
        bRetorno = objDadosAcesso.RecuperarDadosEquipamentos(lContingenciaAcesso_ID, objMensagem)
          
        If objMensagem.TipoMsg <> "" Then
            
           GoTo ErrorHandler
         
        End If
        
        'Caso o Acesso já esteja Cadastrado, Recuperar os dados Conforme a Propriedade do
        'Acesso
        pFlagInclusao = False
        
        If objDadosAcesso.PropridedadeAcesso = "E" Then   'Propriedade Acesso = Embratel
        
           Set objDadosPropriedadeEBT = New clsGenPropriedadeEmbratel
           
           objDadosPropriedadeEBT.InformaBanco frmgenMDI_SGAPlus.objBanco
           
           bRetorno = objDadosPropriedadeEBT.RecuperarDados(lContingenciaAcesso_ID, objMensagem)
            
           If objMensagem.TipoMsg <> "" Then
              
              GoTo ErrorHandler
           
           End If
           
          'Carrega a propriedade Objeto Propriedade Embratel
           If objDadosPropriedadeEBT.TipoAcesso = "R" Then
               
              Set objDadosTipoRadio = New clsGenTipoAcessoRadio
              
              objDadosTipoRadio.InformaBanco frmgenMDI_SGAPlus.objBanco
           
              bRetorno = objDadosTipoRadio.RecuperarDados(lContingenciaAcesso_ID, objMensagem)
              
              If objMensagem.TipoMsg <> "" Then
                 
                 GoTo ErrorHandler
                 
              End If

              objDadosPropriedadeEBT.objTipoAcesso = objDadosTipoRadio

           ElseIf objDadosPropriedadeEBT.TipoAcesso = "F" Then  'Fibra Otica
              
              Set objDadosTipoFibra = New clsGenTipoAcessoFibra
              
              objDadosTipoFibra.InformaBanco frmgenMDI_SGAPlus.objBanco
           
              bRetorno = objDadosTipoFibra.RecuperarDados(lContingenciaAcesso_ID, objMensagem)
              
              If objMensagem.TipoMsg <> "" Then
                 
                 GoTo ErrorHandler
                 
              End If
              
              objDadosPropriedadeEBT.objTipoAcesso = objDadosTipoFibra
           
           ElseIf objDadosPropriedadeEBT.TipoAcesso = "A" Then  'ADE
              
              Set objdadosTipoADE = New clsGenTipoAcessoADE
              
              objdadosTipoADE.InformaBanco frmgenMDI_SGAPlus.objBanco
           
              bRetorno = objdadosTipoADE.RecuperarDados(lContingenciaAcesso_ID, objMensagem)
              
              If objMensagem.TipoMsg <> "" Then
                 
                 GoTo ErrorHandler
                 
              End If
              
              objDadosPropriedadeEBT.objTipoAcesso = objdadosTipoADE
           
           ElseIf objDadosPropriedadeEBT.TipoAcesso = "C" Then  'Cabo Interno
              
              Set objDadosTipoCabo = New clsGenTipoAcessoCabo
              
              objDadosTipoCabo.InformaBanco frmgenMDI_SGAPlus.objBanco
           
              bRetorno = objDadosTipoCabo.RecuperarDados(lContingenciaAcesso_ID, objMensagem)
              
              If objMensagem.TipoMsg <> "" Then
                 
                 GoTo ErrorHandler
                 
              End If
              
              objDadosPropriedadeEBT.objTipoAcesso = objDadosTipoCabo
           
           End If
           
           TipoAcessoAnterior = objDadosPropriedadeEBT.TipoAcesso
           
           objDadosAcesso.objPropriedadeEBT = objDadosPropriedadeEBT
           
           Set objDadosPropriedadeEBT = Nothing
        Else
           
           'Propriedade Acesso = Terceiros
           Set objDadosPropriedadeTerc = New clsGenPropriedadeTerceiros
           
           objDadosPropriedadeTerc.InformaBanco frmgenMDI_SGAPlus.objBanco
           
           bRetorno = objDadosPropriedadeTerc.RecuperarDados(lContingenciaAcesso_ID, objMensagem)
            
           If objMensagem.TipoMsg <> "" Then
              
              GoTo ErrorHandler
           
           End If
           
           TipoAcessoAnterior = objDadosPropriedadeTerc.TipoAcesso
           
          'Carrega a propriedade Objeto Propriedade Terceiros
           If objDadosPropriedadeTerc.TipoAcesso = "<" Then   'Tipo de Acesso Menor 2M
               
              Set objdadosTipoAcessoMenor = New clsGenTipoAcessoMenor2M
              
              objdadosTipoAcessoMenor.InformaBanco frmgenMDI_SGAPlus.objBanco
           
              bRetorno = objdadosTipoAcessoMenor.RecuperarDados(lContingenciaAcesso_ID, objMensagem)
              
              If objMensagem.TipoMsg <> "" Then
                 
                 GoTo ErrorHandler
                 
              End If

              objDadosPropriedadeTerc.objTipoAcesso = objdadosTipoAcessoMenor

           ElseIf objDadosPropriedadeTerc.TipoAcesso = ">" Then  'Tipo de Acesso Maior 2M
              
              Set objdadosTipoAcessoMaior = New clsGenTipoAcessoMaior2M
              
              objdadosTipoAcessoMaior.InformaBanco frmgenMDI_SGAPlus.objBanco
           
              bRetorno = objdadosTipoAcessoMaior.RecuperarDados(lContingenciaAcesso_ID, objMensagem)
              
              If objMensagem.TipoMsg <> "" Then
                 
                 GoTo ErrorHandler
                 
              End If
              
              objDadosPropriedadeTerc.objTipoAcesso = objdadosTipoAcessoMaior
           
           ElseIf objDadosPropriedadeTerc.TipoAcesso = "R" Then  'Tipo de Acesso Rede Determ
              
              Set objdadosTipoAcessoRede = New clsGenTipoAcessoRedeDeterm
              
              objdadosTipoAcessoRede.InformaBanco frmgenMDI_SGAPlus.objBanco
           
              bRetorno = objdadosTipoAcessoRede.RecuperarDados(lContingenciaAcesso_ID, objMensagem)
              
              If objMensagem.TipoMsg <> "" Then
                 
                 GoTo ErrorHandler
                 
              End If
              
              objDadosPropriedadeTerc.objTipoAcesso = objdadosTipoAcessoRede
           
           End If
        
           TipoAcessoAnterior = objDadosPropriedadeTerc.TipoAcesso
           
           objDadosAcesso.objPropriedadeTerceiros = objDadosPropriedadeTerc
           
           Set objDadosPropriedadeTerc = Nothing
        
        End If
    
    End If

    RecuperaDadosAcesso = True
    
Finalizar:
    
    On Error GoTo 0
    
    Exit Function
    
ErrorHandler:
       
       If objMensagem.TipoMsg = "" Then
          
          objMensagem.TipoMsg = "E"
          objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                             vbCrLf & "Houve um erro de execução no método: " & vbCrLf & vbCrLf & "RecuperadadosAcesso do Objeto Controle " & _
                             "clsGenOC_TratarAcessoEmbratel"
       End If
       
       objMensagem.ExibeMensagem
       
       GoTo Finalizar

End Function

Public Function PopulaCamposTela(objDadosAcesso As clsGenAcesso) As Boolean

    FlagRecuperando = True
    
    'Popula os campos da pasta DADOS GERAIS
    objForm.ssTabAcesso.Tab = 0
    
    objForm.TxtEMailGic.Text = objDadosAcesso.eMailUsuario
    objForm.TxtRIT.Text = objDadosAcesso.RIT
    
    objForm.SpdDataPedidoAcesso.Row = 1
    objForm.SpdDataPedidoAcesso.Col = 1
    objForm.SpdDataPedidoAcesso.Text = objDadosAcesso.DataPedidoAcesso
    
    If objDadosAcesso.PadraoInterfaceEBT_ID > 0 Then
       
       Call SelecionaCombo(objForm.CboPadraoInterfaceEBT, _
                           CLng(objDadosAcesso.PadraoInterfaceEBT_ID))
    
    End If
    
    If Familia_Produto = "I" Then
    
        If (Produto_ID <> GbIDProdutoBLD And _
                Produto_ID <> GbIDProdutoVOIP) Then
            If Val(objDadosAcesso.Cd_Veloc) > 0 Then
           
                Call SelecionaCombo(objForm.CboVelocidadeAcesso, _
                               CLng(objDadosAcesso.Cd_Veloc))
            End If
        
        End If
    
    ElseIf Familia_Produto = "D" Then
    
        Call SelecionaComboTexto_DNI(objForm.CboVelocidadeAcesso, objDadosAcesso.Cd_Veloc)
    
    End If
    
    If objDadosAcesso.CaracteristicaAcesso = "P" Then
       
       objForm.optCaracAcessoEmbratel(0).Value = True
       objForm.optCaracAcessoTerceiro(0).Value = True
       objForm.optCaracAcessoEqpto(0).Value = True
       objForm.optCaracAcessoGen(0).Value = True
       
        If (Produto_ID = GbIDProdutoBLF Or _
                Produto_ID = GbIDProdutoBLD Or _
                Produto_ID = GbIDProdutoVOIP) Then
            'O único produto a permitir selecionar Seguranca
            'de acesso será o BLF
           
           objForm.optSegurancaAcesso(0).Enabled = True
           objForm.optSegurancaAcesso(1).Enabled = True
       
       End If
    
    Else
       
       objForm.optCaracAcessoEmbratel(1).Value = True
       objForm.optCaracAcessoTerceiro(1).Value = True
       objForm.optCaracAcessoGen(1).Value = True
       objForm.optCaracAcessoEqpto(1).Value = True
       
    End If
    
    If objDadosAcesso.SegurancaAcesso = "S" Then
       
       objForm.optSegurancaAcesso(0).Value = True
       
       objForm.CmdAcessoContingente.Visible = True

       objForm.CmdAcessoPrincipal.Visible = True
    
    Else
       
       objForm.optSegurancaAcesso(1).Value = True
    
    End If
    
    objForm.optProprietarioModem(0).Value = False
    objForm.optProprietarioModem(1).Value = False
    
    If objDadosAcesso.ProprietarioModem = "E" Then
       
       objForm.optProprietarioModem(0).Value = True
    
    Else
       
       If objDadosAcesso.ProprietarioModem = "C" Then
       
          objForm.optProprietarioModem(1).Value = True
       
       End If
       
    End If

    If objDadosAcesso.ResponsavelInfra = "E" Then
       
       objForm.OptResponsavelInfra(0).Value = True
    
    Else
        
       If objDadosAcesso.ResponsavelInfra = "C" Then
       
          objForm.OptResponsavelInfra(1).Value = True
       
       End If
       
    End If
    
    If objDadosAcesso.TipoRede = "E" Then   'Rede E1
       
       objForm.OptTipoRede(0).Value = True
    
    Else
        
       If objDadosAcesso.TipoRede = "T" Then    'Tellabs
       
          objForm.OptTipoRede(1).Value = True
    
       End If
       
    End If
    
    If objDadosAcesso.PropridedadeAcesso = "E" Then

        objForm.optProprietarioModem(0).Enabled = False
        objForm.optProprietarioModem(1).Enabled = False
        
        objForm.ssTabAcesso.TabEnabled(1) = True
        objForm.ssTabAcesso.TabEnabled(2) = False
        
        objForm.CboTipoAcessoEmbratel.Visible = True
        objForm.CboTipoAcessoTerceiro.Visible = False
        
        If Not objDadosAcesso.objPropriedadeEBT Is Nothing Then
             
           If objDadosAcesso.objPropriedadeEBT.TipoAcesso = "R" Then
                
               objForm.CboTipoAcessoEmbratel.ListIndex = 0
             
           ElseIf objDadosAcesso.objPropriedadeEBT.TipoAcesso = "F" Then
                
               objForm.CboTipoAcessoEmbratel.ListIndex = 1
             
           ElseIf objDadosAcesso.objPropriedadeEBT.TipoAcesso = "A" Then
                
               objForm.CboTipoAcessoEmbratel.ListIndex = 2
                
               objForm.optProprietarioModem(0).Enabled = True
                
               objForm.optProprietarioModem(1).Enabled = True
             
           ElseIf objDadosAcesso.objPropriedadeEBT.TipoAcesso = "C" Then
                
               objForm.CboTipoAcessoEmbratel.ListIndex = 3
                
               objForm.optProprietarioModem(0).Enabled = True
                
               objForm.optProprietarioModem(1).Enabled = True
             
           End If
             
            'Popula os campos da Pasta Propriedade Embratel
            bRetorno = PopularDadosEmbratel(objDadosAcesso.objPropriedadeEBT)
       
        End If
       
        objForm.optPropriedadeAcesso(0).Value = True
        
    Else
    
        objForm.optPropriedadeAcesso(1).Value = True
        
        objForm.CboTipoAcessoTerceiro.Visible = True
        
'        objForm.optProprietarioModem(0).Enabled = True
'        objForm.optProprietarioModem(1).Enabled = True
    
'        objForm.optProprietarioModem(0).Value = False
'        objForm.optProprietarioModem(1).Value = False
       
        objForm.CboTipoAcessoEmbratel.Visible = False
        objForm.CboTipoAcessoTerceiro.Visible = True
        
        If Not objDadosAcesso.objPropriedadeTerceiros Is Nothing Then
        
            If objDadosAcesso.objPropriedadeTerceiros.TipoAcesso = "<" Then
               
               objForm.CboTipoAcessoTerceiro.ListIndex = 0
            
            ElseIf objDadosAcesso.objPropriedadeTerceiros.TipoAcesso = ">" Then
               
               objForm.CboTipoAcessoTerceiro.ListIndex = 1
            
            ElseIf objDadosAcesso.objPropriedadeTerceiros.TipoAcesso = "R" Then
               
               objForm.CboTipoAcessoTerceiro.ListIndex = 2
            
               objForm.optProprietarioModem(0).Enabled = False
               objForm.optProprietarioModem(1).Enabled = False
            
            End If
            
       
            'Popula os campos da Pasta Propriedade Terceiros
            bRetorno = PopularDadosTerceiros(objDadosAcesso.objPropriedadeTerceiros)
        End If
        
        objForm.ssTabAcesso.TabEnabled(2) = True
        objForm.ssTabAcesso.TabEnabled(1) = False
   
    End If
    
    'Popula os campos da Pasta Equipamentos Complementares
    bRetorno = PopularEquipamentos(objDadosAcesso)

    objForm.FlControlaAlteracao = False
    
    If objForm.CmdSalvar.Enabled = False Then
    
       Call DesabilitaCamposTela
    
    End If

    If objForm.FrameExecutarCentroLocal.Visible Then
       
       If objDadosAcesso.FlagExecutaCentrolocal = "S" Then
            
          objForm.optExecCentroLocal(0).Value = True
       
       Else
          
          objForm.optExecCentroLocal(1).Value = True
       
       End If
       
    End If
    
    '--- JRCRUZ -----------------------------------------------------------**
    If objForm.FrameClassificacaoEndereco.Visible And _
            objDadosAcesso.ClassificacaoEndereco <> "" Then
       
       objForm.CboClassificacaoEndereco.ListIndex = CInt(objDadosAcesso.ClassificacaoEndereco)
       
    End If
    '--- JRCRUZ -----------------------------------------------------------**

    objForm.FlControlaAlteracao = False
    FlagRecuperando = False
    
End Function

Public Function LimpaCamposTela() As Boolean

    'Popula os campos da pasta DADOS GERAIS
    objForm.optProprietarioModem(0).Value = False
    objForm.optProprietarioModem(1).Value = False
    
    objForm.OptResponsavelInfra(0).Value = False
    objForm.OptResponsavelInfra(1).Value = False
    
    objForm.OptTipoRede(0).Value = False
    objForm.OptTipoRede(1).Value = False
    
    objForm.CboTipoAcessoEmbratel.ListIndex = -1
    objForm.TxtDesignacaoAcessoEBT.Text = ""
    objForm.TxtIdentificadorAcessoEBT.Text = ""
    objForm.TxtEstacaoEntregaEbtConf.Text = ""
    objForm.TxtEstacaoEntregaEBT.Text = ""
    objForm.TxtEstacaoCliente.Text = ""
    objForm.TxtNumeroOtsTx_EBT.Text = ""
    objForm.TxtItemOtsTxEBT.Text = ""
    objForm.SpdDataAtivacaoAcessoEBT.Text = ""
    objForm.TxtDIDFila_Radio.Text = ""
    objForm.TxtDIDArmacao_Radio.Text = ""
    objForm.TxtDIDRegua_Radio.Text = ""
    objForm.TxtDIDPosicao_Radio.Text = ""
    objForm.TxtNumeroRepeticoes = ""
    objForm.TxtDIDFila_Fibra.Text = ""
    objForm.TxtDIDArmacao_Fibra.Text = ""
    objForm.TxtDidRegua_Fibra.Text = ""
    objForm.TxtDIDPosicao_Fibra.Text = ""
    objForm.TxtNumeroCaboFibra.Text = ""
    objForm.TxtNumeroParFibra.Text = ""
    objForm.TxtPredioEBT_ADE.Text = ""
    objForm.TxtNumeroCabo_ADE.Text = ""
    objForm.TxtNumeroPar_ADE.Text = ""
    objForm.TxtPredioCabo_Cabo.Text = ""
    objForm.TxtNumeroParEBT_Cabo.Text = ""
    objForm.TxtNumeroCaboEBT_Cabo.Text = ""
    objForm.TxtNumeroParCliente_Cabo.Text = ""
    objForm.TxtNumeroCaboCliente_Cabo.Text = ""

    Call LimpaCamposTerceiros
    
    Call LimpaCamposEquipamentos
    
    objForm.optSegurancaAcesso(0).Enabled = False
    objForm.optSegurancaAcesso(1).Enabled = False
            
'    If Familia_Produto <> "D" Then
    
       objForm.optPropriedadeAcesso(0).Value = False
       objForm.optPropriedadeAcesso(1).Value = False
    
       objForm.CboTipoAcessoEmbratel.Visible = False
       objForm.CboTipoAcessoTerceiro.Visible = False
    
       objForm.ssTabAcesso.TabEnabled(1) = False
       objForm.ssTabAcesso.TabEnabled(2) = False
       objForm.ssTabAcesso.TabEnabled(3) = False
    
'    End If
        
End Function


Public Function LimpaCamposTerceiros() As Boolean

    objForm.CboTipoAcessoTerceiro.ListIndex = -1
    objForm.TxtEstacaoEntregaTercConf.Text = ""
    objForm.TxtEstacaoEntregaTerc.Text = ""
    objForm.TxtDesignacaoAcessoTerceiro.Text = ""
    objForm.SpdNumeroAcessoUrbano.Text = ""
    objForm.TxtCNLPontaA_Urbano.Text = ""
    objForm.SpdDataEntregaUrbano.Text = ""
    objForm.TxtNumeroOTSTXTerc.Text = ""
    objForm.TxtItemOTSTxTerc.Text = ""
    objForm.TxtIDentificadorAcessoUrbano.Text = ""
    objForm.SpdDataAtivacaoAcessoTerc.Text = ""
    
    objForm.optTipoTecnologia(0).Value = False
    objForm.optTipoTecnologia(1).Value = False
    objForm.optTipoTecnologia(2).Value = False
    
    objForm.CboProvedorUrbano.ListIndex = -1
    objForm.optExisteAcessoIU(0).Value = False
    objForm.optExisteAcessoIU(1).Value = False
    objForm.CboProvedorInterurbano.ListIndex = -1
    objForm.TxtIdentificadorAcessoEBTInterUrbano.Text = ""
    objForm.SpdNumeroAcessoInterurbano.Text = ""
    objForm.TxtCNLPontaA.Text = ""
    objForm.TxtCNLPontaB.Text = ""
    objForm.SpdDataEntregaInterUrbano.Text = ""
    objForm.CboInterfaceEBTMenor.ListIndex = -1
    objForm.CboInterfaceClienteMenor.ListIndex = -1
    objForm.TxtNumeroCaboAcessoMenor.Text = ""
    objForm.TxtNumeroParAcessoMenor.Text = ""
    objForm.TxtBlocoDGMenor.Text = ""
    objForm.TxtParDGMenor.Text = ""
    objForm.TxtBlocoTransdataMenor.Text = ""
    objForm.TxtParTransdataMenor.Text = ""
    objForm.TxtDIDArmacao.Text = ""
    objForm.TxtDIDFila.Text = ""
    objForm.TxtDIDPosicao.Text = ""
    objForm.TxtDIDRegua.Text = ""
    objForm.TxtDominio.Text = ""
    objForm.txtNo.Text = ""
    
    objForm.CboPorta_Rede.ListIndex = -1
    objForm.CboSlot_Rede.ListIndex = -1
    objForm.CboTSI_Rede.ListIndex = -1
    objForm.CboTSF_Rede.ListIndex = -1
    
    objForm.TxtDIDArmacao_Rede.Text = ""
    objForm.TxtDIDFila_Rede.Text = ""
    objForm.TxtDIDPosicao_Rede.Text = ""
    objForm.TxtDIDRegua_Rede.Text = ""

End Function


Public Function LimpaCamposEquipamentos() As Boolean

    objForm.TxtIDEquipamento1 = ""
    objForm.CboTipoEquipamento_1.ListIndex = -1
    objForm.TxtFabricanteEquipamento_1 = ""
    objForm.optLocalInstalacao_1(0).Value = False
    objForm.optLocalInstalacao_1(1).Value = False
    
    objForm.CboTipoEquipamento_2.ListIndex = -1
    objForm.TxtFabricanteEquipamento_2 = ""
    objForm.optLocalInstalacao_2(0).Value = False
    objForm.optLocalInstalacao_2(1).Value = False
    
    objForm.CboTipoEquipamento_3.ListIndex = -1
    objForm.TxtFabricanteEquipamento_3 = ""
    objForm.optLocalInstalacao_3(0).Value = False
    objForm.optLocalInstalacao_3(1).Value = False
    
    objForm.TxtIDEquipamento1 = ""
    objForm.TxtIDEquipamento2 = ""
    objForm.TxtIDEquipamento3 = ""

End Function

Private Function ExcluiTipoAcessoEBT(pTipoAcesso As String, pContingenciaAcesso As Long, objMensagem As clsgenMensagem) As Boolean

On Error GoTo ErrorHandler

    Dim objDadosTipoRadio      As clsGenTipoAcessoRadio
    Dim objDadosTipoFibra      As clsGenTipoAcessoFibra
    Dim objdadosTipoADE        As clsGenTipoAcessoADE
    Dim objDadosTipoCabo       As clsGenTipoAcessoCabo

    ExcluiTipoAcessoEBT = False
     
    If pTipoAcesso = "R" Then    'Tipo de Acesso Radio
        
        Set objDadosTipoRadio = New clsGenTipoAcessoRadio
        
        objDadosTipoRadio.InformaBanco frmgenMDI_SGAPlus.objBanco
        
        bRetorno = objDadosTipoRadio.ExcluirDados(pContingenciaAcesso, objMensagem)
        
        Set objDadosTipoRadio = Nothing
        
        If objMensagem.TipoMsg <> "" Then
             
           GoTo ErrorHandler
             
        End If
     
    ElseIf pTipoAcesso = "F" Then  'Tipo de Acesso Fibra Otica
    
        Set objDadosTipoFibra = New clsGenTipoAcessoFibra
        
        objDadosTipoFibra.InformaBanco frmgenMDI_SGAPlus.objBanco
        
        bRetorno = objDadosTipoFibra.ExcluirDados(pContingenciaAcesso, objMensagem)
        
        Set objDadosTipoFibra = Nothing
        
        If objMensagem.TipoMsg <> "" Then
             
           GoTo ErrorHandler
             
        End If
     
    ElseIf TipoAcessoAnterior = "A" Then 'Tipo de Acesso ADE
    
        Set objdadosTipoADE = New clsGenTipoAcessoADE
        
        objdadosTipoADE.InformaBanco frmgenMDI_SGAPlus.objBanco
        
        bRetorno = objdadosTipoADE.ExcluirDados(pContingenciaAcesso, objMensagem)
        
        Set objdadosTipoADE = Nothing
        
        If objMensagem.TipoMsg <> "" Then
             
           GoTo ErrorHandler
             
        End If
     
    ElseIf TipoAcessoAnterior = "C" Then 'Tipo de Acesso Cabo
    
        Set objDadosTipoCabo = New clsGenTipoAcessoCabo
        
        objDadosTipoCabo.InformaBanco frmgenMDI_SGAPlus.objBanco
        
        bRetorno = objDadosTipoCabo.ExcluirDados(pContingenciaAcesso, objMensagem)
        
        Set objDadosTipoCabo = Nothing
        
        If objMensagem.TipoMsg <> "" Then
             
           GoTo ErrorHandler
             
        End If
     
    End If
    
    ExcluiTipoAcessoEBT = True
    
Finalizar:
        
    On Error GoTo 0
    
    Exit Function
    
ErrorHandler:
       
    If objMensagem.TipoMsg = "" Then
       
       objMensagem.TipoMsg = "E"
       objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                   "Houve um erro de execução no Método: " & vbCrLf & vbCrLf & "ExcluiTipoAcessoEBT " & _
                   "da classe clsGenOC_TratarAcessoEmbratel"
    End If
    
    GoTo Finalizar

End Function

Private Function ExcluiTipoAcessoTerceiros(pTipoAcesso As String, pContingenciaAcesso As Long, objMensagem As clsgenMensagem) As Boolean

    Dim objDadosTipoMenor2M    As clsGenTipoAcessoMenor2M
    Dim objDadosTipoMaior2M    As clsGenTipoAcessoMaior2M
    Dim objdadosTipoRede       As clsGenTipoAcessoRedeDeterm

On Error GoTo ErrorHandler

     ExcluiTipoAcessoTerceiros = False
     
     If pTipoAcesso = "<" Then    'Tipo de Acesso Menor 2M
        
        Set objDadosTipoMenor2M = New clsGenTipoAcessoMenor2M
        
        objDadosTipoMenor2M.InformaBanco frmgenMDI_SGAPlus.objBanco
        
        bRetorno = objDadosTipoMenor2M.ExcluirDados(pContingenciaAcesso, objMensagem)
        
        Set objDadosTipoMenor2M = Nothing
        
        If objMensagem.TipoMsg <> "" Then
             
           GoTo ErrorHandler
             
        End If
     
     ElseIf pTipoAcesso = ">" Then  'Tipo de Acesso Maior 2M
    
        Set objDadosTipoMaior2M = New clsGenTipoAcessoMaior2M
        
        objDadosTipoMaior2M.InformaBanco frmgenMDI_SGAPlus.objBanco
        
        bRetorno = objDadosTipoMaior2M.ExcluirDados(pContingenciaAcesso, objMensagem)
        
        Set objDadosTipoMaior2M = Nothing
        
        If objMensagem.TipoMsg <> "" Then
             
           GoTo ErrorHandler
             
        End If
     
     ElseIf pTipoAcesso = "R" Then 'Tipo de Acesso Rede deterministica
    
        Set objdadosTipoRede = New clsGenTipoAcessoRedeDeterm
        
        objdadosTipoRede.InformaBanco frmgenMDI_SGAPlus.objBanco
        
        bRetorno = objdadosTipoRede.ExcluirDados(pContingenciaAcesso, objMensagem)
        
        Set objdadosTipoRede = Nothing
        
        If objMensagem.TipoMsg <> "" Then
             
           GoTo ErrorHandler
             
        End If
     
     End If
    
     ExcluiTipoAcessoTerceiros = True

Finalizar:
    
    On Error GoTo 0
    Exit Function
    
ErrorHandler:
       
    If objMensagem.TipoMsg = "" Then
       
       objMensagem.TipoMsg = "E"
       
       objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                   "Houve um erro de execução no Método: " & vbCrLf & vbCrLf & "ExcluiTipoAcessoTerceiros " & _
                   "da classe clsGenOC_TratarAcessoEmbratel"
                   
    End If
    
    GoTo Finalizar

End Function

Private Sub Class_Terminate()

    Set objDadosAcesso = Nothing
    Set objDadosPendencia = Nothing
    Set objDadosAcesso = Nothing
    Set objMensagem = Nothing
    Set objdadosServico = Nothing
    Set objblfOC_SelecionarOTS = Nothing
    Set objForm = Nothing

End Sub

Public Function InformaSelecionarOTS(oObjSelecionarOTS As clsgenOC_SelecionarOTS)

    Set objSelecionarOTS = oObjSelecionarOTS

End Function

Public Sub InformaOC_SelecionarOTS(ByVal objNomeClasseOC As clsgenOC_SelecionarOTS)

    Set objblfOC_SelecionarOTS = objNomeClasseOC
    
End Sub

Public Sub AtualizarListaOTS()

    objblfOC_SelecionarOTS.ExibirListaOTS

End Sub

Private Function CarregarFormAlocacaoCF(psCodigoCFCentroLocal As String) As String
    
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
        
    
    ElseIf lProduto_ID = GbIDProdutoBIV Then
    
        ' VPN
        objFormAlocacao.Caption = "IP VPN"
    
    ElseIf lProduto_ID = GbIDProdutoBLD Then
    
        ' BLD
        objFormAlocacao.Caption = "Business Link Direct"

    ElseIf lProduto_ID = GbIDProdutoVOIP Then
    
        ' BLD
        objFormAlocacao.Caption = "Internet Voice"


    End If
    
    objFormAlocacao.lblCentroFuncionalAlocado.Caption = psCodigoCFCentroLocal
    
    bRetorno = objblfOC_SelecionarOTS.ExibirListaCentroFuncional( _
                                                  objFormAlocacao.cboCentroFuncional, _
                                                  GbCodigoFaseCentroLocal, _
                                                  True)
        
    If Not bRetorno Then
        
        Exit Function
        
    End If
    
    objFormAlocacao.informa_OC Me
    
    objFormAlocacao.Show vbModal
    
    CarregarFormAlocacaoCF = CentroFuncID_Realoc
    
    Set objFormAlocacao = Nothing
    
End Function

Public Function CamposPreenchidosOK() As Boolean

    Dim objDadosFase            As clsgenFase
    Dim bRetorno                As Boolean
    Dim sDataFim                As String
    Dim sMaiorData              As String
    Dim iAux                    As Integer
    Dim bDataValida             As Boolean
    Dim bCriticaData            As Boolean
    
    Dim objDadosLocalidade      As clsGenLocalidade
    
    iAux = iAux + 1
    
    bRetorno = True
        
    bCriticaData = True
    
    If (Sigla_Acao = "ATG" Or Sigla_Acao = "ALT" Or Sigla_Acao = "ALG") And objdadosServico.Origem = "SGA" Then
       
       bCriticaData = False
       
    End If
    
    ReDim vMensagemListaMsg(iAux)
    
    vMensagemListaMsg(0) = "Os campos abaixo foram preenchidos incorretamente ou não foram preenchidos:"
    iAux = iAux + 1
    
    CamposPreenchidosOK = False
    
    If (Not VerificaDataOK(objForm.SpdDataPedidoAcesso.Text)) And _
                      Trim(objForm.SpdDataPedidoAcesso.Text) <> "" Then
        
        bRetorno = False
        
        iAux = iAux + 1
        
        ReDim Preserve vMensagemListaMsg(iAux)
        vMensagemListaMsg(iAux - 1) = "Data de Pedido do Acesso não é uma data válida!"
            
    
    End If
    
    If objForm.optSegurancaAcesso(0).Value = False And _
       objForm.optSegurancaAcesso(1).Value = False Then
        
        bRetorno = False
        
        iAux = iAux + 1
        
        ReDim Preserve vMensagemListaMsg(iAux)
        vMensagemListaMsg(iAux - 1) = "Segurança do acesso " & _
        IIf(objForm.optCaracAcessoGen(0).Value = True, " do Acesso Principal", _
                                                       " do Acesso Contingente") & _
                                      " é obrigatória."
        
    End If
    
    If objForm.optPropriedadeAcesso(0).Value = False And _
       objForm.optPropriedadeAcesso(1).Value = False Then
        
        bRetorno = False
        
        iAux = iAux + 1
        
        ReDim Preserve vMensagemListaMsg(iAux)
        vMensagemListaMsg(iAux - 1) = "Propriedade do acesso" & _
        IIf(objForm.optCaracAcessoGen(0).Value = True, " do Acesso Principal", _
                                                       " do Acesso Contingente") & _
                                      " é obrigatório."
        
    End If
    
    Set objDadosFase = New clsgenFase
    
    objDadosFase.InformaBanco frmgenMDI_SGAPlus.objBanco
    objDadosFase.InformaMensagem objMensagem
    
    sDataFim = objDadosFase.ObtemDataFimRealFaseAnterior(Servico_ID, _
                                                         Sequencia_Acao, Familia_Produto)
    
    Set objDadosFase = Nothing
    
    If objForm.optPropriedadeAcesso(0).Value = True Then 'Propriedade Embratel
        
        If objForm.CboTipoAcessoEmbratel.ListIndex = -1 Then
           
           bRetorno = False
            
           iAux = iAux + 1
            
           ReDim Preserve vMensagemListaMsg(iAux)
           vMensagemListaMsg(iAux - 1) = "Propriedade do acesso Embratel" & _
           IIf(objForm.optCaracAcessoGen(0).Value = True, " do Acesso Principal", _
                                                           " do Acesso Contingente") & _
                                          " é obrigatório."
            
        End If
    
        If Trim(objForm.TxtNumeroOtsTx_EBT.Text) <> "" And _
           Trim(objForm.TxtNumeroOtsTx_EBT.Text) <> "0" And _
           (objForm.TxtNumeroOtsTx_EBT.Enabled = True) Then
        
            If Not MascaraOTSTX_OK(Trim(objForm.TxtNumeroOtsTx_EBT.Text)) Then
                
                objForm.ssTabAcesso.Tab = 1
                
                bRetorno = False
                 
                iAux = iAux + 1
                 
                ReDim Preserve vMensagemListaMsg(iAux)
                vMensagemListaMsg(iAux - 1) = "A máscara do número de OTS TX deve ser " & _
                 "composta de (1 a 5 dígitos) seguido de "
            
                iAux = iAux + 1
                 
                ReDim Preserve vMensagemListaMsg(iAux)
                vMensagemListaMsg(iAux - 1) = "'-' mais (1 a 4) caracteres!"
            
            End If
        
        End If
    
        If (Not VerificaDataOK(objForm.SpdDataAtivacaoAcessoEBT.Text)) And _
                          Trim(objForm.SpdDataAtivacaoAcessoEBT.Text) <> "" Then
            
            bRetorno = False
            
            iAux = iAux + 1
            
            ReDim Preserve vMensagemListaMsg(iAux)
            vMensagemListaMsg(iAux - 1) = "Data de Ativação do Acesso Embratel não é uma data válida!"
            
        Else
                    
            If Trim(objForm.SpdDataAtivacaoAcessoEBT.Text) <> "" And sDataFim <> "" And _
               (Sigla_Acao <> "DES" And Sigla_Acao <> "CAN") Then
                    
                If CDate(sDataFim) > CDate(objForm.SpdDataAtivacaoAcessoEBT.Text) And _
                   bCriticaData Then
                    
                    bRetorno = False
                    
                    iAux = iAux + 1
                    
                    ReDim Preserve vMensagemListaMsg(iAux)
                    vMensagemListaMsg(iAux - 1) = "Data de Ativação do Acesso Embratel não pode ser menor que " & _
                                                  "a data de finalização da fase anterior!"
                    
                Else
                    
                    If objForm.SpdDataAtivacaoAcessoEBT.Text <> "" Then
                        
                        If CDate(objForm.SpdDataAtivacaoAcessoEBT.Text) > Date Then
                        
                            bRetorno = False
                            
                            iAux = iAux + 1
                            
                            ReDim Preserve vMensagemListaMsg(iAux)
                            vMensagemListaMsg(iAux - 1) = "Data de Ativação do Acesso Embratel não pode ser maior que " & _
                                                          "a data corrente!"
                        
                        End If
                    
                    End If
                
                End If
                
            End If
        
        End If
    Else
        
        If objForm.CboTipoAcessoTerceiro.ListIndex = -1 Then
        
           bRetorno = False
             
           iAux = iAux + 1
            
           ReDim Preserve vMensagemListaMsg(iAux)
           vMensagemListaMsg(iAux - 1) = "Tipo do acesso Terceiros é obrigatório!"
        
        End If
    
        If Trim(objForm.TxtNumeroOTSTXTerc) <> "" And _
           Trim(objForm.TxtNumeroOTSTXTerc) <> "0" Then
        
            If Not MascaraOTSTX_OK(Trim(objForm.TxtNumeroOTSTXTerc.Text)) Then
            
                bRetorno = False
                 
                iAux = iAux + 1
                 
                ReDim Preserve vMensagemListaMsg(iAux)
                vMensagemListaMsg(iAux - 1) = "A máscara do número de OTS TX deve ser " & _
                 "composta de (1 a 5 dígitos) seguido de "
            
                iAux = iAux + 1
                 
                ReDim Preserve vMensagemListaMsg(iAux)
                vMensagemListaMsg(iAux - 1) = "'-' mais (1 a 4) caracteres!"
            End If
        
        End If
            
        If Trim(objForm.TxtCNLPontaA_Urbano.Text) <> "" Then
            
            Set objDadosLocalidade = New clsGenLocalidade
            
            objDadosLocalidade.InformaBanco frmgenMDI_SGAPlus.objBanco
            objDadosLocalidade.InformaMensagem objMensagem
            
            If Not objDadosLocalidade.VerificaExistenciaCNL( _
                                      TrataPlick(objForm.TxtCNLPontaA_Urbano.Text)) Then
            
               bRetorno = False
                 
               iAux = iAux + 1
                 
               ReDim Preserve vMensagemListaMsg(iAux)
               vMensagemListaMsg(iAux - 1) = "CNL da ponta A do provedor urbano informado " & _
                                             "não foi encontrado na tabela de"
             
               iAux = iAux + 1
                 
               ReDim Preserve vMensagemListaMsg(iAux)
               vMensagemListaMsg(iAux - 1) = " Localidades!"
               
            End If
        
            Set objDadosLocalidade = Nothing
        
        End If
        
        If objForm.SpdDataEntregaUrbano.Enabled = True Then
            
            If (Not VerificaDataOK(objForm.SpdDataEntregaUrbano.Text)) And _
                              Trim(objForm.SpdDataEntregaUrbano.Text) <> "" Then
                
                bRetorno = False
                 
                iAux = iAux + 1
                 
                ReDim Preserve vMensagemListaMsg(iAux)
                vMensagemListaMsg(iAux - 1) = "Data de Entrega do Acesso Urbano " & _
                                              "não é uma data válida!"
            Else
            
                If objForm.SpdDataEntregaUrbano.Text <> "" Then
                    
                   If CDate(objForm.SpdDataEntregaUrbano.Text) > Date Then
                        
                        bRetorno = False
                         
                        iAux = iAux + 1
                         
                        ReDim Preserve vMensagemListaMsg(iAux)
                        vMensagemListaMsg(iAux - 1) = "Data de Entrega do Acesso Urbano não pode ser maior que " & _
                                                      "a data corrente!"
                        
                   End If
                
                End If
                
            End If
        
        End If
        
        If Trim(objForm.TxtCNLPontaA.Text) <> "" Then
            
            Set objDadosLocalidade = New clsGenLocalidade
            
            objDadosLocalidade.InformaBanco frmgenMDI_SGAPlus.objBanco
            objDadosLocalidade.InformaMensagem objMensagem
            
            If Not objDadosLocalidade.VerificaExistenciaCNL( _
                                      TrataPlick(objForm.TxtCNLPontaA.Text)) Then
            
               bRetorno = False
                 
               iAux = iAux + 1
                 
               ReDim Preserve vMensagemListaMsg(iAux)
               vMensagemListaMsg(iAux - 1) = "CNL da ponta A do provedor interurbano informado " & _
                                             "não foi encontrado na tabela de"
             
               iAux = iAux + 1
                 
               ReDim Preserve vMensagemListaMsg(iAux)
               vMensagemListaMsg(iAux - 1) = "'Localidades!"
               
            End If
            
            Set objDadosLocalidade = Nothing
        
        End If
        
        If Trim(objForm.TxtCNLPontaB.Text) <> "" Then
            
            Set objDadosLocalidade = New clsGenLocalidade
            
            objDadosLocalidade.InformaBanco frmgenMDI_SGAPlus.objBanco
            objDadosLocalidade.InformaMensagem objMensagem
                
            If Not objDadosLocalidade.VerificaExistenciaCNL( _
                                      TrataPlick(objForm.TxtCNLPontaB.Text)) Then
            
               bRetorno = False
                 
               iAux = iAux + 1
                 
               ReDim Preserve vMensagemListaMsg(iAux)
               vMensagemListaMsg(iAux - 1) = "CNL da ponta B do provedor interurbano informado " & _
                                             "não foi encontrado na tabela de"
             
               iAux = iAux + 1
                 
               ReDim Preserve vMensagemListaMsg(iAux)
               vMensagemListaMsg(iAux - 1) = "'Localidades!"
               
            End If
            
            Set objDadosLocalidade = Nothing
        
        End If
        
        If objForm.SpdDataEntregaInterUrbano.Enabled = True Then
            
            If (Not VerificaDataOK(objForm.SpdDataEntregaInterUrbano.Text)) And _
                              Trim(objForm.SpdDataEntregaInterUrbano.Text) <> "" Then
            
                bRetorno = False
                 
                iAux = iAux + 1
                 
                ReDim Preserve vMensagemListaMsg(iAux)
                vMensagemListaMsg(iAux - 1) = "Data de Entrega do Acesso InterUrbano não é uma data válida!"
                
            Else
                If objForm.SpdDataEntregaInterUrbano.Text <> "" Then
                    
                    If CDate(objForm.SpdDataEntregaInterUrbano.Text) > Date Then
                    
                        bRetorno = False
                          
                         iAux = iAux + 1
                          
                         ReDim Preserve vMensagemListaMsg(iAux)
                         vMensagemListaMsg(iAux - 1) = "Data de Entrega do Acesso InterUrbano não pode ser maior que " & _
                                                       "a data corrente!"
                          
                    End If
                
                End If
                
            End If
        
        End If
        
        If objForm.SpdDataAtivacaoAcessoTerc.Enabled Then
            
            If (Not VerificaDataOK(objForm.SpdDataAtivacaoAcessoTerc.Text)) And _
                              Trim(objForm.SpdDataAtivacaoAcessoTerc.Text) <> "" Then
            
                 bRetorno = False
                  
                 iAux = iAux + 1
                  
                 ReDim Preserve vMensagemListaMsg(iAux)
                 vMensagemListaMsg(iAux - 1) = "Data de Ativação do Acesso Terceiros não é uma data válida!"
                
                 bDataValida = False
            Else
            
                bDataValida = True

            End If
        
        End If
        
        If objForm.SpdDataEntregaUrbano.Text <> "" Then
           
           sMaiorData = objForm.SpdDataEntregaUrbano.Text
        
        End If
        
        If sMaiorData <> "" And objForm.SpdDataEntregaInterUrbano.Text <> "" Then
            
            If CDate(objForm.SpdDataEntregaInterUrbano.Text) > CDate(sMaiorData) Then
               
               sMaiorData = objForm.SpdDataEntregaInterUrbano.Text
            
            End If
        
        End If
        
        If objForm.SpdDataAtivacaoAcessoTerc.Text <> "" And sMaiorData <> "" And _
           bDataValida Then
            
            If CDate(objForm.SpdDataAtivacaoAcessoTerc.Text) < CDate(sMaiorData) Then
                
                bRetorno = False
                  
                iAux = iAux + 1
                  
                ReDim Preserve vMensagemListaMsg(iAux)
                vMensagemListaMsg(iAux - 1) = "Data de Ativação do Acesso Terceiro não pode ser menor " & _
                      " que a maior data de Entrega entre o Acesso Urbano e o Interurbano!"
            End If
        
        End If
        
        If Trim(objForm.SpdDataAtivacaoAcessoTerc.Text) <> "" And sDataFim <> "" And _
           (Sigla_Acao <> "DES" And Sigla_Acao <> "CAN") And bDataValida Then
        
            If CDate(sDataFim) > CDate(objForm.SpdDataAtivacaoAcessoTerc.Text) And _
               bCriticaData Then
                
                iAux = iAux + 1
                  
                ReDim Preserve vMensagemListaMsg(iAux)
                vMensagemListaMsg(iAux - 1) = "Data de Ativação do Acesso terceiro não pode ser menor que " & _
                                        "a data de finalização da fase anterior!"
            
            End If
            
        End If
    
    End If
    
    '-- JRCRUZ --------------------------------------------**
    If objForm.FrameClassificacaoEndereco.Visible = True Then
        If objForm.CboClassificacaoEndereco.ListIndex = -1 Then
        
           bRetorno = False
             
           iAux = iAux + 1
            
           ReDim Preserve vMensagemListaMsg(iAux)
           vMensagemListaMsg(iAux - 1) = "Classificação de Endereço é obrigatório!"
        
        End If
    End If
    '-- JRCRUZ --------------------------------------------**
    
    If Not bRetorno Then
    
        Call ExibeListaMensagem
    
    Else
    
        CamposPreenchidosOK = True
        
    End If
    
End Function

'Private Function CarregaCabecalhoAcesso_DNI(Usr As UsrDNICabecalhoAcesso, ByVal pServico_ID As Long, ByVal pFase_ID As Long, objMensagem As clsgenMensagem) As Boolean
'
'    Dim objDadosCabecalho As New clsDNI_DadosCabecalho
'
'    'Instancia objeto Cabecalho para obter dados do cabecalho padrao
'    objDadosCabecalho.InformaBanco frmgenMDI_SGAPlus.ObjBanco
'
'    'Seta o objeto mensagem da classe de dados
'    objDadosCabecalho.InformaMensagem objMensagem
'
'    If Not objDadosCabecalho.RecuperarDadosCabecalho_DNI(pServico_ID, _
'                                                         pFase_ID, GbCodigoFaseAcesso) Then
'
'        If objMensagem.TipoMsg <> "" Then
'
'           objMensagem.ExibeMensagem
'
'           Exit Function
'
'        End If
'    End If
'
'    Usr.OTS = objDadosCabecalho.CodigoOTS
'
'    Usr.ContaCorrente = objDadosCabecalho.ContaCorrente
'    Usr.NomeCliente = objDadosCabecalho.NomeCliente
'    Usr.Velocidade = objDadosCabecalho.Velocidade
'    Usr.Interface_Cliente = objDadosCabecalho.Interface_Cliente
'    Usr.CNL_Instalacao = objDadosCabecalho.CNL_Instalacao
'    Usr.Estacao = objDadosCabecalho.Estacao
'    Usr.DataFinalizacao = objDadosCabecalho.DataFinalizacao
'    Usr.AnoOE = objDadosCabecalho.AnoOE
'    Usr.NumeroOE = objDadosCabecalho.NumeroOE
'    Usr.ItemOE = objDadosCabecalho.ItemOE
'    Usr.Contrato = objDadosCabecalho.Contrato
'    Usr.OtsAnterior = objDadosCabecalho.CodigoOTSAnterior
'    Usr.DesignacaoCCTOAnterior = objDadosCabecalho.Designacao_CCTOAnterior
'
'    Usr.CNLPontaA = objDadosCabecalho.CNLPontaA
'    Usr.CNLPontaB = objDadosCabecalho.CNLPontaB
'    Usr.CodigoNP = objDadosCabecalho.CodigoNP
'
'    If objDadosCabecalho.Produto_ID = GbIDProdutoDiginet Then
'
'        Usr.SequencialDesig = Format(objDadosCabecalho.SequencialDesig, "0000")
'
'    ElseIf objDadosCabecalho.Produto_ID = GbIDProdutoEmbratelIPL Then
'
'        Usr.SequencialDesig = Format(objDadosCabecalho.SequencialDesig, "0000000")
'
'    End If
'
'
'    If Trim(objDadosCabecalho.Designacao_CCTO) <> "0" Then
'
'       Usr.DesignacaoCCTO = objDadosCabecalho.Designacao_CCTO
'
'    End If
'
'    Usr.CarregarDados
'
'    Set objDadosCabecalho = Nothing
'
'End Function


Private Function ExibeListaMensagem() As Boolean

    Dim iAux     As Integer
    
    Load FrmGenListaMensagem
    
    FrmGenListaMensagem.Left = 1335
    FrmGenListaMensagem.Top = 1785
    
    For iAux = 0 To UBound(vMensagemListaMsg, 1)
    
        FrmGenListaMensagem.LstMensagem.AddItem vMensagemListaMsg(iAux)
    
    Next iAux
    
    ReDim vMensagemListaMsg(1)
    FrmGenListaMensagem.Show vbModal
    
End Function

Public Function CarregarFormCFsDaFase(ByVal plIdFase As Long) As Boolean
    
    'IP VPN
    Dim objDadosProduto         As New clsGenProduto
    Dim objAlocCFs              As clsGen_DetalheAlocarCF
    Dim objDadosFaseVariosCfs   As clsgenFase
    Dim objDadosVariosCFs       As clsBLF_AlocacaoCF
    Dim objOcAlocarCF           As clsInterOC_AlocarCF

    On Error GoTo ErrorHandler
    
    CarregarFormCFsDaFase = False
    
    'Formulário
    '''Set objFormVariasFases = New frmgenAlocaCFVariasFases
    
    'Instância a classe para fornecer ao formulário
    Set objOcAlocarCF = New clsInterOC_AlocarCF

    'Dados váris CFs alocados por Fase
    Set objDadosFaseVariosCfs = New clsgenFase
    objDadosFaseVariosCfs.InformaBanco frmgenMDI_SGAPlus.objBanco
    objDadosFaseVariosCfs.InformaMensagem objMensagem
    
    'Dados váris CFs Alocados
    Set objDadosVariosCFs = New clsBLF_AlocacaoCF
    objDadosVariosCFs.InformaBanco frmgenMDI_SGAPlus.objBanco
    objDadosVariosCFs.InformaMensagem objMensagem
    
    'Obtém dados do Produto
    objDadosProduto.InformaBanco frmgenMDI_SGAPlus.objBanco
    objDadosProduto.InformaMensagem objMensagem
        
    bRetorno = objDadosProduto.RecuperarDados(Produto_ID)
    
    If Not bRetorno Then
        GoTo Finalizar
    End If
       
    'Instancia o formulário na classe de controle que fica responsável
    'pela atualização da coleção
    bRetorno = objOcAlocarCF.InstanciaoFormVariasFases
        
    If Not bRetorno Then
        GoTo Finalizar
    End If
       
    'Atualiza o caption da Tela Vários CFs
    'objFormVariasFases.Caption = objDadosProduto.Descricao
    objOcAlocarCF.objFormVariasFases.Caption = objDadosProduto.Descricao
    
    'objFormVariasFases.FramTituloFase.Caption = "Alocação Centros Funcionais da Fase Tratar Centro Local"
    objOcAlocarCF.objFormVariasFases.FramTituloFase.Caption = "Alocação Centros Funcionais da Fase Tratar Centro Local"
        
    'Carrega os CFs alocados para o Fase
    If Not objDadosFaseVariosCfs.MontaCFsDaFase(plIdFase) Then
        Exit Function
    End If
        
    'Retorna os CFs da Fase carregando o combo da tela
    bRetorno = objblfOC_SelecionarOTS.ExibirListaCentroFuncional( _
           objOcAlocarCF.objFormVariasFases.cboCentroFunc, _
            GbCodigoFaseCentroLocal, _
            True)
            
    If Not bRetorno Then
    
        GoTo Finalizar
        
    End If

    'Carrega na tela os CFs já selecionados
    For Each objAlocCFs In objDadosFaseVariosCfs.ColecaoCFsDaFase
        
        objOcAlocarCF.objFormVariasFases.lstCfs.AddItem objAlocCFs.Codigo_Centro_Funcional
        objOcAlocarCF.objFormVariasFases.lstCfs.ItemData(objOcAlocarCF.objFormVariasFases.lstCfs.NewIndex) = objAlocCFs.Centro_Funcional_ID
    
    Next objAlocCFs
    
    With objOcAlocarCF.objFormVariasFases

        .cboCentroFunc.Enabled = True
        .lstCfs.Enabled = True
        .CmdAdd.Enabled = True
        .CmdDel.Enabled = True
        .CmdSalvar.Enabled = True
        .CmdLimpar.Enabled = True
    
    End With
    
    'Informa a classe de Dados Fase Vários CFs
    Call objOcAlocarCF.objFormVariasFases.InformaClasseVariosCFsGEN(objDadosFaseVariosCfs)
        
    'Informa a classe de Controle ALocar CF
    Call objOcAlocarCF.objFormVariasFases.InformaClasseAlocarCFOC(objOcAlocarCF)
    
    'Passa ao controlador do formulário vários CFs a classe com os CFs alocados
    Set objOcAlocarCF.objDadosFaseVariosCfs = objDadosFaseVariosCfs
    
    
    'Carrega o formulário Vários CFs
    objOcAlocarCF.objFormVariasFases.Show 1
    
    'Atualiza e grava a coleção de CF's, relacionada aquela fase da OTS
    Set objAlocCFs = New clsGen_DetalheAlocarCF
    
    If objOcAlocarCF.objDadosFaseVariosCfs.ColecaoCFsDaFase.Count > 0 Then
    
        If objDadosVariosCFs.ExcluiDadosVariasFases(plIdFase) Then
            
            For Each objAlocCFs In objOcAlocarCF.objDadosFaseVariosCfs.ColecaoCFsDaFase
                
                Set objAlocacaoFase = New clsGen_DetalheAlocarCF
                    
                If Not objDadosVariosCFs.GravaDadosVariasFases(plIdFase, _
                        objAlocCFs.Centro_Funcional_ID) Then
                        
                    GoTo ErrorHandler
                    
                End If
            
            Next objAlocCFs
        
        End If
    
    Else
        objMensagem.TipoMsg = "A"
        objMensagem.Descricao = "Ao menos um Centro Funcional deve ser selecionado."
        objMensagem.ExibeMensagem
        CarregarFormCFsDaFase = False
        
        GoTo Finalizar
        
    End If
    
    'Executa método para popular os campos do Controle Cabecalho
    CarregarFormCFsDaFase = True

    GoTo Finalizar
    
Finalizar:
    
    Set objDadosProduto = Nothing
    Set objAlocCFs = Nothing
    
    
    Set objDadosProduto = Nothing
    Set objDadosFaseVariosCfs = Nothing
    Set objDadosVariosCFs = Nothing
    
    Exit Function
    
ErrorHandler:

       objMensagem.TipoMsg = "E"
       
       objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                           vbCrLf & "Houve um erro de execução no método: " & vbCrLf & vbCrLf & "CarregarForm do Objeto Controle " & _
                          "clsGenOC_TratarAcessoEmbratel"
       
       objMensagem.ExibeMensagem
       
       GoTo Finalizar
End Function

Public Function RealocarCF() As Boolean

    Dim objDadosFase           As New clsgenFase

    objDadosFase.InformaBanco frmgenMDI_SGAPlus.objBanco
    objDadosFase.InformaMensagem objMensagem

    If Not objForm.FlControlaAlteracao And _
       Not bFlagControlaAlteracao Then
    
    'Verificar se a fase já foi iniciada, em caso afirmativo, exibe mensagem de alerta
    '
        If objDadosFase.ObtemUltimaFaseRealizada(Servico_ID) <> Fase_ID Then
        
            If CarregarFormRealocarCF(CentroFuncionalID, GbCodigoFaseAcesso, _
                      objblfOC_SelecionarOTS, Me) <> 0 Then
        
                Unload objForm
        
            End If
        
        Else
            objMensagem.TipoMsg = "A"
            objMensagem.Descricao = "Não é possível a realocação do Centro Funcional! Esta fase encontra-se iniciada/concluída!"
            objMensagem.ExibeMensagem
            
        End If
    
    Else
        If objDadosFase.ObtemUltimaFaseRealizada(Servico_ID) <> Fase_ID Then
        
            objMensagem.TipoMsg = "Q"
            objMensagem.Descricao = "Existem dados da tela que serão perdido! Continua assim mesmo? "
            
            If objMensagem.ExibeMensagem = "S" Then
            
                If CarregarFormRealocarCF(CentroFuncionalID, GbCodigoFaseAcesso, _
                          objblfOC_SelecionarOTS, Me) <> 0 Then
            
                    Unload objForm
            
                End If
            
            End If
            
        End If
        
    End If

End Function

Public Function SolicitarDadosSwitch() As Boolean

    Dim objDadosEqptoNGN As VIP_Equipamento_NGN
    
    Set objDadosEqptoNGN = New VIP_Equipamento_NGN
    
    objDadosEqptoNGN.InformaBanco frmgenMDI_SGAPlus.objBanco
    objDadosEqptoNGN.InformaMensagem objMensagem
    
    bRetorno = objDadosEqptoNGN.SolicitarDadosSwitch(Fase_ID)
        
    If objMensagem.TipoMsg <> "" Then
    
        objMensagem.ExibeMensagem
        Exit Function
        
    End If
    
    If bRetorno Then
    
        objForm.CmdObterEnderecoIP.Caption = "Aguardando dados de Switch"
        
    End If
    
End Function

Public Function CarregarComboEqpto() As Boolean

    Dim objDadosEqptoNGN As VIP_Equipamento_NGN
    
    Set objDadosEqptoNGN = New VIP_Equipamento_NGN
    
    objDadosEqptoNGN.InformaBanco frmgenMDI_SGAPlus.objBanco
    objDadosEqptoNGN.InformaMensagem objMensagem
    
    Set objOC_DadosSwitch = New clsVIPoc_CadastrarEqptoNGN
    
    bRetorno = objOC_DadosSwitch.CarregarDadosSwitch(objDadosEqptoNGN, Fase_ID, _
                                 objForm.UsrVIPDadosSwitch)
    
End Function

'Exibe tela com motivo de devolução da OTS pelo ConfServ na fase NGN Acesso
Public Sub ExibeMotivoDevolucao()

    Dim lIDFase As Long
    Dim objFormObs As frmVIPObsDEvolucao
    Dim objDadosCtrlCorrecao As clsGenDadosControleCorrecao
    
    lIDFase = objdadosServico.ObtemIDFase(Servico_ID, GbCodigoFaseConfiguracaoNGN)
    
    Set objDadosCtrlCorrecao = New clsGenDadosControleCorrecao
    
    objDadosCtrlCorrecao.InformaBanco frmgenMDI_SGAPlus.objBanco
    objDadosCtrlCorrecao.InformaMensagem objMensagem
    
    If objDadosCtrlCorrecao.RecuperarDados(lIDFase, "D") Then
        
        Set objFormObs = New frmVIPObsDEvolucao
        
        objFormObs.TxtObsCorrecao.Text = objDadosCtrlCorrecao.Justificativa
        objFormObs.LblDataSolicitacao.Caption = objDadosCtrlCorrecao.DataSolicitacao
        objFormObs.TxtObsCorrecao.Locked = True
        objFormObs.CmdGravar.Enabled = False
        objFormObs.CmdGravar.Visible = False
        
        Set objDadosUsuario = New clsgenUsuario
        
        objDadosUsuario.InformaBanco frmgenMDI_SGAPlus.objBanco
        objDadosUsuario.InformaMensagem objMensagem
        
        bRetorno = objDadosUsuario.RecuperarDados(objDadosCtrlCorrecao.UsuarioIDSolic)
        
        If bRetorno Then
            
            objFormObs.LblUsuarioSolicitacao.Caption = objDadosUsuario.Login
             
        End If
        
        Set objDadosUsuario = Nothing
        
        objFormObs.Show vbModal
        
    End If

End Sub

Public Sub CarregarFormEqptoNGN()

    Dim ObjEqptoNGN As New VIP_Equipamento_NGN
    Dim objOc_EqptoNGN As clsVIPoc_CadastrarEqptoNGN
    
    Set objOc_EqptoNGN = New clsVIPoc_CadastrarEqptoNGN
    
    objOc_EqptoNGN.Fase_ID = Fase_ID
    
    If objForm.CmdSalvar.Enabled = True Then
       
       objOc_EqptoNGN.TipoOperacao = "I"
    
    Else
       
       objOc_EqptoNGN.TipoOperacao = "C"
    
    End If

    Set ObjEqptoNGN = New VIP_Equipamento_NGN
    
    ' Informar o nome do banco desejado
    ObjEqptoNGN.InformaBanco frmgenMDI_SGAPlus.objBanco
    ObjEqptoNGN.InformaMensagem objMensagem
        
    If ObjEqptoNGN.VerificaEqptoNaOTS(Fase_ID) Then
            
        objOc_EqptoNGN.TipoOperacao = "C"
                
    End If

    'Verifica se já existe eqpto Cadastrado
    objOc_EqptoNGN.Contrato_ID = ContratoID
    objOc_EqptoNGN.Endereco_ID = EnderecoID
    
    Call objOc_EqptoNGN.CarregarFormEquipamento
    
    Call CarregarComboEqpto

End Sub
