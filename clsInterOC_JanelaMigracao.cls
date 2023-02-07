VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInterOC_JanelaMigracao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public objForm                  As frmPortJanelaMigracao

Public objDadosPendencia        As New clsblf_Pendencia

Public objOCSelecionarOTS       As clsgenOC_SelecionarOTS

Dim objMensagem                 As New clsgenMensagem
Dim sStatus_Fase                As String
Dim lServico_ID                 As Long
Dim lFase_Id                    As Long
Dim lProduto_ID                 As Long
Dim sDataFinalizacao            As String
Dim lSequencia_Acao             As Long
Dim sTextoAnotacao              As String
Dim sPermiteExecucao            As String
Dim lCentroFuncionalID          As Long
Dim sStatusTelefone              As String
Dim sAcao                         As String
Dim lGrupo_ID                   As Long

Dim blFlagControlaAlteracao   As String


'BIP VPN 16/06/2003
Private bPermiteAtualizar       As Boolean

Public Property Get bFlagControlaAlteracao() As String

    bFlagControlaAlteracao = blFlagControlaAlteracao

End Property

Public Property Let bFlagControlaAlteracao(ByVal sNewValue As String)

    blFlagControlaAlteracao = sNewValue

End Property

Public Property Get Acao() As String

    Acao = sAcao

End Property

Public Property Let Acao(ByVal sNewValue As String)

    sAcao = sNewValue

End Property

Public Property Get StatusTelefone() As String

    StatusTelefone = sStatusTelefone

End Property

Public Property Let StatusTelefone(ByVal sNewValue As String)

    sStatusTelefone = sNewValue

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

Public Property Get CentroFuncionalID() As Long

    CentroFuncionalID = lCentroFuncionalID

End Property

Public Property Let CentroFuncionalID(ByVal lNewValue As Long)

    lCentroFuncionalID = lNewValue

End Property

Public Property Get DescricaoCF() As String

DescricaoCF = sDescricaoCF

End Property

Public Property Let DescricaoCF(ByVal sNewValue As String)

sDescricaoCF = sNewValue

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
    Dim paPermissoes()         As Variant
    Dim sSituacaoOTS           As String
    Dim cColecao               As Collection
    Dim flUsuarioDoCF          As Boolean
    'Dim sDataFinalizacao       As String
    
    Dim objAtenPortab               As New VIP_Aten_Portabilidade
    Dim objSGAP_PUB_BW        As New clsGenPubBW
    
    '================================================================
    'CALBAT - Acrescido para pertir alteração  de HIQ
    '================================================================
    Dim objAtenTecVoz        As New VIP_Aten_Tec_Voz
    '================================================================
    '================================================================
    
    Dim objAtenItemServico    As VIP_Aten_Item_Servico
    Dim objAtenContrato          As VIP_Aten_Contrato
    
    Dim sHrAtual As String
    Dim sHrJanelaPort As String
    
    
    
    
    On Error GoTo ErrorHandler
    
    CarregarForm = False
    
    'Este campo após crítica pode ser desabilitado. IP VPN
    bPermiteAtualizar = True
    
        '#################
        'adeil RN-SGAD-14971ZKH 28/10/2003
        'se o form passado por parâmetro já estiver aberto, fecha-lo.
        'inicio
'        If Verifica_Janela_Aberta("frmPortJanelaMigracao", objForm.Caption & psTitulo) Then
'            GoTo Finalizar
'        End If
    
    Set objAtenItemServico = New VIP_Aten_Item_Servico
    
    objAtenItemServico.InformaBanco frmgenMDI_SGAPlus.objBanco
    objAtenItemServico.InformaMensagem objMensagem
    
    
    Set objSGAP_PUB_BW = New clsGenPubBW
    
    objSGAP_PUB_BW.InformaBanco frmgenMDI_SGAPlus.objBanco
    objSGAP_PUB_BW.InformaMensagem objMensagem
    
   
    bRetorno = objAtenItemServico.RecuperarDados(lServico_ID)

    If (objMensagem.TipoMsg <> "") Then

       GoTo ErrorHandler

    End If
    
    
    '================================================================
    'CALBAT - Acrescido para pertir alteração  de HIQ
    '================================================================
    objAtenTecVoz.InformaBanco frmgenMDI_SGAPlus.objBanco
    objAtenTecVoz.InformaMensagem objMensagem
    
    bRetorno = objAtenTecVoz.RecuperarDados(objAtenItemServico.iD)

    If (objMensagem.TipoMsg <> "") Then

       GoTo ErrorHandler

    End If
    
    '================================================================
    '================================================================
    
    Set objAtenContrato = New VIP_Aten_Contrato
    
    objAtenContrato.InformaBanco frmgenMDI_SGAPlus.objBanco
    objAtenContrato.InformaMensagem objMensagem
    
    bRetorno = objAtenContrato.RecuperarDados(objAtenItemServico.ContratoID)

    If (objMensagem.TipoMsg <> "") Then

       GoTo ErrorHandler

    End If
    
    
    Set AtenPortab = New VIP_Aten_Portabilidade
    
    objAtenPortab.InformaBanco frmgenMDI_SGAPlus.objBanco
    objAtenPortab.InformaMensagem objMensagem
    
    'Obtem dados servico e verifica se a OTS selecionada ainda encontra-se em andamento
    bRetorno = objAtenPortab.RecuperarDados(Servico_ID)
    
    If objMensagem.TipoMsg <> "" Then
       
       objMensagem.ExibeMensagem
       
       Set objAtenPortab = Nothing
       
       GoTo Finalizar
       
    End If
    
             
    'Verifica se a Fase ainda continua Em Andamento
'    objDadosFase.InformaBanco frmgenMDI_SGAPlus.objBanco
'    objDadosFase.InformaMensagem objMensagem
'
    Set objDadosFase = New clsgenFase
    
    objDadosFase.InformaBanco frmgenMDI_SGAPlus.objBanco
    objDadosFase.InformaMensagem objMensagem
    

'    bretorno objDadosFase.RecuperarSituacaoFase(Fase_ID) And DataFinalizacao = "" Then
'
'       objMensagem.ExibeMensagem
'
'       GoTo Finalizar
'    Else
'
'       objMensagem.TipoMsg = ""
'       objMensagem.Descricao = ""
'
'    End If
             
    objAtenPortab.InformaBanco frmgenMDI_SGAPlus.objBanco
    objAtenPortab.InformaMensagem objMensagem
             
    
    'Checa se o produto pode alocar vários CFs para a fase Tratar Centro Local.
    'É instanciado e recebe os objetos Banco e Mensagem apenas neste método
    Set objDadosDuracaoFase = New clsgenDuracaoFase
    objDadosDuracaoFase.InformaBanco frmgenMDI_SGAPlus.objBanco
    objDadosDuracaoFase.InformaMensagem objMensagem
    
    
    'Obtem a Sequencia_Acao da última fase realizada
'    numUltimaFaseRealizada = objDadosFase.ObtemSequenciaUltimaFaseRealizada(Servico_ID, GbCodigoFasePortabilidadeJanelaMigracao)
'
'    If objMensagem.TipoMsg <> "" Then
'
'       objMensagem.ExibeMensagem
'
'       Set objdadosServico = Nothing
'
'       GoTo Finalizar
'
'    End If
    
    Set objAutenticacao = New clsgenAutenticacao
    
    'Verifica as permissoes do usuario
    objAutenticacao.IDUsuario = frmgenMDI_SGAPlus.objAutenticacao.IDUsuario
    objAutenticacao.Login = frmgenMDI_SGAPlus.objAutenticacao.Login
    objAutenticacao.InformaMensagem objMensagem
    
    Call objAutenticacao.InformaBanco(frmgenMDI_SGAPlus.objBanco)
    
   '======================================================================================
   'Instancia o formulario, obtem os dados da fase e carrega os campos da tela
   '======================================================================================
    Set objForm = New frmPortJanelaMigracao
    
    
    
    '=================================================================================
    'Só permite a digitação da inforamação, caso o produto seja RedeVip
    '=================================================================================
    '=================================================================================
    
    
    If lProduto_ID = GbIDProdutoAAC Then
    
            objForm.spdDataAtivacaoTecnica.Enabled = True
            objForm.spdHoraAtivacaoTecnica.Enabled = True
            
            objForm.spdDataAtivacaoTecnica.Value = Format(Date, "dd/mm/yyyy")
            objForm.spdHoraAtivacaoTecnica.Value = Format(Now, "hh:mm")
            
            'GbCodigoFaseCentroLocal - Codigo da fase Tratar Centro Local cadastrado na tabela Tipo fase
            Call objAutenticacao.RetornarPermissoesFase(objAutenticacao.IDUsuario, _
                                                        paPermissoes, _
                                                        GbCodigoFaseCentroLocal, _
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
    
    Else
    
            Call DesabilitaObjetosTela
    
    End If
    
    '=================================================================================
    '=================================================================================


    objDadosProduto.InformaBanco frmgenMDI_SGAPlus.objBanco
    objDadosProduto.InformaMensagem objMensagem
    
    bRetorno = objDadosProduto.RecuperarDados(Produto_ID)
    
    
    Set objPortVipCabecalho = New VIP_Cabecalho_Port
    
    
    objPortVipCabecalho.NomeCliente = objAtenContrato.RazaoSocialCliente
    objPortVipCabecalho.SubConta = objAtenContrato.ContaCorrente
    
    'Se o produto for um NetPhone, mostra o código NET
    'AGRC - SO-TIEMP-49824BGM - 15/11/2017
    'If lProduto_ID = GbIDProdutoNetPhone Then
    If lProduto_ID = GbIDProdutoNetPhone Or lProduto_ID = GbIDProdutoClaroFone Then
        objPortVipCabecalho.Contrato = Format(objAtenContrato.ClienteID, "000000000000")
        objForm.UsrPortabilidadeCabecalho.MostraCodigoNET (True)
    Else
        objPortVipCabecalho.Contrato = ""
        objForm.UsrPortabilidadeCabecalho.MostraCodigoNET (False)
    End If
    
    objPortVipCabecalho.DataJanela = objAtenPortab.DtHrPrevPortabilidade
    objPortVipCabecalho.DDDTelefone = " ( " & objAtenPortab.CodigoArea & " )  " & Left(objAtenPortab.NumeroInstancia, 4) & "-" & Right(objAtenPortab.NumeroInstancia, 4)
    objPortVipCabecalho.NumeroBilhete = objAtenPortab.NumBilheteAtual
    objPortVipCabecalho.EmpresaReceptora = objAtenPortab.NomeEmpresaReceptora
    objPortVipCabecalho.EmpresaDoadora = objAtenPortab.NomeEmpresaDoadora
    objPortVipCabecalho.CNL = objAtenPortab.SiglaCNLDestino
    If objAtenPortab.IndicaPortabilidade = 1 Then
        objPortVipCabecalho.IndPortab = "SIM"
    Else
        objPortVipCabecalho.IndPortab = "NÃO"
    End If
    
    'objForm.lblProduto.Caption = objDadosProduto.Descricao
    objForm.Caption = objDadosProduto.Descricao & " - Janela de Migração"
    
    
    
    Call CarregaCabecalho_Portabilidade(objForm.UsrPortabilidadeCabecalho, objPortVipCabecalho)
    
    

        TipoOperacao = "I"     'Indica que a fase está sendo executada pela 1a vez


    'Carrega Tela
    objForm.txtStatusTelefone.Text = StatusTelefone
    
     objForm.spdDataAtivacaoTecnica.Col = 1
     objForm.spdDataAtivacaoTecnica.Row = 1
     
     objForm.spdHoraAtivacaoTecnica.Col = 1
     objForm.spdHoraAtivacaoTecnica.Row = 1
     
     sDataFinalizacao = objSGAP_PUB_BW.RecuperarDtHrRealPort(objAtenPortab.CodigoArea, objAtenPortab.NumeroInstancia)
     
     objForm.spdDataAtivacaoTecnica.Text = Left(sDataFinalizacao, 10)
     objForm.spdHoraAtivacaoTecnica.Text = Right(sDataFinalizacao, 8)
    
'    If lProduto_ID = GbIDProdutoAAC Then
'
'        If objDadosFase.VerificaFaseFinalizada(Fase_ID) = True Then
'            objForm.spdDataAtivacaoTecnica.Enabled = False
'            objForm.spdHoraAtivacaoTecnica.Enabled = False
'            objForm.CmdSalvar.Enabled = False
'
'            sDataFinalizacao = objDadosFase.ObtemDataFimReal(Fase_ID)
'
'            objForm.spdDataAtivacaoTecnica.Text = sDataFinalizacao
'
'        Else
'
'            sHrAtual = Format(Now, "hh:mm")
'            sHrJanelaPort = Format(objAtenPortab.DtHrPrevPortabilidade, "hh:mm")
'
'            If Date > CDate(objAtenPortab.DtHrPrevPortabilidade) Then
'                objForm.CmdSalvar.Enabled = True
'            End If
'
'            If Date = CDate(objAtenPortab.DtHrPrevPortabilidade) Then
'
'                    If CDate(sHrAtual) > CDate(Format(DateAdd("h", 2, sHrJanelaPort), "hh:mm")) Then
'                            objForm.CmdSalvar.Enabled = True
'                    Else
'                            objForm.CmdSalvar.Enabled = False
'                    End If
'            End If
'
'            If Date < CDate(objAtenPortab.DtHrPrevPortabilidade) Then
'                objForm.CmdSalvar.Enabled = False
'            End If
'
'        End If
'
'    End If

    Grupo_ID = objOCSelecionarOTS.Grupo_ID
    Produto_ID = objOCSelecionarOTS.Produto_ID
    
    '=======================================================================================
    '=======================================================================================
    'Só habilita o botão se a fase de janela de migração estiver com a data Inicio Real preenchida
    If lProduto_ID <> GbIDProdutoAAC Then
        bRetorno = objDadosFase.VerificaFaseIniciada(Fase_ID)
        
        If bRetorno = False Then
            Call DesabilitaObjetosTela
        End If
    End If
    
    '=======================================================================================
    '=======================================================================================
    
    Set objDadosPendencia = New clsblf_Pendencia
    
    objDadosPendencia.InformaBanco frmgenMDI_SGAPlus.objBanco
    objDadosPendencia.InformaMensagem objMensagem
   
        
        'Caso exista pendencia para a fase obtem os dados de pendencia e carrega a
        'propriedade do formulário com o objeto Pendencia
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
        
    '===========================================================================
    'CALBAT - Acrescido para pertir alteração  de HIQ
    '===========================================================================
        
            If Not objDadosFase.FaseFinalizada(lFase_Id) Then
            
                objForm.CmdSalvar.Enabled = True
                
            Else
            
                objForm.CmdSalvar.Enabled = False
            
            End If
        
        'Inicio - SPED - PTS 818 - 09/09/2009 - Joao Carlos
        'objForm.UsrVLRDesignacaoLogica1.HabilitaObjetos (objForm.CmdSalvar.Enabled)
        
'        objForm.UsrVLRDesignacaoLogica1.SiglaLocalidade = objAtenTecVoz.Sigla_Localidade
'        objForm.UsrVLRDesignacaoLogica1.ElementoRede = objAtenTecVoz.Tipo_Elemento_Rede
'        objForm.UsrVLRDesignacaoLogica1.IdentEnderRede = objAtenTecVoz.Identif_Endereco_Elem_Rede
'        objForm.UsrVLRDesignacaoLogica1.IdentCluster = objAtenTecVoz.Identif_Cluster
        
        objForm.TxtSiglaLocalidade.Text = objAtenTecVoz.Sigla_Localidade
        objForm.TxtElementoRede.Text = objAtenTecVoz.Tipo_Elemento_Rede
        objForm.TxtIdentEnderRede.Text = objAtenTecVoz.Identif_Endereco_Elem_Rede
        objForm.TxtIdentCluster.Text = objAtenTecVoz.Identif_Cluster
        
        objForm.txtCNL.Text = objAtenPortab.SiglaCNLDestino
        
        'AGRC - SO-TIEMP-49824BGM - 15/11/2017
        'If lProduto_ID = GbIDProdutoNetPhone Or lProduto_ID = GbIDProdutoVlx Then
        If lProduto_ID = GbIDProdutoNetPhone Or lProduto_ID = GbIDProdutoVlx Or lProduto_ID = GbIDProdutoClaroFone Then
            objForm.TxtSiglaLocalidade.Enabled = objForm.CmdSalvar.Enabled
            objForm.TxtElementoRede.Enabled = objForm.CmdSalvar.Enabled
            objForm.TxtIdentEnderRede.Enabled = objForm.CmdSalvar.Enabled
            objForm.TxtIdentCluster.Enabled = objForm.CmdSalvar.Enabled
            objForm.txtCNL.Enabled = objForm.CmdSalvar.Enabled
        Else
            objForm.TxtSiglaLocalidade.Enabled = False
            objForm.TxtElementoRede.Enabled = False
            objForm.TxtIdentEnderRede.Enabled = False
            objForm.TxtIdentCluster.Enabled = False
            objForm.txtCNL.Enabled = False
        End If
        'Fim - SPED - PTS 818 - 09/09/2009 - Joao Carlos

        
    
    '===========================================================================
    '===========================================================================
    'AGRC - SO-TIEMP-49824BGM - 15/11/2017
    If lProduto_ID = GbIDProdutoClaroFone Then
        objForm.frmDadosTecClaroFone.Visible = True
        objForm.TxtTerminalBTS.Text = objAtenTecVoz.TerminalBTS
        objForm.TxtTelefoneBLD.Text = objAtenTecVoz.TelefoneBLD
        objForm.TxtSimcard.Text = objAtenTecVoz.Simcard
        objForm.TxtSimcardNew.Text = objAtenTecVoz.SimcardNew
        If objAtenTecVoz.HFC_Migrado = 1 Then
            objForm.LblHFC_Migrado.Visible = True
        Else
            objForm.LblHFC_Migrado.Visible = False
        End If
   Else
        objForm.frmDadosTecClaroFone.Visible = False
   End If
   
    objForm.informa_OC Me
    
    objForm.Show
    
    'Executa método para popular os campos do Controle Cabecalho
    CarregarForm = True

Finalizar:
    
    Set objdadosServico = Nothing
    Set objAutenticacao = Nothing
    Set objDadosFase = Nothing
    Set objDadosProduto = Nothing
    
    '================================================================
    'CALBAT - Acrescido para pertir alteração  de HIQ
    '================================================================
    Set objAtenTecVoz = Nothing
    Set objSGAP_PUB_BW = Nothing
    '================================================================
    '================================================================
    
    Set objCentrosLocais = Nothing
    
    Exit Function
    
ErrorHandler:

    If Err.Number <> 0 Then
       objMensagem.TipoMsg = "E"
       objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                   "Houve um erro de execução no Método: " & vbCrLf & vbCrLf & "CarregarForm " & _
                   "da classe clsblfOC_TratarCentroLocal"
    End If
       
    objMensagem.ExibeMensagem
       
    GoTo Finalizar
       
End Function
Public Function VerificaPendenciaEmberto() As Boolean

    Dim bRetorno                As Boolean
    Dim objDetalhePendencia     As clsBlf_DetalhePendencia
    
    VerificaPendenciaEmberto = False
    
    For Each objDetalhePendencia In objDadosPendencia.ColecaoPendencia
        
        If objDetalhePendencia.Status = "A" Then
        
           VerificaPendenciaEmberto = True
           
           Exit For
           
        End If
    
    Next

End Function
Public Function GravarDados(pbFlagDescarregaFormulario As Boolean) As Boolean

    Dim objdadosServico             As New clsGenServico
    Dim objDadosFase                As New clsgenFase
    Dim sSituacaoOTS                As String
    Dim bRetorno                    As Boolean
    Dim FlagIniciouTransacao        As Boolean
    Dim numUltimaFaseRealizada      As Long
    Dim FlagTemPendenciaEmAberto    As Boolean
    Dim FlagFaseFinalizadaOK        As Boolean
    Dim FlagTemPendenciaCliente     As Boolean
    Dim lSequenciaAcao              As Long
    Dim FlagManterPeriodo           As Boolean
    Dim lIdFaseSeguinte             As Long
    
    Dim ilLinha                         As Integer
    Dim objDadosProduto                 As New clsGenProduto
    Dim bVerificaFinalizacaoVariosCFS   As Boolean
    Dim bRetirarFinalizacao             As Boolean
    Dim iSeqAux                         As Integer
    Dim sDataAcao                      As String
    Dim objConfServAuto             As New VIP_NGN_AutoExec
    
    Dim objAtenTecVoz                As VIP_Aten_Tec_Voz
    
    'Inicio - SPED - PTS 818 - 09/09/2009 - Joao Carlos
    Dim objCircuito                 As clsgenCircuito
    Dim objAtenPortab               As VIP_Aten_Portabilidade
    'Fim - SPED - PTS 818 - 09/09/2009 - Joao Carlos


    On Error GoTo ErrorHandler
    
    GravarDados = False
    
    Screen.MousePointer = vbArrowHourglass
    
    FlagIniciouTransacao = False
    
    FlagTemPendenciaCliente = False
    
'    If Grupo_ID = GbIDGrupoVLR Then
'
'    End If
    
   If Grupo_ID <> GbIDGrupoVLR And Grupo_ID <> GbIDGrupoVLI Then
        If CDate(objForm.spdDataAtivacaoTecnica.Text) < Format(Date, "dd/mm/yyyy") Then
    
            objMensagem.TipoMsg = "E"
            objMensagem.Descricao = "Data da Janela não pode ser menor que a data atual."
        
            objForm.spdDataAtivacaoTecnica.Text = Format(Date, "dd/mm/yyyy")
            objForm.spdDataAtivacaoTecnica.SetFocus
        
            GoTo ErrorHandler
    
        End If
    End If
    
    objDadosProduto.InformaBanco frmgenMDI_SGAPlus.objBanco
    objDadosProduto.InformaMensagem objMensagem
    
    bRetorno = objDadosProduto.RecuperarDados(Produto_ID)
    
    If objMensagem.TipoMsg <> "" Then
       
       objMensagem.ExibeMensagem
       Set objDadosProduto = Nothing
       
       GoTo Finalizar
    End If
    
    Set objAtenPortab = New VIP_Aten_Portabilidade
    
    objAtenPortab.InformaBanco frmgenMDI_SGAPlus.objBanco
    objAtenPortab.InformaMensagem objMensagem
    
    objForm.spdDataAtivacaoTecnica.Col = 1
    objForm.spdDataAtivacaoTecnica.Row = 1
    
    objForm.spdHoraAtivacaoTecnica.Col = 1
    objForm.spdHoraAtivacaoTecnica.Row = 1
    
    
    sDataAcao = objForm.spdDataAtivacaoTecnica.Text & " " & objForm.spdHoraAtivacaoTecnica.Text
    
    'Inicia transação da fase de Janela de Migração
    If FlagIniciouTransacao = False Then
        bRetorno = frmgenMDI_SGAPlus.objBanco.Iniciar_Transacao()
        FlagIniciouTransacao = True
    End If
    
    If bRetorno = False Then
        GoTo ErrorHandler
    End If
    
    
    Set objAtenTecVoz = New VIP_Aten_Tec_Voz

    objAtenTecVoz.InformaBanco frmgenMDI_SGAPlus.objBanco
    objAtenTecVoz.InformaMensagem objMensagem
    
    'Inicio - SPED - PTS 818 - 09/09/2009 - Joao Carlos
    'objAtenTecVoz.Sigla_Localidade = objForm.UsrVLRDesignacaoLogica1.SiglaLocalidade
    'objAtenTecVoz.Tipo_Elemento_Rede = objForm.UsrVLRDesignacaoLogica1.ElementoRede
    'objAtenTecVoz.Identif_Endereco_Elem_Rede = objForm.UsrVLRDesignacaoLogica1.IdentEnderRede
    'objAtenTecVoz.Identif_Cluster = objForm.UsrVLRDesignacaoLogica1.IdentCluster
    objAtenTecVoz.Sigla_Localidade = objForm.TxtSiglaLocalidade
    objAtenTecVoz.Tipo_Elemento_Rede = objForm.TxtElementoRede
    objAtenTecVoz.Identif_Endereco_Elem_Rede = objForm.TxtIdentEnderRede
    objAtenTecVoz.Identif_Cluster = objForm.TxtIdentCluster
    
    bRetorno = objAtenTecVoz.AlteraDesignacaoLogica(Servico_ID)
    
    If objMensagem.TipoMsg <> "" Then
    
        objMensagem.ExibeMensagem
    
    End If
    
    'AGRC - SO-TIEMP-49824BGM - 15/11/2017
    objAtenTecVoz.TerminalBTS = objForm.TxtTerminalBTS
    objAtenTecVoz.TelefoneBLD = objForm.TxtTelefoneBLD
    objAtenTecVoz.Simcard = objForm.TxtSimcard
    objAtenTecVoz.SimcardNew = objForm.TxtSimcardNew
      
    bRetorno = objAtenTecVoz.AlteraDadosHFC(Servico_ID)
    
    If objMensagem.TipoMsg <> "" Then
    
        objMensagem.ExibeMensagem
    
    End If
    
    
    'Atualiza o CNL do Circuito
    Set objCircuito = New clsgenCircuito
    
    objCircuito.InformaBanco frmgenMDI_SGAPlus.objBanco
    objCircuito.InformaMensagem objMensagem
    
    bRetorno = objCircuito.RecuperarDados(Servico_ID)
    If objMensagem.TipoMsg <> "" Then
        objMensagem.ExibeMensagem
    End If
    
    objCircuito.CNL_Ponta_A = objForm.txtCNL.Text
    
    bRetorno = objCircuito.AtualizarCNL(objCircuito.ID_CIRCUITO, objCircuito.CNL_Ponta_A)
    If objMensagem.TipoMsg <> "" Then
        objMensagem.ExibeMensagem
    End If
    
    objAtenPortab.SiglaCNLDestino = objForm.txtCNL.Text
    
    bRetorno = objAtenPortab.AtualizaSiglaCNLDestino(Servico_ID, objAtenPortab.SiglaCNLDestino)
    If objMensagem.TipoMsg <> "" Then
        objMensagem.ExibeMensagem
    End If
    'Fim - SPED - PTS 818 - 09/09/2009 - Joao Carlos
    
    
    If Grupo_ID <> GbIDGrupoVLI And Grupo_ID <> GbIDGrupoVLR Then
        
        'callbat
        bRetorno = objAtenPortab.AtualizaDataJanelaMigracao(sDataAcao, Servico_ID, Acao)
    
        If objMensagem.TipoMsg <> "" Then
       
           objMensagem.ExibeMensagem
       
        End If
            
        
    End If
    
    'envia para a confserv_Tipo_Fase
    Set objConfServAuto = New VIP_NGN_AutoExec

    objConfServAuto.InformaBanco frmgenMDI_SGAPlus.objBanco
    
    '================================================================
    'Calbat - 03/02/2009
    'Alteração feita para atender a nova estrutura do envio do Confserv
    '================================================================
    objConfServAuto.FlagExecucao = 0
    objConfServAuto.EmProcesso = 0
    
    
    objAtenTecVoz.InformaBanco frmgenMDI_SGAPlus.objBanco
    objAtenTecVoz.InformaMensagem objMensagem
    

    
    '================================================================
    '================================================================
    
    objConfServAuto.ServicoID = Servico_ID

    If Produto_ID = GbIDProdutoAAC Then
        
        objConfServAuto.CodigoTIpoFase = 4
        
        bRetorno = objConfServAuto.IncluirDados(objMensagem)
        
        If (objMensagem.TipoMsg <> "") Then
            GoTo ErrorHandler
        End If
    
    Else
    
        '===================================================================================
        'CALLBAT - 24/06/2009
        ' Incluso para alterar a Designação Lógica
        '===================================================================================
        
        'objAtenTecVoz.Sigla_Localidade = objForm.UsrVLRDesignacaoLogica1.SiglaLocalidade
        'objAtenTecVoz.Tipo_Elemento_Rede = objForm.UsrVLRDesignacaoLogica1.ElementoRede
        'objAtenTecVoz.Identif_Endereco_Elem_Rede = objForm.UsrVLRDesignacaoLogica1.IdentEnderRede
        'objAtenTecVoz.Identif_Cluster = objForm.UsrVLRDesignacaoLogica1.IdentCluster
        
        '===================================================================================
        '===================================================================================
    
        bRetorno = objConfServAuto.VerificarRegistroFase(Servico_ID, 40, objMensagem)
        
        If bRetorno = True Then
        
            objConfServAuto.CodigoTIpoFase = 40
                    
            'bRetorno = objConfServAuto.IncluirDados(objMensagem)
            
            bRetorno = objConfServAuto.AlterarDados(Servico_ID, objConfServAuto.CodigoTIpoFase, objMensagem)
            
            If (objMensagem.TipoMsg <> "") Then
                GoTo ErrorHandler
            End If
            
            If Grupo_ID <> GbIDGrupoVLI And Grupo_ID <> GbIDGrupoVLR Then
            
                        'finaliza a fase corrente
                        Set objDadosFase = New clsgenFase
                        
                        objDadosFase.InformaBanco frmgenMDI_SGAPlus.objBanco
                        objDadosFase.InformaMensagem objMensagem
                        
                        bRetorno = objDadosFase.FinalizaFase(Fase_ID)
                        
                        If objMensagem.TipoMsg <> "" Then
                            GoTo ErrorHandler
                        End If
                        
                         lSequenciaAcao = objDadosFase.ObtemSequenciadaFase(Servico_ID, GbCodigoFasePortabilidadeJanelaMigracao)
                        
                        'Calcula o período para a proxima fase
                        bRetorno = objDadosFase.CalculaPeriodoProximaFase(lSequenciaAcao, Servico_ID)
                        
                        If (objMensagem.TipoMsg <> "") Then
                        
                            GoTo ErrorHandler
                            
                        End If
                        
                        
                        'Retira, pendencia caso tenha
                        'Verifica se Existe pendencia a ser incluida
                        If objDadosPendencia.ColecaoPendencia.Count <> 0 Then
                        
                           bRetorno = objDadosPendencia.AtualizarDados(Fase_ID, Servico_ID, _
                                                                       Produto_ID, FlagTemPendenciaEmAberto, _
                                                                       objdadosServico.DiretorRegional, _
                                                                       objdadosServico.NumOE, _
                                                                       objdadosServico.AnoOE, _
                                                                       objdadosServico.ItemOE, _
                                                                       sSiglaProdutoSOE, _
                                                                       "")
                                If (Not bRetorno) Then
                                   GoTo ErrorHandler
                                End If
                           
                        End If
                    
            End If
                    
        End If
                
    End If
    
    If FlagIniciouTransacao Then

       bRetorno = frmgenMDI_SGAPlus.objBanco.Finalizar_Transacao
        GravarDados = True
    
    End If
    
    objMensagem.TipoMsg = "A"
    
    objMensagem.Descricao = "Fase finalizada com Sucesso! "
    objMensagem.ExibeMensagem
    
    objForm.CmdSalvar.Enabled = False
    
Finalizar:
    Set objMensagem = Nothing
    Set objDadosTratarCentroLocal = Nothing
    Set objdadosServico = Nothing
    Set objDadosFase = Nothing
    Set objDadosProduto = Nothing
    
    Set objAtenPortab = Nothing
    Set objConfServAuto = Nothing
    Set objDadosFase = Nothing
    
    'Inicio - SPED - PTS 818 - 09/09/2009 - Joao Carlos
    Set objCircuito = Nothing
    'Fim - SPED - PTS 818 - 09/09/2009 - Joao Carlos
    
    Screen.MousePointer = vbDefault
    
    Exit Function
    
ErrorHandler:
    Screen.MousePointer = vbDefault

    ' Cancelar Transação
    If FlagIniciouTransacao Then
       bRetorno = frmgenMDI_SGAPlus.objBanco.Cancelar_Transacao
    End If
    
    If Err.Number <> 0 Then
       objMensagem.TipoMsg = "E"
       
       objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                   "Houve um erro de execução no Método: " & vbCrLf & vbCrLf & "GravarDados " & _
                   "da classe clsblfOC_TratarCentroLocal"
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
    
    If objDadosTratarCentroLocal.DtHoraConclusao <> "" Then
                
        VerificaFinalizacaoOK = True
        
    End If

End Function


Private Function DesabilitaObjetosTela()
    Dim ilLinha     As Integer

    objForm.CmdSalvar.Enabled = False
    objForm.spdDataAtivacaoTecnica.Enabled = False
    objForm.spdHoraAtivacaoTecnica.Enabled = False
    'objForm.spdDataAtivacaoTecnica.EditMode = False
    'objForm.spdHoraAtivacaoTecnica.EditMode = False
    objForm.cmdImpOTS.Enabled = False

    bPermiteAtualizar = False
    
    
End Function

Public Function LimpaTela()
    Dim ilLinha     As Integer

    'objForm.SpdDataConclusao.Text = ""
    
    'Checa se o produto pode alocar vários CFs para a fase Tratar Centro Local
    
    'Localiza o registro com o CF atual.
'    If objDadosDuracaoFase.ChecaPermissaoMultiplosCF(Produto_ID, _
'                GbCodigoFaseCentroLocal) Then
'
'        For ilLinha = 1 To objForm.grdCentroLocal.MaxRows
'
'            With objForm.grdCentroLocal
'
'                .Col = 3        'Centro Funcional ID
'                .Row = ilLinha
'
'                If CentroFuncionalID = .Text Then
'                    .Col = 2
'                    .Text = ""
'
'                End If
'            End With
'
'        Next ilLinha
'    End If
    
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
                            "InformaOCSelecionarOTS da classe clsblfOc_TratarCentroLocal"
    
    GoTo Finalizar

End Function

Public Function AtualizarListaOTS() As Boolean

    objOCSelecionarOTS.ExibirListaOTS
        
End Function

Private Sub Class_Terminate()

    Set objDadosPendencia = Nothing
    Set objMensagem = Nothing
    Set objDadosTratarCentroLocal = Nothing
    Set objForm = Nothing
    Set objDadosDuracaoFase = Nothing

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

Public Function ExcluirDados(pDataFinalizacao As String) As Long

    Dim objdadosServico           As New clsGenServico
    Dim objDadosFase              As New clsgenFase
    
    Dim sSituacaoOTS              As String
    Dim bRetorno                  As Boolean
    Dim FlagIniciouTransacao      As Boolean
    Dim numUltimaFaseRealizada    As Long
    Dim FlagManterPeriodo         As Boolean
    
    Dim bPermiteAtualizar         As Boolean
    
    On Error GoTo ErrorHandler
    
    bPermiteAtualizar = False
    FlagIniciouTransacao = False
    
    ExcluirDados = False
    'Verfica se a OTS selecionada ainda encontra-se em andamento
    objdadosServico.InformaBanco frmgenMDI_SGAPlus.objBanco
    
    'Seta o objeto mensagem da classe de dados
    objdadosServico.InformaMensagem objMensagem
    
    sSituacaoOTS = objdadosServico.RecuperarSituacaoOTS(Servico_ID)
    
    Set objdadosServico = Nothing
    
    If objMensagem.TipoMsg <> "" Then
       
       GoTo ErrorHandler
       
    End If
    
    If sSituacaoOTS = "CAN" Or sSituacaoOTS = "CR" Then
       
       objMensagem.TipoMsg = "A"
       objMensagem.Descricao = "Esta OTS foi paralisada pelo SOE!"
       objMensagem.ExibeMensagem
       
       Set objMensagem = Nothing
       
       GoTo Finalizar
       
    End If

    'Verifica se fase ainda encontra-se em Andamento
    objDadosFase.InformaBanco frmgenMDI_SGAPlus.objBanco
    objDadosFase.InformaMensagem objMensagem

    If (Not objDadosFase.RecuperarSituacaoFase(Fase_ID)) And _
       pDataFinalizacao = "" Then

       Set objDadosFase = Nothing

       objMensagem.ExibeMensagem

       ExcluirDados = True

       GoTo Finalizar
    Else
       
       objMensagem.TipoMsg = ""
    
    End If

    'Obtem o número da última fase realizada
    numUltimaFaseRealizada = objDadosFase.ObtemSequenciaUltimaFaseRealizada(Servico_ID, GbCodigoFaseProgramarCPE)
    
    If objMensagem.TipoMsg <> "" Then
       
       objMensagem.ExibeMensagem
       GoTo Finalizar
    
    End If
    
    If numUltimaFaseRealizada > Sequencia_Acao Then
    
       objMensagem.TipoMsg = "A"
       objMensagem.Descricao = "Os dados desta fase não poderão excluídos! Pois, a fase seguinte já foi iniciada."
       objMensagem.ExibeMensagem
       
       Set objMensagem = Nothing
       ExcluirDados = True
       
       GoTo Finalizar
    
    End If

    If objDadosPendencia.ColecaoPendencia.Count <> 0 Then
    
        objMensagem.TipoMsg = "A"
        objMensagem.Descricao = "Existe(m) Pendência(s) cadastrada(s) para esta fase! " & _
                                "A exclusão não poderá ser realizada!"
        objMensagem.ExibeMensagem
        
        GoTo Finalizar
        
    End If

    'Seta o objeto mensagem da classe de dados
    objDadosTratarCentroLocal.InformaMensagem objMensagem
    objDadosTratarCentroLocal.InformaBanco frmgenMDI_SGAPlus.objBanco

    bRetorno = frmgenMDI_SGAPlus.objBanco.Iniciar_Transacao()
        
    If Not bRetorno Then
    
       GoTo ErrorHandler
       
    End If
    
    FlagIniciouTransacao = True
    
    'Exclui dados de Tratar Centro Local
    ExcluirDados = objDadosTratarCentroLocal.ExcluirDados(Fase_ID)
   
    If Not ExcluirDados Then
       
       GoTo ErrorHandler
    
    End If
    
    'objForm.SpdDataConclusao.Text = ""
    
    'Checa se o produto pode alocar vários CFs para a fase Tratar Centro Local
    '''If objDadosDuracaoFase.ChecaPermissaoMultiplosCF(Produto_ID, _
                GbCodigoFaseCentroLocal) Then
        
    If bPermiteAtualizar Then
        bRetorno = objDadosFase.ExcluiData_VariosCFS(Fase_ID, CentroFuncionalID)
        
        If objMensagem.TipoMsg <> "" Then
           
           objMensagem.ExibeMensagem
           GoTo Finalizar
        
        End If
    End If
    
    '''End If

    objDadosFase.InformaBanco frmgenMDI_SGAPlus.objBanco
    objDadosFase.InformaMensagem objMensagem
    
    objDadosFase.ID_Servico = Servico_ID
    objDadosFase.Id_Fase = Fase_ID
    
    'Quando for OTS de Desativacao ou Cancelamento, o periodo previsto para a fase seguinte
    'deverá ser mantido já que as fases para estas ações serão executadas em paralelo.
    
    If Sigla_Acao = "DES" Or Sigla_Acao = "CAN" Then
    
        FlagManterPeriodo = True
        
    Else
    
        FlagManterPeriodo = False
    
    End If
    
    bRetorno = objDadosFase.RetirarFinalizacaoFase(Servico_ID, Fase_ID, Sequencia_Acao, FlagManterPeriodo)

    If objMensagem.TipoMsg <> "" Then
       
       Set objDadosFase = Nothing
       GoTo ErrorHandler
    
    End If
    
    'Grava Anotacao da fase Vazia
    If Trim(TextoAnotacao) <> "" Then
       
       objDadosFase.GravarAnotacao Fase_ID, ""
         
       If objMensagem.TipoMsg <> "" Then
            
          GoTo ErrorHandler
      
       End If
    
    End If
    
    Set objDadosFase = Nothing
    
    bRetorno = frmgenMDI_SGAPlus.objBanco.Finalizar_Transacao
    
    If Not bRetorno Then
        
        GoTo ErrorHandler
    
    End If
    
    ExcluirDados = True
    
    GoTo Finalizar
    
Finalizar:
    
    Set objdadosServico = Nothing
    Set objDadosFase = Nothing
    
    Exit Function
    
ErrorHandler:
       
    ' Cancelar Transação
    If FlagIniciouTransacao Then
       
       bRetorno = frmgenMDI_SGAPlus.objBanco.Cancelar_Transacao
    
    End If
    
    If Err.Number <> 0 Then
       objMensagem.TipoMsg = "E"
       objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                   "Houve um erro de execução no Método: " & vbCrLf & vbCrLf & "ExcluirDados " & _
                   "da classe clsblfOC_TratarCentroLocal "
    End If
       
    objMensagem.ExibeMensagem
    
    GoTo Finalizar

End Function

Public Function CamposPreenchidosOK() As Boolean

    Dim objDadosFase    As clsgenFase
    Dim bRetorno        As Boolean
    Dim sDataFim        As String
    
    CamposPreenchidosOK = False

'    If objForm.SpdDataConclusao.Text = "" Then
'
'       objMensagem.TipoMsg = "A"
'       objMensagem.Descricao = "Data de Conclusão deve ser informada!"
'       objMensagem.ExibeMensagem
'
'       objForm.SpdDataConclusao.SetFocus
'
'       Exit Function
'
'    End If
    
'''    If Not IsDate(objForm.SpdDataConclusao.Text) And objForm.SpdDataConclusao.Text <> "" Then
'''
'''       objMensagem.TipoMsg = "A"
'''       objMensagem.Descricao = "Data de Conclusão informada não é uma data válida!"
'''       objMensagem.ExibeMensagem
'''
'''       objForm.SpdDataConclusao.SetFocus
'''
'''       Exit Function
'''
'''    End If

    Set objDadosFase = New clsgenFase
    
    objDadosFase.InformaBanco frmgenMDI_SGAPlus.objBanco
    objDadosFase.InformaMensagem objMensagem
    
    sDataFim = objDadosFase.ObtemDataFimRealFaseAnterior(Servico_ID, Sequencia_Acao)
    
    Set objDadosFase = Nothing

''''    If Trim(objForm.SpdDataConclusao.Text) <> "" And sDataFim <> "" And _
''''       (Sigla_Acao <> "DES" And Sigla_Acao <> "CAN") Then
''''
''''        If CDate(sDataFim) > CDate(objForm.SpdDataConclusao.Text) Then
''''
''''            objMensagem.TipoMsg = "A"
''''            objMensagem.Descricao = "Data de Conclusão não pode ser menor que " & _
''''                                    "a data de Finalização da fase anterior!"
''''
''''            objMensagem.ExibeMensagem
''''
''''            objForm.SpdDataConclusao.SetFocus
''''
''''            Exit Function
''''
''''        End If
''''
''''    End If
    
    CamposPreenchidosOK = True

End Function

Public Function InformaBanco(ByVal objNomeBanco As SGAPLUSBanco.clsBanco)

On Error GoTo ErrorHandler:

    InformaBanco = False
    
    Set objBanco = objNomeBanco
    
    InformaBanco = True
    
Finalizar:

    Exit Function
    
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: InformaBanco " & _
                            "da classe clsgenServico"
    
    GoTo Finalizar

End Function

Public Function InformaMensagem(ByVal objNomeMensagem As clsgenMensagem)

On Error GoTo ErrorHandler:

    InformaMensagem = False
    
    Set objMensagem = objNomeMensagem
    
    InformaMensagem = True
    
Finalizar:

    Exit Function
    
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: InformaMensagem " & _
                            "da classe clsgenServico"
    
    GoTo Finalizar

End Function

