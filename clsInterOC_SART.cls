VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInterOC_SART"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private lProduto_ID As Long
Private lGrupo_ID As Long
Private lServico_ID As Long

Private sSigla_Acao As String
Private sStatus_Fase As String

Private lServico_ID_Anterior As Long

Private lFase_Id As Long

Private objMensagem                    As New clsgenMensagem

Dim objOC_TratarAcessoEbt        As clsGenOC_TratarAcessoEmbratel
Dim objDadosFaseAcesso             As clsgenFase

Dim objDadosTipoCabo                As clsGenTipoAcessoCabo

'Característica do acesso
Dim sTipoAcesso                 As String 'ADE(A)   CABO(C) ADE & CABO(CA)

Dim bFlagAcessoNaoExiste   As Boolean
Dim lIdFaseAcesso               As Long

Public objForm                  As frmPortSART

Public objOCSelecionarOTS       As clsgenOC_SelecionarOTS

Public objDadosPendencia        As New clsblf_Pendencia



Private sFamilia_Produto As Long
Public Function AtualizarListaOTS() As Boolean

    objOCSelecionarOTS.ExibirListaOTS
        
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
'Public Function VerificaPendenciaEmberto() As Boolean
'
'    Dim bRetorno                As Boolean
'    Dim objDetalhePendencia     As clsBlf_DetalhePendencia
'
'    VerificaPendenciaEmberto = False
'
'    For Each objDetalhePendencia In objDadosPendencia.ColecaoPendencia
'
'        If objDetalhePendencia.Status = "A" Then
'
'           VerificaPendenciaEmberto = True
'
'           Exit For
'
'        End If
'
'    Next
'
'End Function
Public Function ReenviaSart() As Boolean

Dim FlagTemPendenciaEmAberto As Boolean

Dim lIdFaseSeguinte             As Long
Dim objConfServAuto             As VIP_NGN_AutoExec

    On Error GoTo ErrorHandler

    FlagTemPendenciaEmAberto = False
    
    
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

    
    
    If (objDadosPendencia.ColecaoPendencia.Count < 0) Then
       
        Set objConfServAuto = New VIP_NGN_AutoExec
        
        objConfServAuto.InformaBanco frmgenMDI_SGAPlus.objBanco
        
        objConfServAuto.CodigoTIpoFase = GbCodigoFaseSart
        
        '================================================================
        'Calbat - 03/02/2009
        'Alteração feita para atender a nova estrutura do envio do Confserv
        '================================================================
        objConfServAuto.FlagExecucao = 0
        objConfServAuto.EmProcesso = 0
        '================================================================
        '================================================================
        
        objConfServAuto.ServicoID = Servico_ID
        
        bRetorno = objConfServAuto.IncluirDados(objMensagem)
        
        Set objConfServAuto = Nothing
        
        If objMensagem.TipoMsg <> "" Then
           
           GoTo ErrorHandler

        End If
       
       'frmgenMDI_SGAPlus.objBanco.Parametros_SP_Valor
       FlagFaseFinalizadaOK = True
    
    End If



Finalizar:
    
'    Set objAtivacaoBLF = Nothing
'    Set objdadosTipoADE = Nothing
'    Set objDadosTipoCabo = Nothing
'    Set objDadosFaseAcesso = Nothing
'    Set objOC_TratarAcessoEbt = Nothing
    
    'Set objForm = Nothing
    
    Exit Function
    
ErrorHandler:

    If Err.Number <> 0 Then
       objMensagem.TipoMsg = "E"
       objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                   "Houve um erro de execução no Método: " & vbCrLf & vbCrLf & "ReenviaSart " & _
                   "da classe clsInterOC_SART"
    End If
       
    objMensagem.ExibeMensagem
       
    GoTo Finalizar
End Function

Public Function CarregarForm() As Boolean

Dim objAtivacaoBLF As clsblf_AtivacaoFinalizacao

    On Error GoTo ErrorHandler
    
    CarregarForm = False
    
    Set objAtivacaoBLF = New clsblf_AtivacaoFinalizacao
    
    Set objForm = New frmPortSART
    
    objAtivacaoBLF.InformaBanco frmgenMDI_SGAPlus.objBanco
    objAtivacaoBLF.InformaMensagem objMensagem
    
    bRetorno = objAtivacaoBLF.RecuperarDados(Fase_ID)
    
    If objMensagem.TipoMsg <> "" Then
        GoTo ErrorHandler
    End If
    
    objForm.UsrVozBandaCabecalho.Visible = True
    
    'Call CarregaCabecalho_VIPLINEIP(objForm.UsrVozBandaCabecalho, objBlfCabecalho)
    
    Call CarregaCabecalho_VIPLINEIP(objForm.UsrVozBandaCabecalho, objBlfCabecalho)

    'objForm.informa_OC Me
    
    objForm.spdDataReprogramada.Text = objAtivacaoBLF.DataReprogramada
    objForm.SpdHrReagendamento.Text = objAtivacaoBLF.HoraReprogramada
    
    objForm.SpdDataAgendTesteAtv.Text = ""
    objForm.SpdHrAgendTesteAtv.Text = ""
    
    objForm.SpdDataAtivacao.Text = objAtivacaoBLF.DataAtivacaoTecnica
    objForm.SpdHrAtivacao.Text = objAtivacaoBLF.HoraAtivacaoTecnica
    
    objForm.SpdDataProgramada.Text = objAtivacaoBLF.DataProgramadaRetirada
    objForm.SpdHrProgramada.Text = objAtivacaoBLF.HoraProgramadaRetirada
    
    objForm.SpdDataConclusao.Text = objAtivacaoBLF.DataConclusaoRetirada
    
    objForm.txtNomeClienteAceite.Text = objAtivacaoBLF.NomeClienteAceite
    objForm.txtNomeTerceira.Text = objAtivacaoBLF.NomeTecnicoEmpresaParceira
    
    '******************************************************************************
    'Exibe informações de Modem
    '******************************************************************************
    If (Grupo_ID = GbIDGrupoVLI) Then
    
        If objAtivacaoBLF.FlagUtilizacaoModem = "S" Then
            
            With objForm.usrVipInfTecModem1

                'Flag do Modem obtido na tabela ATIVACAO_FINALIZACAO_BLF
                .IndicaUtilizacaoModem = objAtivacaoBLF.FlagUtilizacaoModem

                .FabricanteModem = "" & objAtivacaoBLF.FabricanteModem
                .ModeloModem = "" & objAtivacaoBLF.ModeloModem
                .NumeroSerieModem = "" & objAtivacaoBLF.NumeroSerieModem

            End With
            
            objForm.lblInterface.Caption = "Modem"
            
            objForm.FrameTipoADE.Enabled = True
            objForm.FrameTipoADE.ForeColor = &H80000012  'Preto
            objForm.FrameTipoADE.Font.Bold = False
            
            objForm.FrameTipoCabo.Enabled = False
            objForm.FrameTipoCabo.ForeColor = &H80000012  'Preto
            objForm.FrameTipoCabo.Font.Bold = False
            
            objForm.FramInfModem.Enabled = False
            objForm.FramInfModem.ForeColor = &H80&            'vinho
            objForm.FramInfModem.Font.Bold = True
            
        End If
        
    End If
    '******************************************************************************
    '******************************************************************************
    
    '******************************************************************************
    'Obtem informações de Acesso
    '******************************************************************************
    
        'Obtém o Id da Fase Acesso
    Set objDadosFaseAcesso = New clsgenFase
    
    objDadosFaseAcesso.InformaBanco frmgenMDI_SGAPlus.objBanco
    objDadosFaseAcesso.InformaMensagem objMensagem

    'Obtém o ID da Fase de Acesso
    lIdFaseAcesso = objDadosFaseAcesso.ObtemIDFase(Servico_ID, _
                        GbCodigoFaseAcesso)
    
    Set objOC_TratarAcessoEbt = New clsGenOC_TratarAcessoEmbratel
    
    bFlagAcessoNaoExiste = False

    Call objOC_TratarAcessoEbt.RecuperaDadosAcesso("P", Servico_ID, lIdFaseAcesso, bFlagAcessoNaoExiste)
    
    
    If objOC_TratarAcessoEbt.objDadosAcesso.PropridedadeAcesso = "E" Then
    
        sTipoAcesso = objOC_TratarAcessoEbt.objDadosAcesso.objPropriedadeEBT.TipoAcesso
    
    Else
        
        sTipoAcesso = objOC_TratarAcessoEbt.objDadosAcesso.objPropriedadeTerceiros.TecnologiaAcesso
    
    End If

    'Acesso ADE
    If sTipoAcesso = "A" Or sTipoAcesso = "CA" And _
         objAtivacaoBLF.TipoAcessoADEID <> 0 Then
    
        Set objdadosTipoADE = New clsGenTipoAcessoADE
        
        objdadosTipoADE.InformaBanco frmgenMDI_SGAPlus.objBanco
        
        bRetorno = objdadosTipoADE.RecuperarDados(objAtivacaoBLF.TipoAcessoADEID, objMensagem)
        
        objForm.TxtNumeroCabo_ADE.Text = objdadosTipoADE.NumeroCaboEBT
        objForm.TxtNumeroPar_ADE.Text = objdadosTipoADE.NumeroParEBT
        objForm.TxtPredioEBT_ADE.Text = objdadosTipoADE.PredioEBT
        
        objForm.lblInterface.Caption = "ADE"
        
        objForm.FrameTipoADE.Enabled = True
        objForm.FrameTipoADE.ForeColor = &H80&           'vinho
        objForm.FrameTipoADE.Font.Bold = True
        
        objForm.FrameTipoCabo.Enabled = False
        objForm.FrameTipoCabo.ForeColor = &H80000012  'Preto
        objForm.FrameTipoCabo.Font.Bold = False
        
        objForm.FramInfModem.Enabled = False
        objForm.FramInfModem.ForeColor = &H80000012  'Preto
        objForm.FramInfModem.Font.Bold = False
        
    End If
    
    'Cabo Interno
    If sTipoAcesso = "C" Or sTipoAcesso = "CA" And _
       objAtivacaoBLF.TipoAcessoCABOID <> 0 Then
        
        Set objDadosTipoCabo = New clsGenTipoAcessoCabo
        
        objDadosTipoCabo.InformaBanco frmgenMDI_SGAPlus.objBanco
        
        bRetorno = objDadosTipoCabo.RecuperarDados(objAtivacaoBLF.TipoAcessoCABOID, objMensagem)
        
        objForm.TxtNumeroParEBT_Cabo.Text = objDadosTipoCabo.NumeroParEBT
        objForm.TxtNumeroCaboEBT_Cabo.Text = objDadosTipoCabo.NumeroCaboEBT
        objForm.TxtNumeroCaboCliente_Cabo.Text = objDadosTipoCabo.NumeroCaboCliente
        objForm.TxtNumeroParCliente_Cabo.Text = objDadosTipoCabo.NumeroParCliente
        objForm.TxtPredioCabo_Cabo.Text = objDadosTipoCabo.Predio
        
        objForm.lblInterface.Caption = "Cabo Interno"
        
        objForm.FrameTipoADE.Enabled = False
        objForm.FrameTipoADE.ForeColor = &H80000012  'Preto
        objForm.FrameTipoADE.Font.Bold = False
        
        objForm.FrameTipoCabo.Enabled = False
        objForm.FrameTipoCabo.ForeColor = &H80000012  'Preto
        objForm.FrameTipoCabo.Font.Bold = False
        
        objForm.FramInfModem.Enabled = True
        objForm.FramInfModem.BackColor = &H80&               'vinho
        objForm.FramInfModem.Font.Bold = True
    
    End If
    
    '******************************************************************************
    '******************************************************************************
    
    '******************************************************************************
    '******************************************************************************
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
    '******************************************************************************
    '******************************************************************************

    objForm.informa_OC Me
    
    objForm.Show
    
    CarregarForm = True

Finalizar:
    
    Set objAtivacaoBLF = Nothing
    Set objdadosTipoADE = Nothing
    Set objDadosTipoCabo = Nothing
    Set objDadosFaseAcesso = Nothing
    Set objOC_TratarAcessoEbt = Nothing
    
    'Set objForm = Nothing
    
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
Public Property Get Fase_ID() As Long

Fase_ID = lFase_Id

End Property

Public Property Let Fase_ID(ByVal vNewValue As Long)

lFase_Id = vNewValue

End Property

Public Property Get Servico_ID() As Long

Servico_ID = lServico_ID

End Property

Public Property Let Servico_ID(ByVal vNewValue As Long)

lServico_ID = vNewValue

End Property


Public Property Get Sigla_Acao() As String

    Sigla_Acao = sSigla_Acao

End Property

Public Property Let Sigla_Acao(ByVal vNewValue As String)

    sSigla_Acao = vNewValue

End Property


Public Property Get Servico_ID_Anterior() As Long

    Servico_ID_Anterior = lServico_ID_Anterior

End Property

Public Property Let Servico_ID_Anterior(ByVal vNewValue As Long)

    lServico_ID_Anterior = vNewValue

End Property
Public Property Get Status_Fase() As String

    Status_Fase = sStatus_Fase

End Property

Public Property Let Status_Fase(ByVal vNewValue As String)

    sStatus_Fase = vNewValue

End Property

