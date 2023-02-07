VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsgenOC_TratarAcessoDSLAM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Public objDadosAcesso           As New clsGenAcesso 'ACESSO EMBRATEL
Public objForm                  As frmgenTratarAcessoDslam
Public objOCSelecionarOTS       As clsgenOC_SelecionarOTS
Public vMensagemListaMsg        As Variant
Public bFlagControlaAlteracao   As Boolean
Public objMensagem              As New clsgenMensagem

Public objDadosPendencia        As New clsblf_Pendencia

'Classe de dados Acesso Contigência do acesso(EBT)
Public objDadosAcesso           As New clsGenAcesso

'Classe de dados Acesso EBT
Dim objDadosAcessoDSLAM         As New clsGenTratarAcessoDSLAM

Dim objOCTratarAcessoEBT        As New clsGenOC_TratarAcessoEmbratel

'Armazena IDs Acesso Principal e COntigente
Dim lIdAcessoPrincipal          As Long
Dim lIdAcessoContigente         As Long

'Classe de dados Acesso propriedade Embratel
Dim objPropriedadeEBT           As clsGenPropriedadeEmbratel

Dim objdadosTipoADE             As clsGenTipoAcessoADE
'Dim objDadosTipoCabo            As clsGenTipoAcessoCabo

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
Dim sPermiteExecucao            As String
Dim lCentroFuncionalID          As Long

Dim sFamilia_Produto            As String
Dim oobjPendencia               As clsDNI_DadosPendencia

Private lUsuarioID              As Long     'Usuário logado no sistema

Private lUsuarioGravadoID       As Long     'Usuário responsável pela gravação da fase
Private sUsuarioLogin           As String   'Login do usuário logado no sistema ou responsável pela gravação da fase
Dim bRetorno                    As Boolean
Dim FlagDescarregaForm          As Boolean

'Dim FlagTemPendenciaADE         As Boolean

'Armazena informações da propriedade EBT
'Dim sTipoAcessoAnterior           As String
'Dim sPropriedadeAcessoAnterior    As String
'Dim sCaracteristicaAcessoAnterior As String

'Public Property Get CaracteristicaAcessoAnterior() As String
'
'    CaracteristicaAcessoAnterior = sCaracteristicaAcessoAnterior
'
'End Property
'
'Public Property Let CaracteristicaAcessoAnterior(ByVal sNewValue As String)
'
'    sCaracteristicaAcessoAnterior = sNewValue
'
'End Property
''''
'
'
'Public Property Get PropriedadeAcessoAnterior() As String
'
'    PropriedadeAcessoAnterior = sPropriedadeAcessoAnterior
'
'End Property

'Public Property Let PropriedadeAcessoAnterior(ByVal sNewValue As String)
'
'    sPropriedadeAcessoAnterior = sNewValue
'
'End Property
''''
''
'Public Property Get TipoAcessoAnterior() As String
'
'    TipoAcessoAnterior = sTipoAcessoAnterior
'
'End Property
'
'Public Property Let TipoAcessoAnterior(ByVal sNewValue As String)
'
'    sTipoAcessoAnterior = sNewValue
'
'End Property
'''

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

'Public Property Get NumeroOE() As Long
'
'NumeroOE = lNumeroOE
'
'End Property
'
'Public Property Let NumeroOE(ByVal lNewValue As Long)
'
'lNumeroOE = lNewValue
'
'End Property
'
'Public Property Get AnoOE() As Long
'
'AnoOE = lAnoOE
'
'End Property
'
'Public Property Let AnoOE(ByVal lNewValue As Long)
'
'lAnoOE = lNewValue
'
'End Property

'Public Property Get NumeroItemOE() As Long
'
'NumeroItemOE = lNumeroITemOE
'
'End Property
'
'Public Property Let NumeroItemOE(ByVal lNewValue As Long)
'
'lNumeroITemOE = lNewValue
'
'End Property

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

'Public Property Get CentroFuncionalCL_ID() As Long
'
'CentroFuncionalCL_ID = lCentroFuncionalCL_ID
'
'End Property
'
'Public Property Let CentroFuncionalCL_ID(ByVal lNewValue As Long)
'
'lCentroFuncionalCL_ID = lNewValue
'
'End Property
'
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
    Dim paPermissoes()         As Variant
    Dim sSituacaoOTS           As String
    Dim numUltimaFaseRealizada As Long
    Dim IdOTS                  As Long
    Dim lFaseAcessoID          As Long
    Dim cColecao               As Collection
    Dim flUsuarioDoCF          As Boolean
    
    'Não resgata os dados da OTS anterior quando se altera o
    'endereço de instalação. IP VPN
    Dim sCNLAtual              As String
    Dim sCNLAnterior           As String
    Dim objdadosCircuito       As New clsgenCircuito
    Dim bResgataAnterior       As Boolean
    
    Dim sTabela_Interface      As String   'JAH
    Dim sCodigoProdutoSOE      As String
    Dim IDFase                 As Long
    
    On Error GoTo ErrorHandler
    
    CarregarForm = False
    
    'Controla a existência de pendência dsa informações do pedido ADE
    'FlagTemPendenciaADE = False
    
    'Acesso Embratel(Contigente)
    Set objDadosAcesso = New clsGenAcesso
    objDadosAcesso.InformaMensagem objMensagem
    objDadosAcesso.InformaBanco frmgenMDI_SGAPlus.objBanco
    
    'Propriedade Embratel
    Set objPropriedadeEBT = New clsGenPropriedadeEmbratel
    objPropriedadeEBT.InformaBanco frmgenMDI_SGAPlus.objBanco
    
    'Acesso DSLAM
    Set objDadosAcessoDSLAM = New clsGenTratarAcessoDSLAM
    objDadosAcessoDSLAM.InformaMensagem objMensagem
    objDadosAcessoDSLAM.InformaBanco frmgenMDI_SGAPlus.objBanco
    
    'Circuito
    Set objdadosCircuito = New clsgenCircuito
    
    objdadosCircuito.InformaMensagem objMensagem
    objdadosCircuito.InformaBanco frmgenMDI_SGAPlus.objBanco

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
             
    objDadosPendencia.InformaBanco frmgenMDI_SGAPlus.objBanco
    objDadosPendencia.InformaMensagem objMensagem
             
    'Obtem a Sequencia_Acao da última fase realizada
    numUltimaFaseRealizada = objDadosFase.ObtemSequenciaUltimaFaseRealizada(Servico_ID, _
                                GbCodigoFaseProgramarCPE, _
                                Familia_Produto)
    
    If objMensagem.TipoMsg <> "" Then
       
       objMensagem.ExibeMensagem
       
       Set objdadosServico = Nothing
       
       GoTo Finalizar
    End If
    
    'Verifica as permissoes do usuario
    objAutenticacao.IDUsuario = frmgenMDI_SGAPlus.objAutenticacao.IDUsuario
    objAutenticacao.Login = frmgenMDI_SGAPlus.objAutenticacao.Login
    objAutenticacao.InformaMensagem objMensagem
    
    'Armazena informações de login nas propriedades da classe
    lUsuarioID = objAutenticacao.IDUsuario
    sUsuarioLogin = objAutenticacao.Login
    
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
    
    flUsuarioDoCF = objAutenticacao.VerificaUsuarioNoCF(objAutenticacao.IDUsuario, _
                                            CentroFuncionalID)
    Set objAutenticacao = Nothing

    If UBound(paPermissoes, 2) = 0 Then
    
        objMensagem.TipoMsg = "E"
        objMensagem.Descricao = "O usuário não possui direito de consulta"
        objMensagem.ExibeMensagem

        GoTo Finalizar
    
    End If
   
   '======================================================================================
   'Instancia o formulario, obtem os dados da fase e carrega os campos da tela
   '======================================================================================
    Set objForm = New frmgenTratarAcessoDslam
    
    If objMensagem.TipoMsg <> "" Then
       
       GoTo ErrorHandler
       
    End If
    
    objForm.LblAcao_Status = Status_Fase
    
    objDadosProduto.InformaBanco frmgenMDI_SGAPlus.objBanco
    objDadosProduto.InformaMensagem objMensagem
    
    bRetorno = objDadosProduto.RecuperarDados(Produto_ID)
    
    objForm.lblProduto.Caption = objDadosProduto.Descricao
    objForm.Caption = objDadosProduto.Sigla & " - Identificar DSLAM"
    
    'Carrega os Cabecalho
    If (Familia_Produto = "I") Then
        
       If (Produto_ID = GbIDProdutoJAH) Then
            'JAH
            Call CarregaCabecalho_JAH(objForm.UsrJAHCabecalho, objBlfCabecalho)
            
       End If
       
    End If
    
    If (DataFinalizacao <> "" And PermiteExecucao <> "S") Or _
        paPermissoes(4, 1) = 0 Or Not flUsuarioDoCF Then
    
       Call DesabilitaObjetosTela
        
    End If
    
    If (Sigla_Acao = "DES" Or Sigla_Acao = "CAN") Then
            
        Call DesabilitaObjetosTela
    
        objForm.CmdSalvar.Enabled = True
        objForm.CmdLimpar.Enabled = False
            
    End If
    
    'Preenche o combo tipo de solução
    bRetorno = ExibirListaTipoSolucao(objForm.CboTipoSolucao)
    
    If (Not bRetorno) Then
    
       objMensagem.ExibeMensagem

       Unload objForm
       Set objForm = Nothing
       
       GoTo Finalizar
    
    Else
        ''Alteração Davi 21/09
        If objForm.CboTipoSolucao.ListCount > 0 Then
        
            'Call SelecionaCombo(objForm.CboTipoSolucao, 0)
                               
            objForm.CboTipoSolucao.ListIndex = 0
            
        End If
    End If
    
    ''Alteração Davi 21/09
    objForm.optSolADE_Nao = True
    
    objForm.SpdDataAtivacaoAcesso.Text = Format(Date, "DD/MM/YYYY")
    
    'Caso exista Servico Anterior e a fase ainda não foi iniciada,
    'Obtem os dados da OTS anterior
    
    If (Servico_ID_Anterior <> 0) Then
        'Não resgata os dados da OTS anterior quando se altera o
        'endereço de instalação. IP VPN
        
            bResgataAnterior = True
                
        If (objdadosServico.ID_ACAO = GbIDAcaoAlteracao Or _
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
        
       'Recupera os dados de Acesso DSLAM e popula os objetos da tela
'       If (Not objDadosAcessoDSLAM.FaseExistente(Fase_ID) And _
                bResgataAnterior) Then
           
        'Valida se apresenta os dados da OTS atual ou da anterior
        objDadosAcesso.CaracteristicaAcesso = "P"
        
        If (Not objDadosAcesso.FaseExistente(Fase_ID) And _
                bResgataAnterior) Then
           
           IdOTS = Servico_ID_Anterior
           
           'GbCodigoFaseAcesso - ID da fase na tabela Tipo_Fase
           IDFase = objdadosServico.ObtemIDFase(Servico_ID_Anterior, _
                        GbCodigoFaseAcesso)
    
           If (objMensagem.TipoMsg <> "") Then
           
              GoTo ErrorHandler
              
           End If
       
       Else
       
           IdOTS = Servico_ID
           IDFase = Fase_ID
           
       End If
       
    Else
       
       IdOTS = Servico_ID
       IDFase = Fase_ID
    
    End If
        
    'Recupera os dados de Acesso Contigente(Embratel)  - Contigente
    bRetorno = objDadosAcesso.RecuperarDados("C", _
                    IdOTS, _
                    IDFase, _
                    objMensagem)
    
    If (objMensagem.TipoMsg <> "") Then

       objMensagem.ExibeMensagem

       Unload objForm
       
       Set objForm = Nothing

       GoTo Finalizar

    End If
    
    lIdAcessoContigente = objDadosAcesso.iD
    
        
    'Recupera os dados de Acesso Contigente(Embratel)  - Principal
    bRetorno = objDadosAcesso.RecuperarDados("P", _
                    IdOTS, _
                    IDFase, _
                    objMensagem)
    
    If (objMensagem.TipoMsg <> "") Then

       objMensagem.ExibeMensagem

       Unload objForm
       Set objForm = Nothing

       GoTo Finalizar

    End If
    
    lIdAcessoPrincipal = objDadosAcesso.iD
    
    'Atribui objetos necessários para gravação do acesso
    objDadosAcesso.objPropriedadeEBT = objPropriedadeEBT
    
    '''
    'Instancia as classes necessárias para gravação dos tipos de acesso
    Set objdadosTipoADE = New clsGenTipoAcessoADE
    
    objdadosTipoADE.InformaBanco frmgenMDI_SGAPlus.objBanco
    
    'Recupera dados do acesso Tipo ADE
    bRetorno = objdadosTipoADE.RecuperarDados(lIdAcessoPrincipal, _
                    objMensagem)

    If (objMensagem.TipoMsg <> "") Then

       objMensagem.ExibeMensagem

       Unload objForm
       Set objForm = Nothing

       GoTo Finalizar

    End If

    objDadosAcesso.objPropriedadeEBT.objTipoAcesso = objdadosTipoADE

    'Recupera dados do acesso ebt
    bRetorno = objDadosAcesso.objPropriedadeEBT.RecuperarDados(lIdAcessoPrincipal, _
                                        objMensagem)

    If (objMensagem.TipoMsg <> "") Then

       objMensagem.ExibeMensagem

       Unload objForm
       Set objForm = Nothing

       GoTo Finalizar

    End If
    
    bRetorno = objDadosAcessoDSLAM.RecuperarDados(lIdAcessoPrincipal)
    
    If objMensagem.TipoMsg <> "" Then
       
       objMensagem.ExibeMensagem
       
       Set objForm = Nothing
       
       GoTo Finalizar
    
    End If

    '''
    If (Not bRetorno) Then
    
        Set cColecao = New Collection
        
        objDadosPendencia.ColecaoPendencia = cColecao
        
        objForm.FlControlaAlteracao = False
        
        TipoOperacao = "I"     'Indica que a fase está sendo executada pela 1a vez
        
        objForm.informa_OC Me
        
        objForm.Show
        
        GoTo Finalizar
    
    End If
    '''
    'Caso a fase não tenha sido iniciada e seja uma OTS de Desativacao ou Cancelamento,
    'Exibe os dados da OTS original
    
    TipoOperacao = "A"     'Indica que a fase já foi iniciada anteriormente
    
    lUsuarioGravadoID = (0 & objDadosAcesso.UsuarioAtualizacao_ID)
       
    objForm.informa_OC Me
    
    'Popula os campos da tela
    Call PopularCamposTela
    
    'Caso exista pendencia para a fase obtem os dados de pendencia e carrega a
    'propriedade do formulário com o objeto Pendencia
    
    If (objForm.cmdPendencia.Enabled = True _
            And Familia_Produto = "I") Then
           
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
    
    'Caso o Servico já esteja Finalizado, exibir os campos desabilitados
    If objdadosServico.DataRealizacao <> "" And Familia_Produto = "I" Then
       
       Call DesabilitaObjetosTela

    End If
    
    objForm.FlControlaAlteracao = False
    
    
    objForm.Show
    
    'Executa método para popular os campos do Controle Cabecalho
    CarregarForm = True

Finalizar:
    
    Set objdadosServico = Nothing
    Set objAutenticacao = Nothing
    Set objDadosFase = Nothing
    Set objDadosProduto = Nothing
    Set cColecao = Nothing
    Set objdadosCircuito = Nothing
    
    Exit Function
    
ErrorHandler:

    If Err.Number <> 0 Then
       objMensagem.TipoMsg = "E"
       objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                   "Houve um erro de execução no Método: " & vbCrLf & vbCrLf & "CarregarForm " & _
                   "da classe clsblfOC_TratarAcessoDSLAM"
    End If
       
    objMensagem.ExibeMensagem
       
    GoTo Finalizar
       
End Function
Public Function GravarDados(pbFlagDescarregaFormulario As Boolean, _
                            Optional Flag_Atualiza_Pendencia As Boolean) As Boolean

    'Esta fase gera dois regitros na contigência de acesso
    'um do tipo ADE e outro do tipo Cabo interno.
    
    'As informações destes acessos serão armazenadas na fase Ativação/Finalização
    
    Dim objdadosServico             As New clsGenServico
    Dim objDadosFase                As New clsgenFase
    Dim objWEBIP                    As clsInterOC_TrataWebIP
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
    
    'Dim lIdFaseCentroLocal          As Long
    'Dim lIDCentroFuncionalCL        As Long
    'Dim sCodigoCFCentroLocal        As String
    
    'Checa se o produto pode alocar vários CFs para a fase Tratar Centro Local
    'Dim objDadosDuracaoFase         As clsgenDuracaoFase
    
    On Error GoTo ErrorHandler
    
    FlagIniciouTransacao = False
    FlagTemPendenciaCliente = False
    
    GravarDados = False
    
    Screen.MousePointer = vbArrowHourglass
    
'    'Instancia as classes necessárias para gravação dos tipos de acesso
    Set objdadosTipoADE = New clsGenTipoAcessoADE
    'Set objDadosTipoCabo = New clsGenTipoAcessoCabo
    
    'Instância a classe de dados Acesso
    'Set objDadosAcesso = New clsGenAcesso
    
    'Seta o objeto mensagem da classe de dados
    'objDadosAcesso.InformaMensagem objMensagem
    'objDadosAcesso.InformaBanco frmgenMDI_SGAPlus.objBanco

    'Verfica se a OTS selecionada ainda encontra-se em andamento
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
    
    If (sSituacaoOTS = "CAN") Or _
            (sSituacaoOTS = "CR") Then
       
       objMensagem.TipoMsg = "A"
       objMensagem.Descricao = "Esta OTS foi paralisada pelo ATEN!"
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
                    GbCodigoFaseProgramarCPE, _
                    Familia_Produto)
    
    If (objMensagem.TipoMsg <> "") Then
       
       objMensagem.ExibeMensagem
       GoTo Finalizar
    
    End If
    
    'CALBAT - 21/09/2007
    '====================================================
    'If (numUltimaFaseRealizada > Sequencia_Acao) And _
            (PermiteExecucao <> "S") Then
            
    If (numUltimaFaseRealizada > Sequencia_Acao And _
            PermiteExecucao <> "S") And lProduto_ID <> GbIDProdutoVBL Then
            
    '====================================================
    
       objMensagem.TipoMsg = "A"
       objMensagem.Descricao = "Os dados desta fase não poderão ser alterados! Pois, a fase seguinte já foi iniciada."
       objMensagem.ExibeMensagem
       
       GravarDados = True
       
       GoTo Finalizar
    
    End If

    'Checa se o produto pode alocar vários CFs para a fase Tratar Centro Local
'    Set objDadosDuracaoFase = New clsgenDuracaoFase
'
'    objDadosDuracaoFase.InformaBanco frmgenMDI_SGAPlus.objBanco
'    objDadosDuracaoFase.InformaMensagem objMensagem
'
    'Carrega as propriedades da classe Contigência do acesso
    
    Call CarregarPropriedadesContigenciaAcesso
    
    'Carrega as propriedades da classe Propriedade Embratel
    Call CarregarPropriedadesAcessoEBT
    
    'Carrega as propriedades da classe Acesso DSLAM
    Call CarregarPropriedadesDSLAM
    
    'Carrega as propriedades da classe Acesso ADE
    Call CarregarPropriedadesADE
    
    bRetorno = frmgenMDI_SGAPlus.objBanco.Iniciar_Transacao()
    
    '=================================================================
    'CALLBAT - 21/05/2012 - BSoD Net Light
    '=================================================================
      objDadosFase.Usuario_ID = GbID_Usuario
    '=================================================================
    '=================================================================
        
    If (Not bRetorno) Then
        
        GoTo ErrorHandler
    
    End If
    
    FlagIniciouTransacao = True
    
    'Indica ser o acesso principal.
    'Na fase de ativação é verificado se há a necessidade de gravar
    'acesso Contigente.
    objDadosAcesso.CaracteristicaAcesso = "P"
    
    'Atualiza dados de Acesso(Contigência do acesso) - Acesso Principal
    If objDadosAcesso.FaseExistente(Fase_ID) Then
       
       TipoOperacao = "A"
       
       'Verifica e atualiza os dados do Acesso - Contigente
       If (lIdAcessoContigente > 0) Then
           
           objDadosAcesso.CaracteristicaAcesso = "C"
           objDadosAcesso.iD = lIdAcessoContigente
    
           bRetorno = objDadosAcesso.AtualizarDados(Fase_ID)
        
           'Trata erro
            If (objMensagem.TipoMsg <> "") Then
               
               GoTo ErrorHandler
               
           End If
           
        End If
       
       'Atualiza os dados do Acesso - Principal
       objDadosAcesso.CaracteristicaAcesso = "P"
       objDadosAcesso.iD = lIdAcessoPrincipal
       
       bRetorno = objDadosAcesso.AtualizarDados(Fase_ID)
    
       'Trata erro
        If (objMensagem.TipoMsg <> "") Then
           
           GoTo ErrorHandler
           
       End If
    
    
    Else
       
       TipoOperacao = "I"
       
       'Inclui os dados do Acesso Principal
       objDadosAcesso.CaracteristicaAcesso = "P"
       
       bRetorno = objDadosAcesso.IncluirDados(Fase_ID)
        
       'Trata erro
        If (objMensagem.TipoMsg <> "") Then
           
           GoTo ErrorHandler
           
       End If
       
       'Atualiza os dados do Acesso Contigente
'       objDadosAcesso.CaracteristicaAcesso = "C"
'
'       bRetorno = objDadosAcesso.IncluirDados(Fase_ID)
'
'       'Trata erro
'        If (objMensagem.TipoMsg <> "") Then
'
'           GoTo ErrorHandler
'
'       End If
       
       'Recupera o ID do Acesso principal
       objDadosAcesso.iD = objDadosAcesso.ObtemIDContingencia(Fase_ID, _
                "P")
    
        If (objMensagem.TipoMsg <> "") Then
           
           GoTo ErrorHandler
           
        End If
        
        lIdAcessoPrincipal = objDadosAcesso.iD
    
'       'Recupera o ID contigente do Acesso
'       objDadosAcesso.ID = objDadosAcesso.ObtemIDContingencia(Fase_ID, _
'                "C")
'
'        If (objMensagem.TipoMsg <> "") Then
'
'           GoTo ErrorHandler
'
'        End If
'
'        lIdAcessoContigente = objDadosAcesso.ID
'
    End If
    
    'Trata erro
    If (objMensagem.TipoMsg <> "") Then
       
       GoTo ErrorHandler
       
    End If
    
    'Grava os dados da Propriedade Embratel - Principal
    objDadosAcesso.CaracteristicaAcesso = "P"
    objDadosAcesso.iD = lIdAcessoPrincipal
    
    
    
'    'Acesso Tipo ADE
    objDadosAcesso.objPropriedadeEBT.objTipoAcesso = objdadosTipoADE
    objDadosAcesso.objPropriedadeEBT.TipoAcesso = "A"
    
    bRetorno = objOCTratarAcessoEBT.GravaDadosPropriedadeEmbratel(objDadosAcesso)

    'Captura o objeto mensagem
    Set objMensagem = objOCTratarAcessoEBT.objMensagem
    
    If (objMensagem.TipoMsg <> "") Then

       GoTo ErrorHandler

    End If

    'Grava os dados da Propriedade Embratel - Contingente
'    objDadosAcesso.CaracteristicaAcesso = "C"
'    objDadosAcesso.ID = lIdAcessoContigente
'
'    'Acesso Cabo interno
'    objDadosAcesso.objPropriedadeEBT.objTipoAcesso = objDadosTipoCabo
'    objDadosAcesso.objPropriedadeEBT.TipoAcesso = "C"
'
'    bRetorno = objOCTratarAcessoEBT.GravaDadosPropriedadeEmbratel(objDadosAcesso)
'
'    If (objMensagem.TipoMsg <> "") Then
'
'       GoTo ErrorHandler
'
'    End If
    
    'Grava os dados do acesso DSLAM - Uma única vez(Acesso Principal)
    If (objDadosAcessoDSLAM.FaseExistente(lIdAcessoPrincipal)) Then
       
       bRetorno = objDadosAcessoDSLAM.AtualizarDados(lIdAcessoPrincipal)
    
    Else
       
       bRetorno = objDadosAcessoDSLAM.IncluirDados(lIdAcessoPrincipal, _
                            objMensagem)
    
    End If
   
    If (Not bRetorno) Then
       
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
            bRetorno = objdadosServico.DevolveDesignacaoOTS(objdadosServico.ID_SERVICO, _
                                    Produto_ID, _
                                    True)
            If (Not bRetorno) Then
                GoTo ErrorHandler
            End If
            
            sDesignacaoAtual = ("" & objdadosServico.DesignacaoCircuito)
            
            'Recupera os dados de Circuito da OTS que está sendo alterada
            bRetorno = objdadosServico.DevolveDesignacaoOTS(objdadosServico.ID_SERVICO_ANTERIOR, _
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
''''              lIdFaseCentroLocal = objDadosFase.ObtemIDFase(Servico_ID, _
''''                            GbCodigoFaseCentroLocal)
''''
''''              If (objMensagem.TipoMsg <> "") Then
''''                GoTo ErrorHandler
''''              End If
            
              'Checa se o produto pode alocar vários CFs para a fase Tratar Centro Local
'''              If (objDadosDuracaoFase.ChecaPermissaoMultiplosCF(Produto_ID, _
'''                                                       GbCodigoFaseCentroLocal)) Then
'''
'''                If objDadosFase.ObtemSequenciaFaseFinalizada(Servico_ID, _
'''                                                   GbCodigoFaseCentroLocal) = 0 And _
'''                                                   (Familia_Produto = "I" And _
'''                                                   lIdFaseCentroLocal > 0) Then
'''
'''                  'Carrega a tela Alocação de Centro Funcional e atualiza os CFs alocados
'''                  Set ocAcessoEmbratel = New clsGenOC_TratarAcessoEmbratel
'''
'''                  'Informa o ID Produto
'''                  ocAcessoEmbratel.Produto_ID = Produto_ID
'''
'''                  'Informa a Controladora Selecionar OTS
'''                  Call ocAcessoEmbratel.InformaOC_SelecionarOTS(objOCSelecionarOTS)
'''
'''                  bretorno = ocAcessoEmbratel.CarregarFormCFsDaFase(lIdFaseCentroLocal)
'''
'''                  If (Not bretorno) Then
'''                      GoTo ErrorHandler
'''                  End If
'''
'''                End If 'ObtemSequenciaFaseFinalizada
'''
'''              End If 'Checa multiplos
'''
          ElseIf (Familia_Produto = "D") Then
    
              'Calcula o período para a proxima fase
              bRetorno = objDadosFase.CalculaPeriodoProximaFase(Sequencia_Acao, Servico_ID, Familia_Produto)
        
              If (objMensagem.TipoMsg <> "") Then
              
                 GoTo ErrorHandler
                 
              End If
          
          End If    'Familia_Produto <> "I"
         
                
          'CALBAT - 21/09/2007
          '=========================================================================
          'If (Sigla_Acao <> "CAN" And _
                Sigla_Acao <> "DES" And _
                Familia_Produto = "I") Then
            
         If (Sigla_Acao <> "CAN" And _
                Sigla_Acao <> "DES" And _
                Familia_Produto = "I") And lProduto_ID <> GbIDProdutoVBL Then
            
                If (objWEBIP Is Nothing) Then
                    Set objWEBIP = New clsInterOC_TrataWebIP
                End If
          '=========================================================================
                
                
                objWEBIP.Grupo_ID = lGrupo_ID
                
''''
''''                bRetorno = objWEBIP.GravaWEBIPIntegracao(Servico_ID, _
''''                            Fase_ID, _
''''                            "C", _
''''                            Produto_ID, _
''''                            objMensagem, _
''''                            Sigla_Acao)
''''
''''                If (objMensagem.TipoMsg <> "") Then
''''                  GoTo ErrorHandler
''''                End If
                
          End If
          
          'Quando for uma OTS de Cancelamento ou Desativacao, verifica se
          'as fases que estão sendo executadas em paralelo já foram finalizadas
          'caso Todas as outras fases já estejam finalizadas, o sistema deverá finalizar
          'o Servico.
          If (Sigla_Acao = "CAN" Or Sigla_Acao = "DES") And _
                (Familia_Produto = "I") Then

             bRetorno = VerificaFinalizacaoDesativacao(objDadosFase, _
                                                       objdadosServico, _
                                                       Servico_ID, Fase_ID, _
                                                       Sigla_Acao, _
                                                       objMensagem)
          End If
       
          'Preenche o periodo realizado indicando que a fase foi finalizada
          bRetorno = objDadosFase.FinalizaFase(Fase_ID)
    
          If (objMensagem.TipoMsg <> "") Then
             
             GoTo ErrorHandler
             
          End If
       
       End If
       
       'Produtos de Dados
'       If (Familia_Produto = "D") Then
              
            'Obtem a sequencia_acao da última fase existente para a OTS
'            lSeqUltimaFaseOTS = objDadosFase.ObtemSequenciaUltimaFase(Servico_ID, Familia_Produto)
'
'            If (objMensagem.TipoMsg <> "") Then
'
'                objMensagem.ExibeMensagem
'
'                GoTo Finalizar
'
'            End If
'
'            If (Sequencia_Acao = lSeqUltimaFaseOTS) Then
'
''                Set objDadosAnalisarOE = New clsDNI_DadosAnalisarOE
''
''                objDadosAnalisarOE.InformaMensagem objMensagem
''                objDadosAnalisarOE.InformaBanco frmgenMDI_SGAPlus.objBanco
''
''                sDesignacaoSOE = Trim(objForm.UsrDNICabecalhoAcesso.CNLPontaA) & "/" & _
''                                  Trim(objForm.UsrDNICabecalhoAcesso.CNLPontaB) & "/" & _
''                                  Trim(objForm.UsrDNICabecalhoAcesso.CodigoNP) & "/"
'
''                If Produto_ID = GbIDProdutoDiginet Then
''
''                    sDesignacaoSOE = sDesignacaoSOE & _
''                    Format(objForm.UsrDNICabecalhoAcesso.SequencialDesig, "0000")
''
''                Else
''
''                    sDesignacaoSOE = sDesignacaoSOE & _
''                            Format(objForm.UsrDNICabecalhoAcesso.SequencialDesig, "0000000")
''
''                End If
'
''                bRetorno = objDadosAnalisarOE.EncerraWorkFlow(Servico_ID, _
''                                Servico_ID_Anterior, _
''                                AnoOE, NumeroOE, NumeroItemOE, _
''                                "02", _
''                                GbUsuario, _
''                                sDesignacaoSOE, _
''                                Sigla_Acao, _
''                                objdadosServico)
''
''                Set objDadosAnalisarOE = Nothing
'
''                If (Sigla_Acao = "DES" Or _
''                        Sigla_Acao = "CAN") Then
''
''                    objdadosServico.FinalizaServico (Servico_ID)
''
''                End If
'
'            End If    'Sequencia_Acao
            
'       End If   'Produtos de Dados
       
       FlagFaseFinalizadaOK = True
       
       pbFlagDescarregaFormulario = True
       
    Else
       'Caso a fase já tenha sido finalizada e uma pendencia tenha sido incluida,
       'A fase deverá ser considerada como Em Andamento
       
       objDadosFase.ID_SERVICO = Servico_ID
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
    
    'Atualiza o label Login
    lUsuarioGravadoID = lUsuarioID
    
    Call PopularCampoLogin

    bRetorno = frmgenMDI_SGAPlus.objBanco.Finalizar_Transacao
    
    
    If (Not bRetorno) Then
        
        GoTo ErrorHandler
    
    End If
    
    objForm.FlControlaAlteracao = False
    bFlagControlaAlteracao = False
    
    GravarDados = True
           
    objMensagem.TipoMsg = "A"
    

    If (FlagFaseFinalizadaOK = True And _
                Not FlagTemPendenciaEmAberto) Then

        objMensagem.Descricao = "Os Dados de Acesso foram " & _
                               "atualizados com Sucesso!"
       
        FlagDescarregaForm = True
                    
                    
    'Controla a existência de pendência dsa informações do pedido ADE
'    ElseIf (Not FlagFaseFinalizadaOK And _
'                FlagTemPendenciaADE) Then
'
'       objMensagem.Descricao = "Os Dados de Acesso foram atualizados com Sucesso. " & vbCrLf & _
'                               " Porém a fase continua em Andamento," & vbCrLf & _
'                               " pois campos obrigatórios do Pedido do Acesso ADE não foram preenchidos."
'''
    ElseIf (Not FlagFaseFinalizadaOK And _
                Not FlagTemPendenciaEmAberto) Then

       objMensagem.Descricao = "Os Dados de Acesso foram atualizados com Sucesso. " & vbCrLf & _
                               " Porém a fase continua em Andamento," & vbCrLf & _
                               " pois campos obrigatórios não foram preenchidos."
    
    
    ElseIf (Not FlagFaseFinalizadaOK And _
                FlagTemPendenciaEmAberto) Then

       objMensagem.Descricao = "Os Dados de Acesso foram atualizados com Sucesso. " & vbCrLf & _
                               " Porém a fase continua em andamento, devido a existência de pendência!"
    
    End If
    
    objMensagem.ExibeMensagem
    
    GoTo Finalizar
    
Finalizar:
    
    Set objdadosServico = Nothing
    Set objDadosFase = Nothing
    Set objWEBIP = Nothing
    
    Screen.MousePointer = vbDefault
    
    Exit Function
    
ErrorHandler:

    ' Cancelar Transação
    If (FlagIniciouTransacao) Then
       
       bRetorno = frmgenMDI_SGAPlus.objBanco.Cancelar_Transacao
    
    End If
    
    'If (Err.Number <> 0) Then
    If (objMensagem.Descricao <> "") Then
       
       objMensagem.TipoMsg = "E"
       objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                   "Houve um erro de execução no Método: " & vbCrLf & vbCrLf & "GravarDados " & _
                   "da classe clsblfOC_TratarAcessoDSLAM"
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

    On Error GoTo ErrorHandler
    
    VerificaFinalizacaoOK = False

    Dim vMensagemListaMsg   As Variant
    Dim iAux                As Integer
    
    
    'Exibe os campos obrigatórios não preenchidos
    iAux = 2
    ReDim vMensagemListaMsg(iAux)
    vMensagemListaMsg(0) = "Os campos abaixo foram preenchidos incorretamente ou não foram preenchidos:"
    vMensagemListaMsg(1) = ""
    
    If (Sigla_Acao = "DES" Or Sigla_Acao = "CAN") Then
    
       VerificaFinalizacaoOK = True
       
    Else
    
        VerificaFinalizacaoOK = True
        
        
        'Data da ativação do acesso
        If (objDadosAcesso.DataFinalizacao = "") Then
            
            VerificaFinalizacaoOK = False
            
            iAux = iAux + 1
            
            ReDim Preserve vMensagemListaMsg(iAux)
            vMensagemListaMsg(iAux - 1) = "A Data da Ativação do acesso " & _
                                          " é obrigatória."
        End If
        
        'Nome do equipamento
        'If (objDadosAcessoDSLAM.NomeDSLAM = "") Then
            
        '    VerificaFinalizacaoOK = False
            
        '    iAux = iAux + 1
            
        '    ReDim Preserve vMensagemListaMsg(iAux)
        '    vMensagemListaMsg(iAux - 1) = "O Nome do equipamento " & _
                                          " é obrigatório."
        'End If
        
        'Identificação do prédio
        'If (objDadosAcessoDSLAM.IdentPredioDSLAM = "") Then
            
        '    VerificaFinalizacaoOK = False
            
        '    iAux = iAux + 1
            
        '    ReDim Preserve vMensagemListaMsg(iAux)
        '    vMensagemListaMsg(iAux - 1) = "O Nome da identificação do prédio " & _
                                          " é obrigatório."
        'End If
        
        'Os campos de Pedido do Acesso ADE são obrigatórios se o
        'usuário opta por gravar o pedido ADE
''''        If (objDadosAcessoDSLAM.SolicitarADE = "S") Then
        
''''            'Identificação do acesso Embratel
''''            If (objPropriedadeEBT.IdentificadorAcesso = "") Then
''''
''''                VerificaFinalizacaoOK = False
''''
''''                iAux = iAux + 1
''''
''''                ReDim Preserve vMensagemListaMsg(iAux)
''''                vMensagemListaMsg(iAux - 1) = "A Identificação do acesso Embratel " & _
''''                                              " é obrigatória."
''''            End If
''''
''''            'Designação do acesso
''''            If (objPropriedadeEBT.DesignacaoAcesso = "") Then
''''
''''                VerificaFinalizacaoOK = False
''''
''''                iAux = iAux + 1
''''
''''                ReDim Preserve vMensagemListaMsg(iAux)
''''                vMensagemListaMsg(iAux - 1) = "A Designação do acesso " & _
''''                                              " é obrigatória."
''''            End If
''''
''''            'Estacao de Entrega do Acesso - Configuracao
''''            If (objPropriedadeEBT.EstacaoEntregaAcessoConfiguracao = "") Then
''''
''''                VerificaFinalizacaoOK = False
''''
''''                iAux = iAux + 1
''''
''''                ReDim Preserve vMensagemListaMsg(iAux)
''''                vMensagemListaMsg(iAux - 1) = "A Estacao de Entrega do Acesso - Configuracao " & _
''''                                              " é obrigatória."
''''
''''            End If
''''
''''            'Estacao de Entrega do Acesso - Embratel
''''            If (objPropriedadeEBT.EstacaoEntregaAcesso = "") Then
''''
''''                VerificaFinalizacaoOK = False
''''
''''                iAux = iAux + 1
''''
''''                ReDim Preserve vMensagemListaMsg(iAux)
''''                vMensagemListaMsg(iAux - 1) = "A Estacao de Entrega do Acesso - Embratel " & _
''''                                              " é obrigatória."
''''
''''            End If
''''
'''            'Data do pedido do acesso
'''            If (objDadosAcesso.DataPedidoAcesso = "") Then
'''
'''                VerificaFinalizacaoOK = False
'''
'''                iAux = iAux + 1
'''
'''                ReDim Preserve vMensagemListaMsg(iAux)
'''                vMensagemListaMsg(iAux - 1) = "A Data do pedido do acesso " & _
'''                                              " é obrigatória."
'''            End If
'''
'''        End If
        
    End If
        
    'Valida a exibição da relação de campos obrigatórios não informados
    If (UBound(vMensagemListaMsg) > 2) Then
        
        Call ExibeListaMensagem(vMensagemListaMsg)
        
    End If

Finalizar:
    
    Exit Function
    
ErrorHandler:
       
    If Err.Number <> 0 Then
       objMensagem.TipoMsg = "E"
       objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                   "Houve um erro de execução no Método: " & vbCrLf & vbCrLf & "VerificaFinalizacaoOK " & _
                   "da classe clsgenOC_TratarAcessoDSLAM"
    End If

    GoTo Finalizar

End Function
Private Function DesabilitaObjetosTela()

    With objForm
        
        .MskNomeDSLAM.Enabled = False
        .TxtIdentPredioDSLAM.Enabled = False
        '.TxtNumeroOTSTX.Enabled = False
        '.TxtItemOTSTx.Enabled = False
        .SpdDataAtivacaoAcesso.Enabled = False
        
        .CmdSalvar.Enabled = False
        .CmdLimpar.Enabled = False
        
        .CboTipoSolucao.Enabled = False
        
        .TxtEMailGic.Enabled = False
        .TxtRIT.Enabled = False
        
'        .TxtPredioEBT_ADE.Enabled = False
'        .TxtNumeroCabo_ADE.Enabled = False
'        .TxtNumeroPar_ADE.Enabled = False
        
        .optSolADE_Sim.Enabled = False
        .optSolADE_Nao.Enabled = False
        
        'Pedido do Acesso ADE
'        .TxtIdentificadorAcessoEBT.Enabled = False
'        .TxtDesignacaoAcessoEBT.Enabled = False
'        .TxtEstacaoEntregaEbtConf.Enabled = False
'        .TxtEstacaoEntregaEbt.Enabled = False
'        .SpdDataPedidoAcesso.Enabled = False

    End With

End Function

Public Function LimpaTela()
    
    With objForm
        
        .MskNomeDSLAM.Mask = ""
        .MskNomeDSLAM.Text = ""
        .MskNomeDSLAM.Mask = "CCCC-CCC-CCC"
        
        .TxtIdentPredioDSLAM.Text = ""
        
        .TxtEMailGic.Text = ""
        .TxtRIT.Text = ""

        .SpdDataAtivacaoAcesso.Col = 1
        .SpdDataAtivacaoAcesso.Row = 1
        .SpdDataAtivacaoAcesso.Text = ""
        
        .optSolADE_Sim.Value = False
        .optSolADE_Nao.Value = False
        
        .CboTipoSolucao.ListIndex = -1
        
        'Pedido do Acesso ADE
'        .TxtIdentificadorAcessoEBT.Text = ""
'        .TxtDesignacaoAcessoEBT.Text = ""
'        .TxtEstacaoEntregaEbtConf.Text = ""
'        .TxtEstacaoEntregaEbt.Text = ""
'        .SpdDataPedidoAcesso.Col = 1
'        .SpdDataPedidoAcesso.Row = 1
'        .SpdDataPedidoAcesso.Text = ""
        
 End With

    
End Function

'Public Function LimpaTab_ADE()
'    'Limpa apenas os campos da Tab DSLAM para o caso do usuário desmarcar esta opções
'
'    With objForm
'
'        'Pedido do Acesso ADE
''        .TxtIdentificadorAcessoEBT.Text = ""
''        .TxtDesignacaoAcessoEBT.Text = ""
''        .TxtEstacaoEntregaEbtConf.Text = ""
''        .TxtEstacaoEntregaEbt.Text = ""
''        .SpdDataPedidoAcesso.Col = 1
''        .SpdDataPedidoAcesso.Row = 1
''        .SpdDataPedidoAcesso.Text = ""
'
' End With
'
'End Function


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
                            "InformaOCSelecionarOTS da classe clsblfOC_TratarAcessoDSLAM"
    
    GoTo Finalizar

End Function



'Private Function GravaDadosPropriedadeEmbratel(objDadosAcesso As clsGenAcesso) As Boolean
'
'    Dim bRetorno              As Boolean
'    Dim lContingenciaAcesso   As Long
'    Dim FlJaexiste            As Boolean
'    'Dim objPropriedadeTerc    As clsGenPropriedadeTerceiros
'
'On Error GoTo ErrorHandler
'
'    lContingenciaAcesso = objDadosAcesso.ID
'
'    GravaDadosPropriedadeEmbratel = False
'
'    objDadosAcesso.objPropriedadeEBT.InformaBanco frmgenMDI_SGAPlus.objBanco
'
'    FlJaexiste = objDadosAcesso.objPropriedadeEBT.FaseExistente(lContingenciaAcesso)
'
'    If objMensagem.TipoMsg <> "" Then
'
'       GoTo ErrorHandler
'
'    End If
'
'    'Caso o Tipo de Acesso seja o mesmo, realiza somente uma Atualiza
'    'Caso contrario, exclui o tipo anterior e incluir o novo
'
'    If objDadosAcesso.objPropriedadeEBT.TipoAcesso <> TipoAcessoAnterior Then
'
'        'Exclui o tipo de Acesso anterior
'        bRetorno = ExcluiTipoAcessoEBT(TipoAcessoAnterior, lContingenciaAcesso, objMensagem)
'
'    End If
'
'    If Not FlJaexiste Then  'Inclusao
'
'       bRetorno = objDadosAcesso.objPropriedadeEBT.IncluirDados(lContingenciaAcesso, objMensagem)
'
'       If objMensagem.TipoMsg <> "" Then
'
'          GoTo ErrorHandler
'
'       End If
'
'       'Caso o usuário tenha alterado a propriedade de Acesso para a mesma caracteristica
'       'de acesso, exclui os dados da propriedade anterior
'
''       If PropriedadeAcessoAnterior <> "" And _
''          CaracteristicaAcessoAnterior = objDadosAcesso.CaracteristicaAcesso Then
''
''          Set objPropriedadeTerc = New clsGenPropriedadeTerceiros
''
''          'Exclui dados de Propriedade Terceiros
''
''         bRetorno = ExcluiTipoAcessoTerceiros(TipoAcessoAnterior, lContingenciaAcesso, objMensagem)
''
''         If objMensagem.TipoMsg <> "" Then
''
''             GoTo ErrorHandler
''
''         End If
''
''         objPropriedadeTerc.InformaBanco frmgenMDI_SGAPlus.objBanco
''
''          bRetorno = objPropriedadeTerc.ExcluirDados(lContingenciaAcesso, objMensagem)
''
''          If objMensagem.TipoMsg <> "" Then
''
''             GoTo ErrorHandler
''
''         End If
''
''         TipoAcessoAnterior = ""
''
''       End If
'
'    Else
'
'       bRetorno = objDadosAcesso.objPropriedadeEBT.AtualizarDados(lContingenciaAcesso, objMensagem)
'
'    End If
'
'    If Not bRetorno Then
'
'        GoTo ErrorHandler
'
'    End If
'
'    'Grava dados de Tipo de Acesso
'    objDadosAcesso.objPropriedadeEBT.objTipoAcesso.InformaBanco frmgenMDI_SGAPlus.objBanco
'
'    If Not objDadosAcesso.objPropriedadeEBT.objTipoAcesso.VerificaExistencia(lContingenciaAcesso, objMensagem) Then
'
'       bRetorno = objDadosAcesso.objPropriedadeEBT.objTipoAcesso.IncluirDados( _
'                                                        lContingenciaAcesso, objMensagem)
'    Else
'
'       bRetorno = objDadosAcesso.objPropriedadeEBT.objTipoAcesso.AtualizarDados( _
'                                                        lContingenciaAcesso, objMensagem)
'
'    End If
'
'    If objMensagem.TipoMsg <> "" Then
'
'        GoTo ErrorHandler
'
'    End If
'
'    GravaDadosPropriedadeEmbratel = True
'
'Finalizar:
'
'    On Error GoTo 0
'
'    Exit Function
'
'ErrorHandler:
'
'    If objMensagem.TipoMsg = "" Then
'
'       objMensagem.TipoMsg = "E"
'       objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
'                   "Houve um erro de execução no Método: " & vbCrLf & vbCrLf & "GravarDadosPropriedadeEmbratel " & _
'                   "da classe clsGenOC_TratarAcessoDSLAM"
'
'    End If
'
'    GoTo Finalizar
'
'End Function
Public Function AtualizarListaOTS() As Boolean

    objOCSelecionarOTS.ExibirListaOTS
        
End Function

Private Sub Class_Terminate()

    Set objDadosPendencia = Nothing
    Set objMensagem = Nothing
    Set objForm = Nothing
    
    Set objDadosAcesso = Nothing
    Set objPropriedadeEBT = Nothing
    Set objDadosAcessoDSLAM = Nothing

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

On Error GoTo ErrorHandler

    With objForm
    
        If (Len(objDadosAcessoDSLAM.NomeDSLAM) = 10) Then
        
            .MskNomeDSLAM.Text = Left(objDadosAcessoDSLAM.NomeDSLAM, 4) & "-" & _
                                        Mid(objDadosAcessoDSLAM.NomeDSLAM, 5, 3) & "-" & _
                                        Right(objDadosAcessoDSLAM.NomeDSLAM, 3)
        
        ElseIf (Len(objDadosAcessoDSLAM.NomeDSLAM) = 9) Then
        
            .MskNomeDSLAM.Text = Left(objDadosAcessoDSLAM.NomeDSLAM, 3) & " " _
                                        & "-" & _
                                        Mid(objDadosAcessoDSLAM.NomeDSLAM, 5, 3) & "-" & _
                                        Right(objDadosAcessoDSLAM.NomeDSLAM, 3)
        Else
            .MskNomeDSLAM.Text = "____-___-___"
        End If
        
        .TxtIdentPredioDSLAM.Text = ("" & objDadosAcessoDSLAM.IdentPredioDSLAM)

        .SpdDataAtivacaoAcesso.Col = 1
        .SpdDataAtivacaoAcesso.Row = 1
        .SpdDataAtivacaoAcesso.Text = (IIf(IsDate(objPropriedadeEBT.DataAtivacaoAcesso), Format(objPropriedadeEBT.DataAtivacaoAcesso, "dd/mm/yyyy"), ""))
        
        If Val(objDadosAcessoDSLAM.IdTipoSolucao) > 0 Then
           
           Call SelecionaCombo(objForm.CboTipoSolucao, _
                               CLng(objDadosAcessoDSLAM.IdTipoSolucao))
        End If


        .TxtRIT.Text = objDadosAcesso.RIT
        .TxtEMailGic.Text = objDadosAcesso.eMailUsuario
        
        If (objDadosAcessoDSLAM.SolicitarADE = "S") Then

            .optSolADE_Sim.Value = True
            'Habilita o Tab Acesso DSLAM
            '.tabDadosAcesso.TabEnabled(1) = True

        ElseIf (objDadosAcessoDSLAM.SolicitarADE = "N") Then

            .optSolADE_Nao.Value = True
            'Desabilita o Tab Acesso DSLAM
            '.tabDadosAcesso.TabEnabled(1) = False

        End If
        
        'Pedido do Acesso ADE
'        .TxtIdentificadorAcessoEBT.Text = objPropriedadeEBT.IdentificadorAcesso
'        .TxtDesignacaoAcessoEBT.Text = objPropriedadeEBT.DesignacaoAcesso
'        .TxtEstacaoEntregaEbtConf.Text = objPropriedadeEBT.EstacaoEntregaAcessoConfiguracao
'        .TxtEstacaoEntregaEbt.Text = objPropriedadeEBT.EstacaoEntregaAcesso
'        .SpdDataPedidoAcesso.Col = 1
'        .SpdDataPedidoAcesso.Row = 1
'        .SpdDataPedidoAcesso.Text = objDadosAcesso.DataPedidoAcesso
        
        'Obtém o login
        Call PopularCampoLogin
        
    End With
    
Finalizar:

    Exit Function
    
ErrorHandler:

    If Err.Number <> 0 Then
       objMensagem.TipoMsg = "E"
       objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                   "Houve um erro de execução no Método: " & vbCrLf & vbCrLf & "PopularCamposTela " & _
                   "da classe clsblfOC_TratarAcessoDSLAM"
    End If
       
    objMensagem.ExibeMensagem
       
    GoTo Finalizar
    
End Function
Private Function PopularCampoLogin()

On Error GoTo ErrorHandler

    Dim objDadosUsuario         As clsgenUsuario

    Set objDadosUsuario = New clsgenUsuario

    objDadosUsuario.InformaBanco frmgenMDI_SGAPlus.objBanco
    objDadosUsuario.InformaMensagem objMensagem
    
    bRetorno = objDadosUsuario.RecuperarDados(lUsuarioGravadoID)
    
    If Not bRetorno Then

        GoTo Finalizar

    End If
    
    
    With objForm
        
        .LblNomeUsuario.Caption = objDadosUsuario.Login
        
    End With
    
Finalizar:

    Set objDadosUsuario = Nothing
    
    Exit Function
    
ErrorHandler:

    If Err.Number <> 0 Then
       objMensagem.TipoMsg = "E"
       objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                   "Houve um erro de execução no Método: " & vbCrLf & vbCrLf & "PopularCamposTela " & _
                   "da classe clsblfOC_TratarAcessoDSLAM"
    End If
       
    objMensagem.ExibeMensagem
       
    GoTo Finalizar
    
End Function

Private Function ObtemCD_Velocidade() As Boolean

On Error GoTo ErrorHandler

    Dim objDadosUsuario         As clsgenUsuario
    Dim objDadosAcessoInterface As clsblfDadosAcesso
    Dim sCodigoProdutoSOE       As String
    Dim sTabela_Interface       As String

'    Set objDadosUsuario = New clsgenUsuario
    ObtemCD_Velocidade = False
    
    If (Produto_ID = GbIDProdutoJAH) Then

        Set objDadosAcessoInterface = New clsblfDadosAcesso

        objDadosAcessoInterface.InformaBanco frmgenMDI_SGAPlus.objBanco
        objDadosAcessoInterface.InformaMensagem objMensagem

        'Recupera Dados de Acesso
        sTabela_Interface = ""

        If (Produto_ID = GbIDProdutoJAH) Then

            sTabela_Interface = "ATEN_ACESS_FACIL"
            sCodigoProdutoSOE = GbCodigo_Produto_JAH

        End If

        bRetorno = objDadosAcessoInterface.RecuperarDados(ContratoID, _
                        NrLocalInsta, _
                        sCodigoProdutoSOE, _
                        sTabela_Interface)

        If (objMensagem.TipoMsg <> "") Then

            GoTo ErrorHandler

        End If

        objDadosAcesso.Cd_Veloc = objDadosAcessoInterface.CdVelocidadeAcesso

    End If
    
    ObtemCD_Velocidade = True
    
Finalizar:

    Set objDadosUsuario = Nothing
    Set objDadosAcessoInterface = Nothing
    
    Exit Function
    
ErrorHandler:

    If Err.Number <> 0 Then
       objMensagem.TipoMsg = "E"
       objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                   "Houve um erro de execução no Método: " & vbCrLf & vbCrLf & "PopularCamposTela " & _
                   "da classe clsblfOC_TratarAcessoDSLAM"
    End If
       
    objMensagem.ExibeMensagem
       
    GoTo Finalizar
    
End Function

Private Function CarregarPropriedadesDSLAM()

    With objForm
    
        objDadosAcessoDSLAM.NomeDSLAM = TrataPlick("" & .MskNomeDSLAM.ClipText)
        objDadosAcessoDSLAM.IdentPredioDSLAM = TrataPlick("" & .TxtIdentPredioDSLAM.Text)
        
        objDadosAcessoDSLAM.DataAtivacaoAcesso = objPropriedadeEBT.DataAtivacaoAcesso
        
        'Tipo da solução elecionada
        If (objForm.CboTipoSolucao.ListIndex > -1) Then
        
            objDadosAcessoDSLAM.IdTipoSolucao = (0 & .CboTipoSolucao.ItemData(.CboTipoSolucao.ListIndex))
            
        Else
        
            objDadosAcessoDSLAM.IdTipoSolucao = 0
        
        End If
        
        If (.optSolADE_Sim.Value = True) Then

            objDadosAcessoDSLAM.SolicitarADE = "S"

        ElseIf (.optSolADE_Nao.Value = True) Then

            objDadosAcessoDSLAM.SolicitarADE = "N"

        End If
        
    End With
    
End Function

Private Function CarregarPropriedadesADE()

    With objForm

        objdadosTipoADE.iD = objDadosAcesso.objPropriedadeEBT.iD

'        objdadosTipoADE.NumeroCaboEBT = Val(.TxtNumeroCabo_ADE)
'        objdadosTipoADE.NumeroParEBT = TrataPlick("" & .TxtNumeroPar_ADE)
'        objdadosTipoADE.PredioEBT = TrataPlick("" & .TxtPredioEBT_ADE)
'
    End With

End Function
Private Function CarregarPropriedadesContigenciaAcesso()
    'Carrega as propriedades necessárias para gravação da tabela Contigência de acesso
    
    On Error GoTo ErrorHandler
    
    'Indica não ter segurança de acesso.
    
    'Na fase de ativação é verificado se há a necessidade de gravar
    'alterar para Segurança de acesso sim.
    objDadosAcesso.SegurancaAcesso = "N"
    
    'Embratel
    objDadosAcesso.PropridedadeAcesso = "E"
    
    'Grava apenas o acesso principal
    objDadosAcesso.CaracteristicaAcesso = "P"
    
    objDadosAcesso.UsuarioAtualizacao_ID = frmgenMDI_SGAPlus.objAutenticacao.IDUsuario
    lUsuarioGravadoID = objDadosAcesso.UsuarioAtualizacao_ID
    
    objDadosAcesso.Cd_Veloc = objBlfCabecalho.Cod_Velocidade
    
    Call PopularCampoLogin
    
    With objForm
        'Data de ativação do acesso
        .SpdDataAtivacaoAcesso.Col = 1
        .SpdDataAtivacaoAcesso.Row = 1
        objDadosAcesso.DataFinalizacao = (IIf(IsDate(.SpdDataAtivacaoAcesso.Text), Format(.SpdDataAtivacaoAcesso.Text, "dd/mm/yyyy"), ""))
        
        'Obtém o código de velocidade do circuito
        bRetorno = ObtemCD_Velocidade
        
        If (Not bRetorno) Then
        
            GoTo Finalizar
        
        End If
        
        objDadosAcesso.RIT = .TxtRIT.Text
        objDadosAcesso.eMailUsuario = .TxtEMailGic.Text
        
        'Pedido do Acesso ADE
        '.SpdDataPedidoAcesso.Row = 1
        '.SpdDataPedidoAcesso.Col = 1
        'objDadosAcesso.DataPedidoAcesso = .SpdDataPedidoAcesso.Text
    
    End With
    
Finalizar:
    
    Exit Function
    
ErrorHandler:

    If Err.Number <> 0 Then
       objMensagem.TipoMsg = "E"
       objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                   "Houve um erro de execução no Método: " & vbCrLf & vbCrLf & "CarregarPropriedadesContigenciaAcesso " & _
                   "da classe clsblfOC_TratarAcessoDSLAM"
    End If
       
    objMensagem.ExibeMensagem
       
    GoTo Finalizar

End Function
Private Function CarregarPropriedadesAcessoEBT()
    'Carrega as propriedades necessárias para gravação da tabela Propriedade do Acesso
    
    On Error GoTo ErrorHandler
    
    With objForm
        'Inicial grava como ADE. Será atualizado na fase Ativação/Finalização
        objPropriedadeEBT.TipoAcesso = "A"
        
        'Data de ativação do acesso
        .SpdDataAtivacaoAcesso.Col = 1
        .SpdDataAtivacaoAcesso.Row = 1
        objPropriedadeEBT.DataAtivacaoAcesso = (IIf(IsDate(.SpdDataAtivacaoAcesso.Text), Format(.SpdDataAtivacaoAcesso.Text, "dd/mm/yyyy"), ""))
        
'        'ADE
'        objPropriedadeEBT.objTipoAcesso.PredioEBT = "" & TrataPlick(Trim(.TxtPredioEBT_ADE.Text))
'        objdadosTipoADE.NumeroCaboEBT = Val(.TxtNumeroCabo_ADE.Text)
'        objdadosTipoADE.NumeroParEBT = "" & TrataPlick(Trim(.TxtNumeroPar_ADE.Text))
        
        'Pedido do Acesso ADE
'        objPropriedadeEBT.IdentificadorAcesso = .TxtIdentificadorAcessoEBT.Text
'        objPropriedadeEBT.DesignacaoAcesso = .TxtDesignacaoAcessoEBT.Text
'        objPropriedadeEBT.EstacaoEntregaAcessoConfiguracao = .TxtEstacaoEntregaEbtConf.Text
'        objPropriedadeEBT.EstacaoEntregaAcesso = .TxtEstacaoEntregaEbt.Text
        
    End With
    
Finalizar:
    
    Exit Function
    
ErrorHandler:

    If Err.Number <> 0 Then
       objMensagem.TipoMsg = "E"
       objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                   "Houve um erro de execução no Método: " & vbCrLf & vbCrLf & "CarregarPropriedadesAcessoEBT " & _
                   "da classe clsblfOC_TratarAcessoDSLAM"
    End If
       
    objMensagem.ExibeMensagem
       
    GoTo Finalizar

End Function

'
'    objFormAlocacao.informa_OC Me
'
'    objFormAlocacao.Show vbModal
'
'    CarregarFormAlocacaoCF = CentroFuncionalCL_ID
'
'    Set objFormAlocacao = Nothing
'
'End Function
'
Public Function CamposPreenchidosOK() As Boolean

    Dim objDadosFase    As clsgenFase
    Dim bRetorno        As Boolean
    Dim sDataFim        As String
    
    CamposPreenchidosOK = False

    objForm.SpdDataAtivacaoAcesso.Row = 1
    objForm.SpdDataAtivacaoAcesso.Col = 1
    

    'Nenhum campo preenchido
    If Trim(objForm.MskNomeDSLAM.ClipText) = "" And _
       Trim(objForm.TxtIdentPredioDSLAM.Text) = "" And _
       Trim(objForm.SpdDataAtivacaoAcesso.Text) = "" And _
       (objForm.TxtRIT.Text = "") And _
       (objForm.TxtEMailGic.Text = "") And _
       (objForm.optSolADE_Sim.Enabled = False And _
            objForm.optSolADE_Nao.Enabled = False) Then
       
       objMensagem.TipoMsg = "A"
       objMensagem.Descricao = "Nenhuma Informação foi Preenchida!"
       objMensagem.ExibeMensagem
       
       objForm.MskNomeDSLAM.SetFocus
       
       Exit Function
    
    End If

    ''Campos obrigatórios não preenchidos
    'If Trim(objForm.MskNomeDSLAM.ClipText) = "" Or _
    '   Trim(objForm.TxtIdentPredioDSLAM.Text) = "" Then
    '
    '   objMensagem.TipoMsg = "A"
    '   objMensagem.Descricao = "Nome do Equipamento de Acesso e Identificação " & _
    '                        vbCrLf & "do Prédio são informações obrigatórias para gravação da fase!"
    '
    '   objMensagem.ExibeMensagem
    '
    '   objForm.MskNomeDSLAM.SetFocus
    '
    '   Exit Function
    '
    'End If


    Set objDadosFase = New clsgenFase
    
    objDadosFase.InformaBanco frmgenMDI_SGAPlus.objBanco
    objDadosFase.InformaMensagem objMensagem
    
    sDataFim = objDadosFase.ObtemDataFimRealFaseAnterior(Servico_ID, _
                                                         Sequencia_Acao, _
                                                         Familia_Produto)
    
    Set objDadosFase = Nothing
    

    If (objForm.SpdDataAtivacaoAcesso.Text) <> "" And _
        Not IsDate(objForm.SpdDataAtivacaoAcesso.Text) Then
        
        objMensagem.TipoMsg = "A"
        objMensagem.Descricao = "Data de Ativação do Acesso Cliente não é uma data válida "
        objMensagem.ExibeMensagem
        
        objForm.SpdDataAtivacaoAcesso.SetFocus
          
       Exit Function
    
    End If

    If Trim(objForm.SpdDataAtivacaoAcesso.Text) <> "" And sDataFim <> "" And _
       (Sigla_Acao <> "DES" And Sigla_Acao <> "CAN") Then
    
        If CDate(sDataFim) > CDate(objForm.SpdDataAtivacaoAcesso.Text) Then
        
            objMensagem.TipoMsg = "A"
            objMensagem.Descricao = "Data de Ativação do Acesso não pode ser menor que " & _
                                    "a data de Finalização da fase anterior!"
            
            objMensagem.ExibeMensagem
            
            objForm.SpdDataAtivacaoAcesso.SetFocus
            
            Exit Function
        
        End If
        
    End If
    
    CamposPreenchidosOK = True
    
End Function


Private Function ExibirListaTipoSolucao(ByRef cbo As ComboBox, _
                                            Optional ByVal plIdTipoSolucao As Long) As Boolean
On Error GoTo ErrorHandler:

    Dim bRetorno As Boolean
    
    Dim aListaTipoSolucao() As Variant
    Dim objTipoSolucao  As clsgenTipoSolucao
    
    ExibirListaTipoSolucao = False
    
    ' Classe de centro funcional
    Set objTipoSolucao = New clsgenTipoSolucao
    
    ' Informar o objeto do banco para as classes
    '
    ' Classe de centro funcional
    objTipoSolucao.InformaBanco frmgenMDI_SGAPlus.objBanco
    
    ' Informar o objeto de mensagem para as classes
    '
    ' Classe de centro funcional
    objTipoSolucao.InformaMensagem objMensagem
    
    ' Limpa a combo
    cbo.Clear
    
    ' Executar o método de montagem da lista de centro funcional na classe de analisar order entry
    bRetorno = objTipoSolucao.ListarDados(aListaTipoSolucao)
    
    ' Se ocorreu algum erro na DLL exibi-lo
    If Not bRetorno Then

        GoTo Finalizar

    End If
    
    'Preenche o combo
    PreencheCombo cbo, aListaTipoSolucao()
    
      
    ExibirListaTipoSolucao = True
    
Finalizar:

    Screen.MousePointer = vbDefault
    
    If objMensagem.TipoMsg <> "" Then
    
        objMensagem.ExibeMensagem
        objMensagem.TipoMsg = ""
        
    End If
    
    Set objTipoSolucao = Nothing
    
    On Error GoTo 0
    
    Exit Function
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: ExibirListaTipoSolucao " & _
                            "da classe clsgenOC_TratarAcessoDSLAM"
    
    GoTo Finalizar

End Function
