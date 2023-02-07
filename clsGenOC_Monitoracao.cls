VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGenOC_Monitoracao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private objMensagem         As clsgenMensagem

Dim lCentroFuncional_ID As Long
Dim VetorServico        As Variant
Dim VetorFull           As Variant
Dim lTipoAcao           As Long
Dim lTipoServico        As Long
Dim lProduto_ID         As Long
Dim lGrupo_ID           As Long
Dim sAcaoString         As String
Dim sFamilia_Produto    As String

Public Property Get Familia_Produto() As String

    Familia_Produto = sFamilia_Produto

End Property

Public Property Let Familia_Produto(ByVal sNewValue As String)

    sFamilia_Produto = sNewValue
    
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

'
Public Function CarregarFormPortMonitoracao() As Boolean

    Dim objServico              As clsgenOC_SelecionarOTS
    Dim objFormSelecionarOTS    As frmPortMonitoracao
    
    Dim aListaCentroFuncional()
    
    'RN_SGAD_21148YBJ
    Dim vTodos As String
    vTodos = "[Todos]"
    
    CarregarFormPortMonitoracao = False
    
    ' Instânciar a classe de controle de serviço
    Set objServico = New clsgenOC_SelecionarOTS

    Set objFormSelecionarOTS = frmPortMonitoracao
    
    Load objFormSelecionarOTS
    
    ' Define período
    objFormSelecionarOTS.dtpDataIni.Value = Format(Now - 10, "dd/mm/yyyy")
    objFormSelecionarOTS.dtpDataFim.Value = Format(Now, "dd/mm/yyyy")
    
'==========================================================================
'==========================================================================
        
    ' Executar o método de exibição dos centros funcionais
    bRetorno = objServico.ExibirListaCentroFuncional(objFormSelecionarOTS.cboCentroFuncional)
    
    If Not bRetorno Then
       
       Set objServico = Nothing
       Exit Function
    
    End If
    
    'RN_SGAD_21148YBJ
    If (GbFlagRN_SGAD_21148YBJ) Then
        
        objFormSelecionarOTS.cboCentroFuncional.AddItem vTodos
        objFormSelecionarOTS.cboCentroFuncional.ItemData(objFormSelecionarOTS.cboCentroFuncional.NewIndex) = 0
        
    End If
    
'    'Carrega combo de Acao
'    Call ExibirListaAcao(objFormSelecionarOTS.cboAcao)
'
'    bRetorno = ExibirListaProduto(objFormSelecionarOTS.CboServico, Familia_Produto)
'
'    objFormSelecionarOTS.CboServico.AddItem "", 0
'
'    If Not bRetorno Then
'
'        Unload objFormSelecionarOTS
'
'        Exit Function
'
'    End If
    
    Set objServico = Nothing

    bRetorno = ExibirListaProduto(objFormSelecionarOTS.CboServico, Familia_Produto, , 1)
    
    objFormSelecionarOTS.CboServico.AddItem "", 0
    
    If Not bRetorno Then
        
        Unload objFormSelecionarOTS
        
        Exit Function
        
    End If
    
    
'    'Carrega combo de Acao
'    Call ExibirListaAcao(objFormSelecionarOTS.cboAcao)
'
'    bRetorno = ExibirListaProduto(objFormSelecionarOTS.CboServico, Familia_Produto)
'
'    objFormSelecionarOTS.CboServico.AddItem "", 0
'
'    If Not bRetorno Then
'
'        Unload objFormSelecionarOTS
'
'        Exit Function
'
'    End If
'
'    Set objServico = Nothing
'
'    bRetorno = ExibirListaProduto(objFormSelecionarOTS.CboServico, Familia_Produto)
'
'    objFormSelecionarOTS.CboServico.AddItem "", 0
'
'    If Not bRetorno Then
'
'        Unload objFormSelecionarOTS
'
'        Exit Function
'
'    End If
'
'    If Familia_Produto = "I" Then
'
'        objFormSelecionarOTS.Caption = "Produtos de Internet/IP - Monitoração de OTS em Andamento"
'
'    Else
'
'        objFormSelecionarOTS.Caption = "Produtos de Dados - Monitoração de OTS em Andamento"
'
'    End If
'
'    objFormSelecionarOTS.InformaOC Me
'==========================================================================
'==========================================================================

    objFormSelecionarOTS.Show


End Function

Public Function CarregarFormMonitoracao() As Boolean

    Dim objServico              As clsgenOC_SelecionarOTS
    Dim objFormSelecionarOTS    As frmgenMonitoracaoOTS
    
    Dim aListaCentroFuncional()
    
    'RN_SGAD_21148YBJ
    Dim vTodos As String
    vTodos = "[Todos]"
    
    CarregarFormMonitoracao = False
    
    ' Instânciar a classe de controle de serviço
    Set objServico = New clsgenOC_SelecionarOTS
    
    '#################
    'adeil RN-SGAD-14971ZKH 28/10/2003
    'se o form passado por parâmetro já estiver aberto, fecha-lo.
    '################# inicio
    If Familia_Produto = "I" Then
        If Verifica_Janela_Aberta("frmgenMonitoracaoOTS", "Produtos de Internet - Monitoração de OTS em Andamento") Then
            Exit Function
        End If
    Else
        If Verifica_Janela_Aberta("frmgenMonitoracaoOTS", "Produtos de Dados - Monitoração de OTS em Andamento") Then
            Exit Function
        End If
    End If
    'fim
    '#################
    
    Set objFormSelecionarOTS = New frmgenMonitoracaoOTS
    
    Load objFormSelecionarOTS
    
    '*******************************************************************
    '
    ' Responsável : Accenture – Paulo Roberto
    ' Data        : 15/04/2008
    ' BO          : CH789557
    ' Observação  :
    '
    '*******************************************************************
    objFormSelecionarOTS.spdCircuitos.Row = 0
    
    objFormSelecionarOTS.spdCircuitos.Col = 2
    objFormSelecionarOTS.spdCircuitos.Text = "Número OTS"
    
    objFormSelecionarOTS.spdCircuitos.Col = 3
    objFormSelecionarOTS.spdCircuitos.Text = "Designação"
        
    ' Executar o método de exibição dos centros funcionais
    bRetorno = objServico.ExibirListaCentroFuncional(objFormSelecionarOTS.cboCentroFuncional)
    
    If Not bRetorno Then
       
       Set objServico = Nothing
       Exit Function
    
    End If
    
    'RN_SGAD_21148YBJ
    If (GbFlagRN_SGAD_21148YBJ) Then
        
        objFormSelecionarOTS.cboCentroFuncional.AddItem vTodos
        objFormSelecionarOTS.cboCentroFuncional.ItemData(objFormSelecionarOTS.cboCentroFuncional.NewIndex) = 0
        
    End If
    
    
    'Carrega combo de Acao
    Call ExibirListaAcao(objFormSelecionarOTS.cboAcao)
    
    bRetorno = ExibirListaProduto(objFormSelecionarOTS.CboServico, Familia_Produto)
    
    objFormSelecionarOTS.CboServico.AddItem "", 0
    
    If Not bRetorno Then
        
        Unload objFormSelecionarOTS
        
        Exit Function
        
    End If
    
    Set objServico = Nothing

    bRetorno = ExibirListaProduto(objFormSelecionarOTS.CboServico, Familia_Produto)
    
    objFormSelecionarOTS.CboServico.AddItem "", 0
    
    If Not bRetorno Then
        
        Unload objFormSelecionarOTS
        
        Exit Function
        
    End If
    
    If Familia_Produto = "I" Then
    
        objFormSelecionarOTS.Caption = "Produtos de Internet/IP - Monitoração de OTS em Andamento"
        
    Else
    
        objFormSelecionarOTS.Caption = "Produtos de Dados - Monitoração de OTS em Andamento"
    
    End If
    
    objFormSelecionarOTS.InformaOC Me
    
    objFormSelecionarOTS.Show
    
End Function

Private Sub Class_Initialize()

Set objMsg = New clsgenMensagem

End Sub

Private Sub Class_Terminate()

Set objMsg = Nothing

End Sub

Public Function MontarListaOTS(Frm As frmgenMonitoracaoOTS) As Boolean

    Dim objServico                   As clsGenServico
    Dim objDadosPendencia            As New clsblf_Pendencia
    Dim aListaMonitoracao()          As Variant
    Dim vCabecalho()                 As Variant
    '''Dim VetorOrdenado()              As Variant
    Dim sCabecalhoColunas            As String
    Dim sMeta                        As String
    Dim JaExiste                     As Boolean
    Dim VarAuxCol0                   As String
    Dim VarAuxCol1                   As Integer
    
    Dim iLinha, iLinhaII, iAux, ColAux, QtColunas, PosIni, PosFim, ColFull  As Integer
    
    Dim RowAux, CentroFuncional_ID   As Long
    Dim strNomeFase, sNomeFase       As String
    Dim objMensagem                  As clsgenMensagem
    Dim sFiltroAlteracao             As String
    Dim iCol                         As Integer
    Dim lIDFase                      As Long
    
    'Indice para loop
    Dim i, J                         As Integer
    Dim vSequencia                   As Integer
    Dim vTipoFase                    As String
    Dim iProdutoId                   As Integer
    Dim objProduto                   As clsGenProduto
    Dim sOrdena             As Long   'Se valor da variavel  psOrdena = 1 é pra ordernar por Nome, se for 0 é pra ordenar por OTS
    
    MontarListaOTS = False
    
    ' Instânciar a classe de controle de serviço
    Set objServico = New clsGenServico
    
    Set objMensagem = New clsgenMensagem
    
    objDadosPendencia.InformaBanco frmgenMDI_SGAPlus.objBanco
    objDadosPendencia.InformaMensagem objMensagem
    
    objServico.InformaBanco frmgenMDI_SGAPlus.objBanco
    objServico.InformaMensagem objMensagem
    
    If Frm.cboCentroFuncional.ListIndex > -1 Then
        CentroFuncional_ID = Frm.cboCentroFuncional.ItemData(Frm.cboCentroFuncional.ListIndex)
    Else
        CentroFuncional_ID = 0
    End If
    
    '*******************************************************************
    '
    ' Responsável : Accenture – Gladson M França
    ' Data        : 15/04/2008
    ' BO          : CH789557
    ' Observação  : Verificação se a nova opção de ordenação (por Designação) está selecionada e variável sOrdena = 2 em caso positivo.
    '
    '*******************************************************************
    If Frm.optDesignacao.Value = True Then
        sOrdena = 2
    '*******************************************************************
    ElseIf Frm.optNome.Value = True Then
        sOrdena = 1
    Else
        sOrdena = 0
    End If
    
    ' Executar o método que obtem as OTS para monitoracao
    If TipoAcao = GbIDAcaoAlteracao Or TipoAcao = GbIDAcaoCorrecaoAtivacao Or _
       TipoAcao = GbIDAcaoCorrecaoAlteracao Then
       
        sFiltroAlteracao = "ALT"
        
    Else
    
        sFiltroAlteracao = ""
    
    End If
    
    '*******************************************************************
    '
    ' Responsável : Accenture – Gladson M França
    ' Data        : 15/04/2008
    ' BO          : CH789557
    ' Observação  : Os valores dos campos dos novos filtros são aqui salvos nas novas propriedades "CNL", "Sigla", "Numero_Designacao", e "Cliente" da classe clsGenServico.
    '
    '*******************************************************************
    objServico.CNL = TrataPlick(Frm.TxtCnl)
    objServico.Sigla = TrataPlick(Frm.txtSigla)
    objServico.Numero_Designacao = TrataPlick(Frm.txtNumeroDesignacao)
    objServico.Cliente = TrataPlick(Frm.txtCliente)
    '*******************************************************************
    
    If Trim(Frm.TxtCnl.Text) = "" And Trim(Frm.txtSigla.Text) = "" And Trim(Frm.txtNumeroDesignacao.Text) = "" Then

        'AGRC - SO-TIEMP-49824BGM - 15/11/2017
        'If Frm.CboServico.ItemData(Frm.CboServico.ListIndex) = GbIDProdutoNetPhone Then
        If Frm.CboServico.ItemData(Frm.CboServico.ListIndex) = GbIDProdutoNetPhone Or Frm.CboServico.ItemData(Frm.CboServico.ListIndex) = GbIDProdutoClaroFone Then
            
            bRetorno = objServico.ObtemDadosMonitoracao_NetPhone(TipoServico, _
                                                        TipoAcao, _
                                                        CentroFuncional_ID, _
                                                        aListaMonitoracao(), _
                                                        sCabecalhoColunas, _
                                                        sFiltroAlteracao)
        Else
            
            bRetorno = objServico.ObtemDadosMonitoracao(TipoServico, _
                                                        TipoAcao, _
                                                        CentroFuncional_ID, _
                                                        sOrdena, _
                                                        Frm.CboServico.ItemData(Frm.CboServico.ListIndex), _
                                                        aListaMonitoracao(), _
                                                        sCabecalhoColunas, _
                                                        sFiltroAlteracao)
        End If
        
    Else
       If Trim(Frm.txtSigla.Text) <> "" Then
        
            Set objProduto = New clsGenProduto
            
            objProduto.InformaBanco frmgenMDI_SGAPlus.objBanco
            objProduto.InformaMensagem objMensagem
    
            If Frm.txtSigla.Text = "IP" Then
               
               Call objProduto.RecuperarDados(Frm.CboServico.ItemData(Frm.CboServico.ListIndex))
               
               iProdutoId = objProduto.ID_Produto
               TipoServico = iProdutoId
            Else
               
               iProdutoId = objProduto.ObtemID(Frm.txtSigla.Text)
               TipoServico = iProdutoId
            
            End If
        
        End If
        
        bRetorno = objServico.ObtemDadosMonitoracao(TipoServico, _
                                                    TipoAcao, _
                                                    CentroFuncional_ID, _
                                                    sOrdena, _
                                                    iProdutoId, _
                                                    aListaMonitoracao(), _
                                                    sCabecalhoColunas, _
                                                    sFiltroAlteracao)
    
    End If
    
    If Not bRetorno Then
       
       If objMensagem.TipoMsg <> "" Then
          
          Frm.spdCircuitos.MaxRows = 0
          
          objMensagem.ExibeMensagem
          
          Set objServico = Nothing
          
          Exit Function
       
       End If
       
       objMensagem.TipoMsg = "A"
       objMensagem.Descricao = "Não foi possível carregar a lista de OTS's!"
       objMensagem.ExibeMensagem
       
       Set objServico = Nothing
       
       Exit Function
    
    End If
    
    Set objServico = Nothing
    
    sAcaoString = Frm.cboAcao.List(Frm.cboAcao.ListIndex)
    
    ColAux = 1
    RowAux = 0
    
    Frm.spdCircuitos.MaxRows = 0
    '*********************************************************************
    ' Responsável: Accenture – Soraia
    ' Data       : 16/04/2008
    ' BO         : CH789557
    ' Observação : Alteração do número máximo de colunas
    '*********************************************************************
    Frm.spdCircuitos.MaxCols = 17
    
    QtColunas = 1
    
    PosIni = 2
    PosFim = 0
    
    sNomeFase = ""
    
    'Monta Cabecalho
    For iAux = 2 To QtColunas
        
        With Frm.spdCircuitos
             .Row = 0
             .Col = iAux
             .ColHidden = True
             .ColWidth(iAux) = 14.5
             .Text = ""
        End With
        
    Next iAux
    
    'Monta Cabecalho
    
    iCol = 0
    iLinhaII = 0
    ReDim vCabecalho(2, 1)
    
    For iAux = 0 To UBound(aListaMonitoracao, 2)
    
        JaExiste = False
        
        For iLinha = 0 To UBound(vCabecalho, 2)
        
            If vCabecalho(0, iLinha) = aListaMonitoracao(3, iAux) Then
                
                JaExiste = True
                Exit For
            
            End If
            
        Next iLinha
        
        If Not JaExiste Then
           
           vCabecalho(0, iLinhaII) = aListaMonitoracao(3, iAux)
           vCabecalho(1, iLinhaII) = aListaMonitoracao(12, iAux)
           
           ReDim Preserve vCabecalho(2, iLinhaII + 1)
           
           iLinhaII = iLinhaII + 1
        
        End If
        
    Next iAux
    
    'IP VPN
    'Lógica de ordenação do vetor Cabeçalho reescrita
    For i = 0 To UBound(vCabecalho, 2) - 1
    
        For J = 0 To UBound(vCabecalho, 2) - 1
    
            If vCabecalho(1, i) < vCabecalho(1, J) Then
                
                vSequencia = vCabecalho(1, i)
                vTipoFase = vCabecalho(0, i)
                
                vCabecalho(1, i) = vCabecalho(1, J)
                vCabecalho(0, i) = vCabecalho(0, J)
                
                vCabecalho(1, J) = vSequencia
                vCabecalho(0, J) = vTipoFase
    
            End If
        Next J
    
    Next i

    
    'Ordena o Vetor de Cabecalho
'''    iLinha = 0
'''    iLinhaII = 0
'''    iAux = 0
'''    PosIni = 0
'''
'''    ReDim VetorOrdenado(2, 1)
    
''    VarAuxCol0 = vCabecalho(0, 0)
''    VarAuxCol1 = vCabecalho(1, 0)
    
'''    Do While True
'''
'''        For iLinhaII = PosIni To UBound(vCabecalho, 2)
'''
'''            If vCabecalho(0, iLinhaII) <> "" Then
'''
'''                If vCabecalho(1, iLinhaII) < VarAuxCol1 Then
'''
'''                   VarAuxCol0 = vCabecalho(0, iLinhaII)
'''                   VarAuxCol1 = vCabecalho(1, iLinhaII)
'''
'''                   vCabecalho(0, iLinhaII) = ""
'''                   vCabecalho(1, iLinhaII) = ""
'''
'''                   PosIni = PosIni - 1
'''
'''                End If
'''
'''            End If
'''
'''        Next iLinhaII
'''
'''
'''        ReDim Preserve VetorOrdenado(2, iLinha + 1)
'''
'''        VetorOrdenado(0, iLinha) = VarAuxCol0
'''        VetorOrdenado(1, iLinha) = VarAuxCol1
'''
'''        iLinha = iLinha + 1
'''        PosIni = PosIni + 1
'''
'''        If vCabecalho(0, PosIni) = "" Then
'''
'''            Exit Do
'''
'''        End If
'''
'''        VarAuxCol0 = vCabecalho(0, PosIni)
'''        VarAuxCol1 = vCabecalho(1, PosIni)
'''
'''    Loop
    
    'Monta o detalhe do Grid
    ReDim VetorServico(12, 1)
    
    ReDim VetorFull(6, 1)
    
    ColAux = 1
    ColFull = 0
    iLinha = 0
    
    For iAux = 0 To UBound(aListaMonitoracao, 2)
        
        If IsEmpty(aListaMonitoracao(0, iAux)) Then
        
           Exit For
        
        End If
        
        With Frm.spdCircuitos
           
           If aListaMonitoracao(0, iAux) <> "****" Then
             
              ColFull = 0
              RowAux = RowAux + 1
             
              .MaxRows = RowAux
              .Row = RowAux
              
               '*********************************************************************
               ' Responsável: Accenture – Soraia
               ' Data       : 16/04/2008
               ' BO         : CH789557
               ' Observação : Inclusão do texto separadamente nas três colunas
               '*********************************************************************
             
              .Col = 1
              .Text = aListaMonitoracao(2, iAux)
              
              .Col = 2
              .Text = aListaMonitoracao(23, iAux)
              
              .Col = 3
              .Text = aListaMonitoracao(24, iAux)
                      
             
              'Armazena o ID do Servico para cada OTs da Lista
              ReDim Preserve VetorServico(12, RowAux)
             
              'Servico_ID
              VetorServico(0, RowAux) = aListaMonitoracao(7, iAux)
              
              'Servico_Anterior_ID
              VetorServico(1, RowAux) = aListaMonitoracao(8, iAux)
              
              'Fase_ID
              VetorServico(2, RowAux) = aListaMonitoracao(9, iAux)
              
              'Acao_ID
              VetorServico(3, RowAux) = aListaMonitoracao(10, iAux)
              
              'Produto_ID
              VetorServico(4, RowAux) = aListaMonitoracao(11, iAux)
              
              'Sequencia_Acao
              VetorServico(5, RowAux) = aListaMonitoracao(12, iAux)
              
              'Sigla_Acao
              VetorServico(6, RowAux) = aListaMonitoracao(13, iAux)
              
              'Descricao - Status
              VetorServico(7, RowAux) = aListaMonitoracao(14, iAux)
              
              'Data Finalizacao
              VetorServico(8, RowAux) = aListaMonitoracao(15, iAux)
              
              'Contrato ID
              VetorServico(9, RowAux) = aListaMonitoracao(16, iAux)
              
              'Status Pendencia
              VetorServico(10, RowAux) = aListaMonitoracao(17, iAux)
              
              'Codigo Tipo Fase
              VetorServico(11, RowAux) = aListaMonitoracao(18, iAux)
           
           End If
           
           ColFull = ColFull + 1
           
           ReDim Preserve VetorFull(6, iAux)
           
           VetorFull(0, iAux) = aListaMonitoracao(3, iAux)   'Titulo da Coluna
           VetorFull(1, iAux) = RowAux                       'Número da Linha
           VetorFull(2, iAux) = aListaMonitoracao(9, iAux)   'Fase_ID correspondente a
                                                             'Coluna
           VetorFull(3, iAux) = aListaMonitoracao(18, iAux)  'CODIGO TIPO FASE
           VetorFull(4, iAux) = aListaMonitoracao(7, iAux)   'Servico_ID
           VetorFull(5, iAux) = aListaMonitoracao(19, iAux)  'Data fim Previsto
           
           'Subsitui o vetor VetorOrdenado pelo vCabecalho a partir deste ponto do
           'código
           For iLinha = 0 To UBound(vCabecalho, 2)
           
                 If aListaMonitoracao(3, iAux) = vCabecalho(0, iLinha) Then
                    '*********************************************************************
                    ' Responsável: Accenture – Paulo
                    ' Data       : 16/04/2008
                    ' BO         : CH789557
                    ' Observação : Inclusão do título no cabeçalho do Grid
                    '*********************************************************************
                   ColAux = iLinha + 3
                   .Row = 0
                                       
                   .Col = 2
                   .Text = "Número OTS"
                   .Col = 3
                   .Text = "Designação"
                                
                   .Col = ColAux + 1
                   .Text = aListaMonitoracao(3, iAux)
                   
                   Exit For
                   
                End If
                
           Next iLinha
           
           .Row = RowAux
           .Col = ColAux + 1
           
           .Text = IIf(aListaMonitoracao(6, iAux) = 0, "", aListaMonitoracao(6, iAux))

           Select Case aListaMonitoracao(5, iAux)   'Situacao
                  
                  Case "NI"                         'Não iniciada
                       
                       .BackColor = &H80000005
                  
                  Case "IF"                         'Inicio Futuro
                       
                       .BackColor = &HFFFF00
                  
                  Case "DM"                         'Dentro da Meta
                       
                       .BackColor = &HFF00&
                  
                  Case "FM"                         'Fora da Meta
                       
                       .BackColor = &H8080FF
                  
                  Case "CO"                         'Concluída
                       
                       .BackColor = &HFFC0C0
                       
                  Case "CR"                         'Correção
                  
                        .BackColor = &HC00000
                        
                  '===========================================================
                  'CALBAT - 20/10/2008
                  '
                  '===========================================================
                  Case "RL"                         'Correção
                  
                        .BackColor = &H8080&
                  '===========================================================
                  '===========================================================
                  
                  Case Else
                       
                       .BackColor = &H80000005
           
           End Select
           
           'Caso exista fase em aberto,
           If aListaMonitoracao(17, iAux) = "A" Then
            
              If aListaMonitoracao(20, iAux) = "C" Then   'Pendencia de Cliente
              
                 .BackColor = &H80C0FF
              
              Else  'Pendencia Embratel
                 
                 .BackColor = &H80FFFF
              
              End If
           
           End If
           
           
        End With
        
        ColAux = ColAux + 1
    
    Next iAux
    
    Frm.spdCircuitos.MaxRows = Frm.spdCircuitos.DataRowCnt
    
    '*********************************************************************
    ' Responsável: Accenture – Paulo
    ' Data       : 16/04/2008
    ' BO         : CH789557
    ' Observação : Inserção de mais duas colunas no Grid
    '*********************************************************************
    Frm.spdCircuitos.MaxCols = UBound(vCabecalho, 2) + 1 + 2
    
    For iAux = 2 To UBound(vCabecalho, 2) + 1
        
        With Frm.spdCircuitos
        
             .Row = 0
             .Col = iAux
             .ColWidth(iAux) = 14.5
             
        End With
        
    Next iAux
    
    If UBound(vCabecalho, 2) + 1 > 6 Then
       
       Frm.spdCircuitos.ScrollBars = ScrollBarsBoth
    
    End If
    
    Frm.spdCircuitos.Row = 1
    Frm.spdCircuitos.Col = 1
    Frm.spdCircuitos.Action = 1
    Frm.spdCircuitos.Action = 0
    
    '&H80000005& - Branco       - Nao iniciada  - NI
    '&H00FFFF00& - Azul claro   - Inicio Futuro
    '&H0000FF00& - Verde Claro  - Dentro da Meta
    '&H000000FF& - Red          - Fora da Meta
    '&H00FF0000& - Azul escuro  - Concluido
    
    Set objMensagem = Nothing

    ' Linhas
    MontarListaOTS = True

End Function

Public Function MontarListaOTSPort(Frm As frmPortMonitoracao) As Boolean

    Dim objServico                   As clsGenServico
    Dim objDadosPendencia            As New clsblf_Pendencia
    Dim aListaMonitoracao()          As Variant
    Dim vCabecalho()                 As Variant
    '''Dim VetorOrdenado()              As Variant
    Dim sCabecalhoColunas            As String
    Dim sMeta                        As String
    Dim JaExiste                     As Boolean
    Dim VarAuxCol0                   As String
    Dim VarAuxCol1                   As Integer
    
    Dim sTelefone                 As String
    Dim sTipoRelatorio          As Long
    
    Dim iLinha, iLinhaII, iAux, ColAux, QtColunas, PosIni, PosFim, ColFull  As Integer
    
    Dim sServico                      As Long
    Dim sServicoSeguinte        As Long
        
    Dim sTipoFaseId               As Long
    
    ' Instânciar a classe de controle de serviço
    Set objServico = New clsGenServico
    Set objMensagem = New clsgenMensagem
    
    objServico.InformaBanco frmgenMDI_SGAPlus.objBanco
    objServico.InformaMensagem objMensagem
    
    If Frm.CboServico.ListIndex <= 0 And Trim(Frm.CboServico.Text) = "" Then
    
        objMensagem.TipoMsg = "A"
        objMensagem.Descricao = " Favor Selecionar um Servico."
        
        objMensagem.ExibeMensagem
        
        Frm.CboServico.SetFocus
        
        Exit Function
    
    End If
    
    If Frm.cboCentroFuncional.ListIndex <= 0 And Trim(Frm.cboCentroFuncional.Text) = "" Then
    
        objMensagem.TipoMsg = "A"
        objMensagem.Descricao = " Favor Selecionar um Centro Funcional."
        
        objMensagem.ExibeMensagem
        
        Frm.cboCentroFuncional.SetFocus
        
        Exit Function
    
    End If
    
    
    If Trim(Frm.txtPortaCliente.Text) <> "" Then
        objServico.Cliente = TrataPlick(Frm.txtPortaCliente.Text)
    End If
    
    If Trim(Frm.TxtPortDDD.Text) <> "" Then
        objServico.DDD = Frm.TxtPortDDD.Text
    End If
    
    sTelefone = Replace(Frm.mskTelefone.Text, "_", "")
    
    sTelefone = Replace(sTelefone, "-", "")
    
    If sTelefone <> "" Then
        objServico.Telefone = sTelefone
    End If
    
    If Frm.txtPortNumBilhete.Text <> "" Then
        objServico.NumeroBilhete = TrataPlick(Frm.txtPortNumBilhete.Text)
    End If
    
    If Frm.OptPortDDD.Value = True Then
        sTipoRelatorio = 1
    ElseIf Frm.OptPortJanela.Value = True Then
        sTipoRelatorio = 2
    ElseIf Frm.OptPortCliente.Value = True Then
        sTipoRelatorio = 3
    ElseIf Frm.OptPortDesignacao.Value = True Then
        sTipoRelatorio = 4
    Else
        sTipoRelatorio = 0
    End If
    
    objServico.GbCodigoFasePortabilidadeAgendamento = GbCodigoFasePortabilidadeAgendamento
    objServico.GbCodigoFasePortabilidadeJanelaMigracao = GbCodigoFasePortabilidadeJanelaMigracao
    objServico.GbCodigoFasePortabilidadeEstorno = GbCodigoFasePortabilidadeEstorno
    
    objServico.InformaMensagem objMensagem
    
    bRetorno = objServico.ObtemDadosMonitoracaoPort(Frm.CboServico.ItemData(Frm.CboServico.ListIndex), _
                                                                                    Frm.dtpDataIni.Value, Frm.dtpDataFim.Value, _
                                                                                    Frm.cboCentroFuncional.ItemData(Frm.cboCentroFuncional.ListIndex), sTipoRelatorio, _
                                                                                    aListaMonitoracao(), _
                                                                                    sCabecalhoColunas, _
                                                                                    "", "")
             
    If Not bRetorno Then
       
       If objMensagem.TipoMsg <> "" Then
          
          Frm.sprMonitoracao.MaxRows = 0
          
          objMensagem.ExibeMensagem
          
          Set objServico = Nothing
          
          Exit Function
       
       End If
       
       objMensagem.TipoMsg = "A"
       objMensagem.Descricao = "Não foi possível carregar a lista de OTS's!"
       objMensagem.ExibeMensagem
       
       Set objServico = Nothing
       
       Exit Function
    
    End If
    
    iCol = 0
    iLinhaII = 0
    ReDim vCabecalho(2, 1)
    
    ColAux = 1
    RowAux = 0
    
    Frm.sprMonitoracao.MaxRows = 0
    '*********************************************************************
    ' Responsável: Accenture – Soraia
    ' Data       : 16/04/2008
    ' BO         : CH789557
    ' Observação : Alteração do número máximo de colunas
    '*********************************************************************
    Frm.sprMonitoracao.MaxCols = 5
    
     QtColunas = 1
    
    PosIni = 2
    PosFim = 0

    'Monta o detalhe do Grid
    ReDim VetorServico(12, 1)
    
    ReDim VetorFull(6, 1)
    
    ColAux = 1
    ColFull = 0
    iLinha = 0
    
    Grupo_ID = objServico.ID_Grupo
    
    For iAux = 0 To UBound(aListaMonitoracao, 2)
        
        If IsEmpty(aListaMonitoracao(1, iAux)) Then
        
           Exit For
        
        End If
        
        With Frm.sprMonitoracao
           
           If aListaMonitoracao(1, iAux) <> "****" Then
           'If aListaMonitoracao(12, iAux) <> aListaMonitoracao(12, iAux + 1) Then
             
                      ColFull = 0
                      RowAux = RowAux + 1
                     
                      .MaxRows = RowAux
                      .Row = RowAux
                     
                      .Col = 1
                      .Text = aListaMonitoracao(1, iAux)
                      
                      .Col = 2
                      .Text = aListaMonitoracao(11, iAux)
                      
                      If CLng(aListaMonitoracao(16, iAux)) <> 5 Then
                            .Col = 3
                            .Text = aListaMonitoracao(8, iAux)
                      End If
                              
                    sTipoRelatorio = aListaMonitoracao(13, iAux)
                                                            
                    If sTipoRelatorio = GbCodigoFasePortabilidadeAgendamento Then
                        .Col = 4
                    ElseIf sTipoRelatorio = GbCodigoFasePortabilidadeJanelaMigracao Then
                        .Col = 5
                    End If
                      
                   Select Case aListaMonitoracao(9, iAux)   'Situacao
        
                          Case "NI"                         'Não iniciada
        
                               .BackColor = &H80000005
        
                          Case "IF"                         'Inicio Futuro
        
                               .BackColor = &HFFFF00
        
                          Case "DM"                         'Dentro da Meta
        
                               .BackColor = &HFF00&
        
                          Case "FM"                         'Fora da Meta
        
                               .BackColor = &H8080FF
                               .Text = CInt(aListaMonitoracao(10, iAux))
        
                          Case "CO"                         'Concluída
        
                               .BackColor = &HFFC0C0
        
                          Case "CR"                         'Correção
        
                                .BackColor = &HC00000
        
                          Case Else

                                .BackColor = &HFFFF00
        
                   End Select
                   
                   'Verifica se há pendencia
                    If aListaMonitoracao(14, iAux) = "C" Then   'Pendencia de Cliente
                    
                       .BackColor = &H80C0FF
                    
                    ElseIf aListaMonitoracao(14, iAux) = "E" Then   'Pendencia Embratel
                       
                       .BackColor = &H80FFFF
                    
                    End If
                    
                    'Carrega Vetor para capturar a pendencia, caso haja
                    ReDim Preserve VetorFull(6, iAux)
                    
                    VetorFull(0, iAux) = aListaMonitoracao(15, iAux)  'Titulo da Coluna
                    VetorFull(1, iAux) = RowAux                                 'Número da Linha
                    VetorFull(2, iAux) = aListaMonitoracao(2, iAux)   'Fase_ID correspondente a
                                                                                                'Coluna
                    VetorFull(3, iAux) = aListaMonitoracao(13, iAux)  'CODIGO TIPO FASE
                    VetorFull(4, iAux) = aListaMonitoracao(12, iAux)  'Servico_ID
                    VetorFull(5, iAux) = aListaMonitoracao(7, iAux)   'Data fim Previsto


           End If

            sServico = aListaMonitoracao(12, iAux)
            sServicoSeguinte = aListaMonitoracao(12, iAux + 1)

            If sServico = sServicoSeguinte Then
            
                        sTipoRelatorio = aListaMonitoracao(13, iAux + 1)
                                                                
                        If sTipoRelatorio = GbCodigoFasePortabilidadeAgendamento Then
                            .Col = 4
                        ElseIf sTipoRelatorio = GbCodigoFasePortabilidadeJanelaMigracao Then
                            .Col = 5
                        End If
                          
                       Select Case aListaMonitoracao(9, iAux + 1) 'Situacao
            
                              Case "NI"                         'Não iniciada
            
                                   .BackColor = &H80000005
            
                              Case "IF"                         'Inicio Futuro
            
                                   .BackColor = &HFFFF00
            
                              Case "DM"                         'Dentro da Meta
            
                                   .BackColor = &HFF00&
            
                              Case "FM"                         'Fora da Meta
            
                                   .BackColor = &H8080FF
                                   .Text = CInt(aListaMonitoracao(10, iAux + 1))
            
                              Case "CO"                         'Concluída
            
                                   .BackColor = &HFFC0C0
            
                              Case "CR"                         'Correção
            
                                    .BackColor = &HC00000
                                    
                              Case Else
    
                                    .BackColor = &HFFFF00
           
                       End Select
                
                    'Verifica se há pendencia
                     If aListaMonitoracao(14, iAux + 1) = "C" Then  'Pendencia de Cliente
                     
                        .BackColor = &H80C0FF
                     
                     ElseIf aListaMonitoracao(14, iAux + 1) = "E" Then 'Pendencia Embratel
                        
                        .BackColor = &H80FFFF
                     
                     End If
                     
                    'Carrega Vetor para capturar a pendencia, caso haja
                    ReDim Preserve VetorFull(6, iAux + 1)
                    
                    VetorFull(0, iAux + 1) = aListaMonitoracao(15, iAux + 1) 'Titulo da Coluna
                    VetorFull(1, iAux + 1) = RowAux                               'Número da Linha
                    VetorFull(2, iAux + 1) = aListaMonitoracao(2, iAux + 1) 'Fase_ID correspondente a
                                                                                                'Coluna
                    VetorFull(3, iAux + 1) = aListaMonitoracao(13, iAux + 1) 'CODIGO TIPO FASE
                    VetorFull(4, iAux + 1) = aListaMonitoracao(12, iAux + 1) 'Servico_ID
                    VetorFull(5, iAux + 1) = aListaMonitoracao(7, iAux + 1) 'Data fim Previsto
                     
                
                iAux = iAux + 1
            
            End If

           .Row = RowAux
           .Col = ColAux + 1
        
        End With
    
    Next iAux
    
    Frm.sprMonitoracao.MaxRows = Frm.sprMonitoracao.DataRowCnt
    
    If UBound(vCabecalho, 2) + 1 > 6 Then
       
       Frm.sprMonitoracao.ScrollBars = ScrollBarsBoth
    
    End If
    
    Frm.sprMonitoracao.Row = 1
    Frm.sprMonitoracao.Col = 1
    Frm.sprMonitoracao.Action = 1
    Frm.sprMonitoracao.Action = 0

    '&H80000005& - Branco       - Nao iniciada  - NI
    '&H00FFFF00& - Azul claro   - Inicio Futuro
    '&H0000FF00& - Verde Claro  - Dentro da Meta
    '&H000000FF& - Red          - Fora da Meta
    '&H00FF0000& - Azul escuro  - Concluido
    
    Set objMensagem = Nothing

    ' Linhas
    MontarListaOTSPort = True

End Function

Public Property Get TipoServico() As Long

TipoServico = lTipoServico

End Property

Public Property Let TipoServico(ByVal vNewValue As Long)

lTipoServico = vNewValue

End Property

Public Property Get TipoAcao() As Long

TipoAcao = lTipoAcao

End Property

Public Property Let TipoAcao(ByVal vNewValue As Long)

lTipoAcao = vNewValue

End Property

Public Property Get CentroFuncional_ID() As Long

CentroFuncional_ID = lCentroFuncional_ID

End Property

Public Property Let CentroFuncional_ID(ByVal vNewValue As Long)

lCentroFuncional_ID = vNewValue

End Property

Public Function ExecutaCronograma(plRow As Long) As Boolean

    ' Declara o objeto controle Programar CPE
    Dim objOC_ProgramarCPE       As New clsDNIOC_ProgramarCPE
    Dim objMensagem              As New clsgenMensagem
    Dim objdadosServico          As New clsGenServico
    Dim objDadosFase             As New clsgenFase
    Dim objOCSelecionarOTS       As clsgenOC_SelecionarOTS
    Dim bRetorno                 As Boolean
    Dim lID_FASE                 As Long
    
    If Familia_Produto <> "D" Or plRow = 0 Then
    
        GoTo Finalizar
        
    End If
    
    objOC_ProgramarCPE.Servico_ID = VetorServico(0, plRow)
    
    objdadosServico.InformaBanco frmgenMDI_SGAPlus.objBanco
    objdadosServico.InformaMensagem objMensagem
    
    bRetorno = objdadosServico.RecuperarDados(objOC_ProgramarCPE.Servico_ID)
    
'''    If objdadosServico.Possui_CPE = "N" Then
'''
'''        GoTo Finalizar
'''
'''    End If
    
    'Servico_ID
    
    'Servico_Anterior_ID
    objOC_ProgramarCPE.Servico_ID_Anterior = CLng("0" & VetorServico(1, plRow))
    
    objDadosFase.InformaBanco frmgenMDI_SGAPlus.objBanco
    objDadosFase.InformaMensagem objMensagem
    
    'Fase_ID
    lID_FASE = objDadosFase.ObtemIDFase(objOC_ProgramarCPE.Servico_ID, GbCodigoFaseProgramarCPE)
    
    If objMensagem.TipoMsg <> "" Then
    
        GoTo ErrorHandler
        
    End If
    
    objOC_ProgramarCPE.Fase_ID = lID_FASE
    
    'Acao_ID
    objOC_ProgramarCPE.Acao_ID = VetorServico(3, plRow)
    'objOC_ProgramarCPE.Status_Fase = VetorServico(7, plRow)
    
    'Sequencia_Acao
    objOC_ProgramarCPE.Sequencia_Acao = objDadosFase.ObtemSequenciadaFase( _
                                        objOC_ProgramarCPE.Servico_ID, _
                                        GbCodigoFaseProgramarCPE, lID_FASE, "D")
    
    If objMensagem.TipoMsg <> "" Then
    
        GoTo ErrorHandler
        
    End If
    
    'Sigla_Acao
    objOC_ProgramarCPE.Sigla_Acao = VetorServico(6, plRow)
    
    'Produto_ID
    objOC_ProgramarCPE.Produto_ID = Produto_ID
    
    'Descricao - Status
    
    If objdadosServico.StatusServico = "AN" Then
       
       objOC_ProgramarCPE.Status_Fase = VetorServico(7, plRow) & " / Em Andamento"
    
    Else
       
       objOC_ProgramarCPE.Status_Fase = VetorServico(7, plRow) & " / Finalizado"
    
    End If
    
    'Data finalizacao da fase
    objOC_ProgramarCPE.DataFinalizacao = objDadosFase.ObtemDataFimReal(lID_FASE)
    
    If objMensagem.TipoMsg <> "" Then
    
        GoTo ErrorHandler
        
    End If
    
    'ID do contrato
    objOC_ProgramarCPE.ContratoID = VetorServico(9, plRow)
    objOC_ProgramarCPE.progChamador = "MONITORACAO"
    
    objOC_ProgramarCPE.NumeroOE = objdadosServico.NumOE
    objOC_ProgramarCPE.AnoOE = objdadosServico.AnoOE
    objOC_ProgramarCPE.NumeroItemOE = objdadosServico.ItemOE
    
    Set objOCSelecionarOTS = New clsgenOC_SelecionarOTS
    
    objOC_ProgramarCPE.InformaOC_SelecionarOTS objOCSelecionarOTS
    
    Call objOC_ProgramarCPE.CarregarForm
    
Finalizar:
    
    Set objMensagem = Nothing
    Set objServico = Nothing
    Set objDadosFase = Nothing
    
    Exit Function
    
ErrorHandler:
    
    If Err.Number <> 0 Then
       
       objMensagem.TipoMsg = "E"
       objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                   "Houve um erro de execução no Método: " & vbCrLf & vbCrLf & "ExecutaCronograma " & _
                   "da classe clsGenOC_Monitoracao"
    
    End If
       
    objMensagem.ExibeMensagem
    
    GoTo Finalizar

End Function

Public Function ConsultaPendenciaDiginet(plLinha As Long, psTituloColuna As String) As Boolean

    Dim iAux                As Integer
    Dim objDadosPendencia   As New clsDNI_DadosPendencia
    Dim objMensagem         As New clsgenMensagem
    Dim lCodigoTipoFase     As Long
    Dim lFaseId             As Long
    Dim lServicoID          As Long
    Dim sDataFimPrevisto    As String
    
    ConsultaPendenciaDiginet = False
    
    objDadosPendencia.InformaBanco frmgenMDI_SGAPlus.objBanco
    objDadosPendencia.InformaMensagem objMensagem

    For iAux = 0 To UBound(VetorFull, 2)
        
        If VetorFull(1, iAux) = plLinha And _
                                       Trim(VetorFull(0, iAux)) = Trim(psTituloColuna) Then
        
           lFaseId = VetorFull(2, iAux)
           
           lCodigoTipoFase = VetorFull(3, iAux)
           
           lServicoID = VetorFull(4, iAux)
           
           sDataFimPrevisto = "" & VetorFull(5, iAux)
           
           If objDadosPendencia.RecuperarDados(lFaseId) Then
    
              If objMensagem.TipoMsg <> "" Then
                 
                 Set objDadosPendencia = Nothing
                 
                 GoTo ErrorHandler
                 
              End If
           
              Load FrmDNI_IncluiPendencia
        
              bRetorno = ExibirListaTipoPendencia(Produto_ID, lCodigoTipoFase, _
                                       FrmDNI_IncluiPendencia.CboTipoPendencia, _
                                       objMensagem)
                                      
              If Not bRetorno Then
           
                 Set FrmDNI_IncluiPendencia = Nothing
                
                Exit Function
        
              End If
                                                            
              Select Case objDadosPendencia.Classificacao
                     Case "N"
                          
                          FrmDNI_IncluiPendencia.OptTipoPendencia(0).Value = True
                     
                     Case "P"
                          
                          FrmDNI_IncluiPendencia.OptTipoPendencia(1).Value = True
                     
                     Case "E"
                          
                          FrmDNI_IncluiPendencia.OptTipoPendencia(2).Value = True
              
              End Select
              
              If objDadosPendencia.Classificacao = "P" Or _
                 objDadosPendencia.Classificacao = "E" Then
                 
                 FrmDNI_IncluiPendencia.SpdDataPrevista.Row = 1
                 FrmDNI_IncluiPendencia.SpdDataPrevista.Col = 1
                 FrmDNI_IncluiPendencia.SpdDataPrevista.Text = objDadosPendencia.DataPrevisao
              
              End If
              
              Call SelecionaCombo(FrmDNI_IncluiPendencia.CboTipoPendencia, _
                                                       objDadosPendencia.Tipo_Pendencia_ID)
              
              If objDadosPendencia.Responsabilidade = "C" Then
                
                 FrmDNI_IncluiPendencia.OptResponsabilidade(0).Value = True
            
              ElseIf objDadosPendencia.Responsabilidade = "E" Then
               
                 FrmDNI_IncluiPendencia.OptResponsabilidade(1).Value = True
              
              ElseIf objDadosPendencia.Responsabilidade = "T" Then
               
                 FrmDNI_IncluiPendencia.OptResponsabilidade(2).Value = True
                             
              End If
              
              FrmDNI_IncluiPendencia.TxtObservacao.Text = objDadosPendencia.Observacao
              
              'Recupera dados do Cabecalho
              Call CarregaCabecalho_Diginet(FrmDNI_IncluiPendencia.UsrDNICabecalhoAcesso, _
                                  lServicoID, lFaseId, objMensagem, lCodigoTipoFase)

              ' Em Andamento Atrasado
              If sDataFimPrevisto <> "" Then
                If (Date > CDate(sDataFimPrevisto)) Then
                  
                   sStatus = "Em Andamento (Atrasado)"
                      
                ' Em Andamento à Realizar
                ElseIf (Date <= CDate(sDataFimPrevisto)) Then
                  
                      sStatus = "Em Andamento (A Realizar)"
                
                End If
              
              Else
                  
                  sStatus = "Em Andamento (A Realizar)"
              
              End If
              
              FrmDNI_IncluiPendencia.OptTipoPendencia(0).Enabled = False
              FrmDNI_IncluiPendencia.OptTipoPendencia(1).Enabled = False
              FrmDNI_IncluiPendencia.OptTipoPendencia(2).Enabled = False
              FrmDNI_IncluiPendencia.SpdDataPrevista.Enabled = False
              FrmDNI_IncluiPendencia.CboTipoPendencia.Enabled = False
              FrmDNI_IncluiPendencia.OptResponsabilidade(0).Enabled = False
              FrmDNI_IncluiPendencia.OptResponsabilidade(1).Enabled = False
              FrmDNI_IncluiPendencia.OptResponsabilidade(2).Enabled = False
              FrmDNI_IncluiPendencia.TxtObservacao.Locked = True
              FrmDNI_IncluiPendencia.CmdSalvar.Enabled = False
              FrmDNI_IncluiPendencia.CmdLimpar.Enabled = False
              FrmDNI_IncluiPendencia.LblAcaoStatus.Caption = Trim(sAcaoString) & " / " & sStatus
                      
              If Produto_ID = GbIDProdutoEmbratelIPL Then 'Embratel IPL
               
                 FrmDNI_IncluiPendencia.lblProduto = "E M B R A T E L  I P L"
                 FrmDNI_IncluiPendencia.Caption = "Embratel IPL - Cadastro de Pendência"
            
              Else
            
                 FrmDNI_IncluiPendencia.lblProduto = "D I G I N E T"
                 FrmDNI_IncluiPendencia.Caption = "Diginet - Cadastro de Pendência"
            
              End If
              
              FrmDNI_IncluiPendencia.Show vbModal
               
           End If
        
        End If
    
    Next iAux
    
    ConsultaPendenciaDiginet = True
    
Finalizar:
    
    Set objMensagem = Nothing
    Set objServico = Nothing
    Set objDadosFase = Nothing
    
    Exit Function
    
ErrorHandler:
    
    If Err.Number <> 0 Then

       objMensagem.TipoMsg = "E"
       objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                   "Houve um erro de execução no Método: " & vbCrLf & vbCrLf & "ExecutaCronograma " & _
                   "da classe clsGenOC_Monitoracao"
    
    End If
       
    objMensagem.ExibeMensagem
    
    GoTo Finalizar

End Function

Public Function ConsultaPendenciaProdInternet(plLinha As Long, psTituloColuna As String) As Boolean
    
    
    Dim iAux                As Long
    Dim objDadosPendencia   As New clsblf_Pendencia
    Dim objMensagem         As New clsgenMensagem
    Dim lCodigoTipoFase     As Long
    Dim lFaseId             As Long
    Dim lServicoID          As Long
    Dim sDataFimPrevisto    As String
    Dim bRetorno            As Boolean
    Dim objDetalhePendencia As clsBlf_DetalhePendencia
    Dim lRow                As Long
    
    'IP VPN
    Dim objProduto          As clsGenProduto


    ConsultaPendenciaProdInternet = False

    'Instancia a Produto
    Set objProduto = New clsGenProduto
    
    'Banco e mensagem para a Dados Produto
    objProduto.InformaBanco frmgenMDI_SGAPlus.objBanco
    objProduto.InformaMensagem objMensagem
    
    'Obtém os dados Produto
    bRetorno = objProduto.RecuperarDados(Produto_ID)
    
    If (Not bRetorno) Then
        GoTo ErrorHandler
    End If

    'Banco e mensagem para a Dados Pendência
    objDadosPendencia.InformaBanco frmgenMDI_SGAPlus.objBanco
    objDadosPendencia.InformaMensagem objMensagem
    
    Screen.MousePointer = vbDefault
    
    For iAux = 0 To UBound(VetorFull, 2)

        If VetorFull(1, iAux) = plLinha And _
                                       Trim(VetorFull(0, iAux)) = Trim(psTituloColuna) Then

           lFaseId = VetorFull(2, iAux)

           lCodigoTipoFase = VetorFull(3, iAux)

           lServicoID = VetorFull(4, iAux)

           sDataFimPrevisto = "" & VetorFull(5, iAux)

           If objDadosPendencia.RecuperarDados(lFaseId) Then

              If objMensagem.TipoMsg <> "" Then

                 GoTo ErrorHandler

              End If
                
              frmgenCadastroPendencia.Grupo_ID = lGrupo_ID
              
              Load frmgenCadastroPendencia
                
              bRetorno = ExibirListaTipoPendencia(Produto_ID, lCodigoTipoFase, _
                                          frmgenCadastroPendencia.cmbPendencia, objMensagem)
                                              
              If (Not bRetorno) Then
                   
                 Set frmgenCadastroPendencia = Nothing
                   
                 Exit Function
                
              End If
              
            'CALBAT - 06/10/2008  ====================================================
            'Alterado para ajustar a fase de monitoração, para visualizar pendencias, caso haja
            '=========================================================================
              'If (Produto_ID <> GbIDProdutoJAH And Produto_ID <> GbIDProdutoNetPhone And Produto_ID <> GbIDProdutoVIP) Then  'Outros
              If (Produto_ID <> GbIDProdutoJAH And Produto_ID <> GbIDProdutoNetPhone And objProduto.ID_Grupo <> GbIDGrupoVLI) Then   'Outros
            '=========================================================================
              
                    If objProduto.ID_Grupo = GbIDGrupoBLD Or objProduto.ID_Grupo = GbIDGrupoBLF Then
                        'CALBAT - 13/05/2008  ====================================================
                        'Implementação para acerto da visualização da pendencia                                        =
                        '=========================================================================
                        frmgenCadastroPendencia.UserNEtPhoneCabecalho.Visible = False
                        frmgenCadastroPendencia.usrBIPVPNCabecalho.Visible = False
                        frmgenCadastroPendencia.UsrJAHCabecalho.Visible = False
                        frmgenCadastroPendencia.UsrVozBandaCabecalho.Visible = False
                        frmgenCadastroPendencia.UsrPortabilidadeCabecalho.Visible = False
                        frmgenCadastroPendencia.UsrBLFCabecalho_Acesso.Visible = True
                        '=========================================================================
                        '=========================================================================
                    
                    End If
              
                Call CarregaCabecalho_Business( _
                          frmgenCadastroPendencia.UsrBLFCabecalho_Acesso, _
                          lServicoID, _
                          lFaseId, _
                          objMensagem, _
                          lCodigoTipoFase, _
                          Produto_ID)
                          
              ElseIf (Produto_ID = GbIDProdutoJAH) Then 'JAH
                
                frmgenCadastroPendencia.UserNEtPhoneCabecalho.Visible = False
                frmgenCadastroPendencia.UsrBLFCabecalho_Acesso.Visible = False
                frmgenCadastroPendencia.usrBIPVPNCabecalho.Visible = False
                frmgenCadastroPendencia.UsrJAHCabecalho.Visible = True
                frmgenCadastroPendencia.UsrPortabilidadeCabecalho.Visible = False
                frmgenCadastroPendencia.UsrPortabilidadeCabecalho = False
        
                Call CarregaCabecalho_Business( _
                          frmgenCadastroPendencia.UsrJAHCabecalho, _
                          lServicoID, _
                          lFaseId, _
                          objMensagem, _
                          lCodigoTipoFase, _
                          Produto_ID)
              
              'JRCRUZ - 22/02/2007
              'ElseIf (Produto_ID = GbIDProdutoVIP) Then 'VIP
              ElseIf (objProduto.ID_Grupo = GbIDGrupoVLI) Then  'VIP
                
                frmgenCadastroPendencia.UserNEtPhoneCabecalho.Visible = False
                frmgenCadastroPendencia.UsrBLFCabecalho_Acesso.Visible = False
                frmgenCadastroPendencia.usrBIPVPNCabecalho.Visible = False
                frmgenCadastroPendencia.UsrJAHCabecalho.Visible = False
                frmgenCadastroPendencia.UsrPortabilidadeCabecalho.Visible = False
                frmgenCadastroPendencia.UsrVozBandaCabecalho.Visible = True
                
                Call ObtemDadosCabecalhoVIP( _
                           frmgenCadastroPendencia.UsrVozBandaCabecalho, _
                           lServicoID, _
                          lFaseId, _
                          objMensagem, _
                          lCodigoTipoFase, _
                          Produto_ID)
              
              ElseIf (Produto_ID = GbIDProdutoNetPhone) Then 'NetPhone
                
                frmgenCadastroPendencia.UsrBLFCabecalho_Acesso.Visible = False
                frmgenCadastroPendencia.usrBIPVPNCabecalho.Visible = False
                frmgenCadastroPendencia.UsrJAHCabecalho.Visible = False
                frmgenCadastroPendencia.UsrPortabilidadeCabecalho = False
                frmgenCadastroPendencia.UserNEtPhoneCabecalho.Visible = True
                
                Call ObtemDadosCabecalhoNetPhone( _
                          frmgenCadastroPendencia.UserNEtPhoneCabecalho, _
                          lServicoID, _
                          lFaseId, _
                          objMensagem, _
                          lCodigoTipoFase, _
                          Produto_ID)
              End If
              
            
              If (objMensagem.TipoMsg <> "") Then
              
                 GoTo ErrorHandler
                 
              End If


              ' Em Andamento Atrasado
              If (Date > CDate(sDataFimPrevisto)) Then

                 sStatus = "Em Andamento (Atrasado)"

              ' Em Andamento à Realizar
              ElseIf (Date <= CDate(sDataFimPrevisto)) Then

                    sStatus = "Em Andamento (A Realizar)"

              End If

              frmgenCadastroPendencia.LblAcao_Status.Caption = Trim(sAcaoString) & " / " & sStatus

            
              If objDadosPendencia.ColecaoPendencia.Count <> 0 Then
            
                  frmgenCadastroPendencia.spdListaPendencia.MaxRows = objDadosPendencia.ColecaoPendencia.Count
                
                  lRow = 0
                
                  For Each objDetalhePendencia In objDadosPendencia.ColecaoPendencia
                
                      lRow = lRow + 1
                    
                      frmgenCadastroPendencia.spdListaPendencia.Row = lRow
                      frmgenCadastroPendencia.spdListaPendencia.Col = 1
                      frmgenCadastroPendencia.spdListaPendencia.Text = objDetalhePendencia.DataPrevistaRetirada
                 
                      frmgenCadastroPendencia.spdListaPendencia.Col = 2
                      frmgenCadastroPendencia.spdListaPendencia.Text = objDetalhePendencia.DataInclusao
                 
                      frmgenCadastroPendencia.spdListaPendencia.Col = 3
                      frmgenCadastroPendencia.spdListaPendencia.Text = objDetalhePendencia.DataRetirada
                
                      frmgenCadastroPendencia.spdListaPendencia.Col = 4
                      frmgenCadastroPendencia.spdListaPendencia.Text = objDetalhePendencia.DescrPendencia
                
                      frmgenCadastroPendencia.spdListaPendencia.Col = 5
                    
                      If objDetalhePendencia.Status = "A" Then
                       
                         frmgenCadastroPendencia.spdListaPendencia.Text = "Aberto"
                    
                      ElseIf objDetalhePendencia.Status = "F" Then
                    
                         frmgenCadastroPendencia.spdListaPendencia.Text = "Fechado"
                    
                      End If
                    
                      frmgenCadastroPendencia.spdListaPendencia.Col = 6
                    
                      If objDetalhePendencia.Responsabilidade = "C" Then
                       
                         frmgenCadastroPendencia.spdListaPendencia.Text = "Cliente"
                            
                      ElseIf objDetalhePendencia.Responsabilidade = "E" Then
                    
                         frmgenCadastroPendencia.spdListaPendencia.Text = "Embratel"
                    
                      End If
                    
                      frmgenCadastroPendencia.spdListaPendencia.Col = 7
                      frmgenCadastroPendencia.spdListaPendencia.Text = objDetalhePendencia.iD
                      
                      frmgenCadastroPendencia.spdListaPendencia.Col = 8
                      frmgenCadastroPendencia.spdListaPendencia.Text = objDetalhePendencia.TipoPendenciaID
                      
                      frmgenCadastroPendencia.spdListaPendencia.Col = 9
                      frmgenCadastroPendencia.spdListaPendencia.Text = objDetalhePendencia.ObsPendencia
                      
                      frmgenCadastroPendencia.spdListaPendencia.Col = 10
                      frmgenCadastroPendencia.spdListaPendencia.Text = "N"
                            
                      frmgenCadastroPendencia.spdListaPendencia.Col = 11
                      frmgenCadastroPendencia.spdListaPendencia.Text = objDetalhePendencia.ComandoConfserv
                            
                  Next
        
              End If
            
              frmgenCadastroPendencia.SpdDataPrevistaRetirada.Enabled = False
              frmgenCadastroPendencia.txtOBS.Enabled = False
              frmgenCadastroPendencia.optCliente.Enabled = False
              frmgenCadastroPendencia.optEmbratel.Enabled = False
              frmgenCadastroPendencia.cmbPendencia.Enabled = False
              frmgenCadastroPendencia.CmdLimpar.Enabled = False
              frmgenCadastroPendencia.CmdSalvar.Enabled = False
        
              'IP VPN
              frmgenCadastroPendencia.Caption = Trim(objProduto.Descricao) & " - " & _
                                      "Cadastrar Pendência"

              frmgenCadastroPendencia.lblProduto.Caption = objProduto.Descricao
              
              frmgenCadastroPendencia.informa_OC Me
              
              
              frmgenCadastroPendencia.Grupo_ID = Grupo_ID
              
              Screen.MousePointer = vbDefault
              frmgenCadastroPendencia.Show vbModal
                
              Set frmgenCadastroPendencia = Nothing
              
              Exit For

           End If

        End If

    Next iAux

    ConsultaPendenciaProdInternet = True

Finalizar:

    Screen.MousePointer = vbDefault
    Set objMensagem = Nothing
    Set objServico = Nothing
    Set objDadosPendencia = Nothing
    Set objProduto = Nothing

    Exit Function

ErrorHandler:

    If (Err.Number <> 0) Then

       objMensagem.TipoMsg = "E"
       objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                   "Houve um erro de execução no Método: " & vbCrLf & vbCrLf & "ConsultaPendenciaProdInternet " & _
                   "da classe clsGenOC_Monitoracao"
    End If

    objMensagem.ExibeMensagem

    GoTo Finalizar

End Function
