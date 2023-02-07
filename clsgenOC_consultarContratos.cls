VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsgenOC_ConsultarContratos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private objMensagem                As New clsgenMensagem
Private objBanco                   As SGAPLUSBanco.clsBanco
Private objListaConsultaContrato   As VIP_ConsultaContrato
Private objFormConsultaContratos   As frmgenConsultaContratos
Private objlistadummy              As VIP_ListaDummy
Private sFamilia_Produto           As String

Public Property Get Familia_Produto() As String

    Familia_Produto = sFamilia_Produto

End Property

Public Property Let Familia_Produto(ByVal sNewValue As String)

    sFamilia_Produto = sNewValue
    
End Property

Public Function TelaConsultarConsultarContrato(ByVal psTitulo As String) As Boolean

On Error GoTo ErrorHandler:

    Dim bRetorno As Boolean
     
    Screen.MousePointer = vbArrowHourglass
     
    'Instância a classe de mensagem
    Set objMensagem = New clsgenMensagem
     
    Set objFormConsultaContratos = New frmgenConsultaContratos
    
    If Familia_Produto = "I" Then
    
        objFormConsultaContratos.Caption = "Produtos de Internet - Consultar Contratos"
        
    Else
    
        objFormConsultaContratos.Caption = "Produtos de Dados - Consultar Contratos"
        
    End If
    
    'Informar o objeto controlador
    objFormConsultaContratos.informa_OC Me
    
    'Executar o método de exibição dos produtos
    bRetorno = ExibirListaProduto(objFormConsultaContratos.cmbProdutos, Familia_Produto)
    
    If Not bRetorno Then
    
        GoTo Finalizar
        
    End If
    
    objFormConsultaContratos.Show
     
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
                            "Houve um erro de execução no método: TelaConsultarConsultarContrato " & _
                            "da classe clsgenOC_ConsultarContratos"
    GoTo Finalizar

End Function

Public Function PreencheGrid() As Boolean

On Error GoTo ErrorHandler:

    Dim iLinha        As Long
    Dim sStatus       As String
    Dim bRetorno      As Boolean
    Dim sServico      As String
    Dim sNmCliente    As String
    Dim sCmbPendencia As String
    Dim sCmbStatus    As String
    Dim sDataIni      As String
    Dim sDataFim      As String
    Dim sTipoGrid     As String
    Dim id_Prod       As Integer
    Dim id_Cliente    As Long
    Dim id_Status     As Integer
    Dim Id_Pendencia  As Integer
    Dim objDadosProduto As clsGenProduto
    PreencheGrid = False
    
    Set objListaConsultaContrato = New VIP_ConsultaContrato
    'A o objeto de coleção da lista de consulta contratos
    objListaConsultaContrato.InformaBanco frmgenMDI_SGAPlus.objBanco
    objListaConsultaContrato.InformaMensagem objMensagem
    
    'Critica Preenchimento do Periodo Inicial e Final
    If (IsNull(objFormConsultaContratos.dtpDataIni.Value) = True) And _
       (IsNull(objFormConsultaContratos.dtpDataFim.Value) = True) Then
        
        sDataIni = ""
        sDataFim = ""
    
    ElseIf (IsNull(objFormConsultaContratos.dtpDataIni.Value) = False) And _
           (IsNull(objFormConsultaContratos.dtpDataFim.Value) = False) Then

        sDataIni = Format(objFormConsultaContratos.dtpDataIni.Value, "dd/mm/yyyy")
        sDataFim = Format(objFormConsultaContratos.dtpDataFim.Value, "dd/mm/yyyy")
    
    ElseIf (IsNull(objFormConsultaContratos.dtpDataIni.Value) = True) Then
        
        objMensagem.TipoMsg = "A"
        objMensagem.Descricao = "Data Inicial do Período de Envio Obrigatória."
        
        GoTo Finalizar
    
    ElseIf (IsNull(objFormConsultaContratos.dtpDataFim.Value) = True) Then
        
        objMensagem.TipoMsg = "A"
        objMensagem.Descricao = "Data Final do Período Envio Obrigatória."
         
        GoTo Finalizar
    
    End If
    
    'Critica Servico Designação Circuito
    If objFormConsultaContratos.txtServ.Text = "" And _
       objFormConsultaContratos.cmbProdutos.ListIndex = -1 Then

        objMensagem.TipoMsg = "A"
        
        If objFormConsultaContratos.CkcRejeitado.Value = 1 Or _
           objFormConsultaContratos.Option2.Value = True Then
        
            objMensagem.Descricao = "O Tipo de Produto Deve Estar Preecnhido"

        Else
            objMensagem.Descricao = "O Tipo de Produto ou Servico Devem Estar Preecnhidos"
            
        End If
        
         GoTo Finalizar
    
    End If


    If objFormConsultaContratos.txtServ.Text <> "" Then
        
        sServico = objFormConsultaContratos.txtServ.Text
    
    ElseIf objFormConsultaContratos.cmbProdutos.ListIndex <> -1 Then
    
        Set objDadosProduto = New clsGenProduto
        
        objDadosProduto.InformaBanco frmgenMDI_SGAPlus.objBanco
        objDadosProduto.InformaMensagem objMensagem
                
        bRetorno = objDadosProduto.RecuperarDados(CLng(objFormConsultaContratos.cmbProdutos.ItemData(objFormConsultaContratos.cmbProdutos.ListIndex)))
        
        sServico = objDadosProduto.Sigla
        
        Set objDadosProduto = Nothing

'        sServico = objListaConsultaContrato.CriticaValidaProduto("", objFormConsultaContratos.cmbProdutos.ItemData(objFormConsultaContratos.cmbProdutos.ListIndex))
    
    End If
    
    'Condição: Somente se o Caixa Rejeitado estiver Desmarcada
    If objFormConsultaContratos.CkcRejeitado.Value = 0 Then
    
        'Preencher o Combo de Cliente
        If objFormConsultaContratos.CmbCliente.ListIndex <> -1 Then
        
            sNmCliente = objFormConsultaContratos.CmbCliente.List(objFormConsultaContratos.CmbCliente.ListIndex)
        
        End If
        
        'Critica Status do Contrato
        If objFormConsultaContratos.cmbStatus.ListIndex <> -1 Then
            
            If objFormConsultaContratos.cmbStatus.ItemData(objFormConsultaContratos.cmbStatus.ListIndex) = 0 Then
                sCmbStatus = "AN"
            Else
                sCmbStatus = "FF"
            End If
         
        End If
        
        'Critica Status Pendencia
        If objFormConsultaContratos.cmbPendencia.ListIndex <> -1 Then
            
            If objFormConsultaContratos.cmbPendencia.ItemData(objFormConsultaContratos.cmbPendencia.ListIndex) = 0 Then
                sCmbPendencia = "A"
            Else
                sCmbPendencia = "F"
            End If
         
        End If
    
    End If
    
    'Verifica o Tipo de Grid
    'se for grid rejeitado desconsidera o valor da option detalhado / consolidado
    If objFormConsultaContratos.CkcRejeitado.Value = 1 Then
        
        sTipoGrid = "R"
    
    ElseIf objFormConsultaContratos.Option1.Value = True Then
            
        sTipoGrid = "D"
    
    Else
    
        sTipoGrid = "C"

    End If
    
    Screen.MousePointer = vbArrowHourglass
    
    
    'Carrega o Grid
    bRetorno = objListaConsultaContrato.MontarListaConsultaContrato(sServico, _
                                                                    objFormConsultaContratos.TxtCdDiretRegio.Text, _
                                                                    objFormConsultaContratos.TxtNrContr.Text, _
                                                                    objFormConsultaContratos.TxtAAContr.Text, _
                                                                    objFormConsultaContratos.TxtAditivo.Text, _
                                                                    objFormConsultaContratos.TxtContaCorrente.Text, _
                                                                    sNmCliente, _
                                                                    objFormConsultaContratos.TxtCnl.Text, _
                                                                    objFormConsultaContratos.txtServ.Text, _
                                                                    objFormConsultaContratos.TxtSeq.Text, _
                                                                    objFormConsultaContratos.txtCodAreaTelef.Text, _
                                                                    objFormConsultaContratos.TxtNumTelefone.Text, _
                                                                    sDataIni, _
                                                                    sDataFim, _
                                                                    objFormConsultaContratos.txtIDSolic.Text, _
                                                                    objFormConsultaContratos.cmbStatus.List(objFormConsultaContratos.cmbStatus.ListIndex), _
                                                                    objFormConsultaContratos.cmbPendencia.List(objFormConsultaContratos.cmbPendencia.ListIndex), _
                                                                    sTipoGrid)
    
    If Not bRetorno Then
    
        
        GoTo Finalizar
        
    End If
    
    iLinha = 0
    
    For Each objlistadummy In objListaConsultaContrato.ListaConsultaContrato

        If sTipoGrid = "D" Then
        
            iLinha = iLinha + 1
            objFormConsultaContratos.vaSpread1.MaxRows = iLinha
            objFormConsultaContratos.vaSpread1.MaxCols = 18
            
            With objFormConsultaContratos.vaSpread1
    
                'Atribui o valor da linha e define um range para cores
                .BlockMode = False
                .Redraw = True
                
                .Row = iLinha
                .Col = 1
                .Text = objlistadummy.Diretoria
                
                .Row = iLinha
                .Col = 2
                .Text = objlistadummy.NumeroContrato
                
                .Row = iLinha
                .Col = 3
                .Text = objlistadummy.AnoContrato
                
                .Row = iLinha
                .Col = 4
                .Text = objlistadummy.aditivo
                
                .Row = iLinha
                .Col = 5
                .Text = objlistadummy.ContaCorrente
                
                .Row = iLinha
                .Col = 6
                .Text = objlistadummy.NomeCliente
                
                .Row = iLinha
                .Col = 7
                .Text = objlistadummy.DataRecebimento
                
                .Row = iLinha
                .Col = 8
                .Text = objlistadummy.statusCcto
                
                .Row = iLinha
                .Col = 9
                .Text = objlistadummy.codArea
                
                .Row = iLinha
                .Col = 10
                .Text = objlistadummy.NumTerminal
                
                .Row = iLinha
                .Col = 11
                .Text = objlistadummy.Designacao
                
                .Row = iLinha
                .Col = 12
                .Text = objlistadummy.Fase
                
                .Row = iLinha
                .Col = 13
                
                If objlistadummy.Acao = "ATV" Then
                   .Text = "Ativação"
                ElseIf objlistadummy.Acao = "ALT" Then
                   .Text = "Alteração"
                ElseIf objlistadummy.Acao = "CAN" Then
                   .Text = "Cancelamento"
                ElseIf objlistadummy.Acao = "DES" Then
                   .Text = "Desativação"
                ElseIf objlistadummy.Acao = "ATG" Or objlistadummy.Acao = "ALG" Then
                   .Text = "Correção"
                End If
                
                .Row = iLinha
                .Col = 14
                .Text = objlistadummy.Piloto
                
                .Row = iLinha
                .Col = 15
                '.Text = objlistadummy.Descricao
                .Text = objlistadummy.Pendencia
                .Col = 16
                .Text = objlistadummy.StatusPendencia
                .Col = 17
                .Text = objlistadummy.DatainclusaoPendencia
                .Col = 18
                .Text = objlistadummy.DataRetiradaPendencia
                
            End With
        
        ElseIf sTipoGrid = "R" Then
            
            iLinha = iLinha + 1
            objFormConsultaContratos.vaSpread3.MaxRows = iLinha
            objFormConsultaContratos.vaSpread3.MaxCols = 7
                
            With objFormConsultaContratos.vaSpread3
    
                ' Atribui o valor da linha e define um range para cores
                .BlockMode = False
                .Redraw = True
                
                .Row = iLinha
                .Col = 1
                .Text = objlistadummy.Diretoria
                
                .Row = iLinha
                .Col = 2
                .Text = objlistadummy.NomeCliente
                
                .Row = iLinha
                .Col = 3
                .Text = objlistadummy.ContaCorrente
                
                .Row = iLinha
                .Col = 4
                .Text = objlistadummy.NumeroContrato
                
                .Row = iLinha
                .Col = 5
                .Text = objlistadummy.AnoContrato
                
                .Row = iLinha
                .Col = 6
                .Text = objlistadummy.MotivoRejeicao
                
            End With
            
        Else
        
            iLinha = iLinha + 1
            objFormConsultaContratos.vaSpread2.MaxRows = iLinha
            objFormConsultaContratos.vaSpread2.MaxCols = 5
            
            With objFormConsultaContratos.vaSpread2
    
                ' Atribui o valor da linha e define um range para cores
'                .BlockMode = False
                .Redraw = True
                
                .Row = iLinha
                .Col = 1
                .Text = objlistadummy.Fase
                
                .Row = iLinha
                .Col = 2
                .Text = objlistadummy.Andamento
                
                .Row = iLinha
                .Col = 3
                .Text = objlistadummy.Atraso
                
                .Row = iLinha
                .Col = 4
                .Text = objlistadummy.Pendente
                
                .Row = iLinha
                .Col = 5
                .Text = objlistadummy.Total
            
            End With
        
        End If
        
    Next
    
    Screen.MousePointer = vbDefault
    'poe o foco na segunda tab
    objFormConsultaContratos.tabCritConsulta(0).Tab = 1
    PreencheGrid = True

Finalizar:

    Set objListaConsultaContrato = Nothing
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
                            "Houve um erro de execução no método: PreencheGrid " & _
                            "da classe clsgenOC_ConsultarContratos"
    GoTo Finalizar


End Function

Public Function ExibirListaClientes()
On Error GoTo ErrorHandler:

    Dim bRetorno As Boolean
    Dim aListaClientes() As Variant
    
    'Intância a classe de mensagem
    Set objMensagem = New clsgenMensagem
    
     Set objListaConsultaContrato = New VIP_ConsultaContrato

    'Informar o nome do banco desejado
    objListaConsultaContrato.InformaBanco frmgenMDI_SGAPlus.objBanco
                
    'Informar o nome do objeto de mensagem
    objListaConsultaContrato.InformaMensagem objMensagem
    
    'Executar o método de montagem da lista de centro funcional na classe de analisar order entry
    bRetorno = objListaConsultaContrato.MontarListaClientes(aListaClientes, objFormConsultaContratos.cmbProdutos.ItemData(objFormConsultaContratos.cmbProdutos.ListIndex), UCase(objFormConsultaContratos.CmbCliente.Text))
    
    'Se ocorreu algum erro na DLL
    If Not bRetorno Then
        Exit Function
    End If
    
    If objMensagem.TipoMsg = "" Then
    
        PreencheCombo objFormConsultaContratos.CmbCliente, aListaClientes
        
    End If
    
Finalizar:

    Screen.MousePointer = vbDefault
    
    If objMensagem.TipoMsg <> "" Then
    
        objMensagem.ExibeMensagem
        
    End If
    
    Set objListaConsultaContrato = Nothing
    Set objMensagem = Nothing
    
    Exit Function
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: ExibirListaClientes " & _
                            "da classe clsGenOC_ConsultarContratos"
    
    GoTo Finalizar

End Function

Public Function ExportaListaExcel()

On Error GoTo ErrorHandler:
    
    Dim objExcel                 As Object
    Dim sDocExcel                As String
    Dim objMensagem              As clsgenMensagem
    Dim nLinha                   As Long
    Dim i                        As Long
    
    Screen.MousePointer = vbArrowHourglass
    
    Set objMensagem = New clsgenMensagem
    
    If objFormConsultaContratos.Option1.Value And objFormConsultaContratos.vaSpread1.DataRowCnt = 0 Then
       
       objMensagem.TipoMsg = "A"
       objMensagem.Descricao = "Não existem dados a serem exportados para a planilha excel!"
       GoTo Finalizar
        
    End If
    
    If objFormConsultaContratos.Option2.Value And objFormConsultaContratos.vaSpread2.DataRowCnt = 0 Then
       
       objMensagem.TipoMsg = "A"
       objMensagem.Descricao = "Não existem dados a serem exportados para a planilha excel!"
       GoTo Finalizar
        
    End If
    
    If objFormConsultaContratos.CkcRejeitado.Value = 1 And objFormConsultaContratos.vaSpread3.DataRowCnt = 0 Then
       
       objMensagem.TipoMsg = "A"
       objMensagem.Descricao = "Não existem dados a serem exportados para a planilha excel!"
       GoTo Finalizar
        
    End If
    
    ' Intância a classe de mensagem

    ' Instancia o objeto do Word
    Set objExcel = CreateObject("Excel.Application")
    
    'Verifica o Tipo de Grid
    If objFormConsultaContratos.CkcRejeitado.Value = 1 Then
        
        sDocExcel = App.Path & "\ConsultaContratoRejeitado.xlt"
    
    ElseIf objFormConsultaContratos.Option1.Value = True Then
            
        sDocExcel = App.Path & "\ConsultaContratoDetalhado.xlt"
    
    Else
    
        sDocExcel = App.Path & "\ConsultaContratoConsolidado.xlt"

    End If
    
    nLinha = 11

    ' Prepara os campos a serem preechidos
    With objExcel

        .Visible = False
        .Application.WindowState = 1
        .Workbooks.Open sDocExcel, , True
        .CommandBars("Standard").Visible = True
        
        'Verifica o Tipo de Grid
        If objFormConsultaContratos.CkcRejeitado.Value = 1 Then
            
        For i = 1 To objFormConsultaContratos.vaSpread3.MaxRows
            
                'Posiciona na linha do Grid
                objFormConsultaContratos.vaSpread3.Row = i
                
                'Captura na coluna do grid Diretoria
                objFormConsultaContratos.vaSpread3.Col = 1
                .Cells(nLinha, 1) = objFormConsultaContratos.vaSpread3.Text
                
                'Captura na coluna do grid Conta Corrente
                objFormConsultaContratos.vaSpread3.Col = 2
                .Cells(nLinha, 2) = objFormConsultaContratos.vaSpread3.Text
                
                'Captura na coluna do grid Numero do Contrato
                objFormConsultaContratos.vaSpread3.Col = 3
                .Cells(nLinha, 3) = objFormConsultaContratos.vaSpread3.Text
                
                'Captura na coluna do grid Ano Contrato
                objFormConsultaContratos.vaSpread3.Col = 4
                .Cells(nLinha, 4) = objFormConsultaContratos.vaSpread3.Text
                
                'Captura na coluna do grid Aditivo
                objFormConsultaContratos.vaSpread3.Col = 5
                .Cells(nLinha, 5) = " " & objFormConsultaContratos.vaSpread3.Text
                
                'Captura na coluna do grid Motivo Rejeicao
                objFormConsultaContratos.vaSpread3.Col = 6
                .Cells(nLinha, 6) = " " & objFormConsultaContratos.vaSpread3.Text
                
                
                
                nLinha = nLinha + 1
            
            Next
            
            .Columns("A:R").EntireColumn.AutoFit
        
        
        ElseIf objFormConsultaContratos.Option1.Value = True Then
                
            For i = 1 To objFormConsultaContratos.vaSpread1.MaxRows
            
                'Posiciona na linha do Grid
                objFormConsultaContratos.vaSpread1.Row = i
                
                'Captura na coluna do grid Diretoria
                objFormConsultaContratos.vaSpread1.Col = 1
                .Cells(nLinha, 1) = objFormConsultaContratos.vaSpread1.Text
                
                'Captura na coluna do grid Numero do Contrato
                objFormConsultaContratos.vaSpread1.Col = 2
                .Cells(nLinha, 2) = objFormConsultaContratos.vaSpread1.Text
                
                'Captura na coluna do grid Ano Contrato
                objFormConsultaContratos.vaSpread1.Col = 3
                .Cells(nLinha, 3) = objFormConsultaContratos.vaSpread1.Text
                
                'Captura na coluna do grid Aditivo
                objFormConsultaContratos.vaSpread1.Col = 4
                .Cells(nLinha, 4) = objFormConsultaContratos.vaSpread1.Text
                
                'Captura na coluna do grid Conta Corrente
                objFormConsultaContratos.vaSpread1.Col = 5
                .Cells(nLinha, 5) = " " & objFormConsultaContratos.vaSpread1.Text
                
                'Captura na coluna do grid Nome do Cliente
                objFormConsultaContratos.vaSpread1.Col = 6
                .Cells(nLinha, 6) = objFormConsultaContratos.vaSpread1.Text
                
                'Captura na coluna do grid Data de Recebimento
                objFormConsultaContratos.vaSpread1.Col = 7
                .Cells(nLinha, 7) = objFormConsultaContratos.vaSpread1.Text
                
                'Captura na coluna do grid Status do Ccto
                objFormConsultaContratos.vaSpread1.Col = 8
                .Cells(nLinha, 8) = objFormConsultaContratos.vaSpread1.Text
                
                'Captura na coluna do grid Cod.Area
                objFormConsultaContratos.vaSpread1.Col = 9
                .Cells(nLinha, 9) = objFormConsultaContratos.vaSpread1.Text
                
                'Captura na coluna do grid N.Terminal
                objFormConsultaContratos.vaSpread1.Col = 10
                .Cells(nLinha, 10) = objFormConsultaContratos.vaSpread1.Text
                
                'Captura na coluna do grid Designação
                objFormConsultaContratos.vaSpread1.Col = 11
                .Cells(nLinha, 11) = objFormConsultaContratos.vaSpread1.Text
                
                'Captura na coluna do grid Fase
                objFormConsultaContratos.vaSpread1.Col = 12
                .Cells(nLinha, 12) = objFormConsultaContratos.vaSpread1.Text
                
                'Captura na coluna do grid Ação
                objFormConsultaContratos.vaSpread1.Col = 13
                .Cells(nLinha, 13) = objFormConsultaContratos.vaSpread1.Text
                
                'Captura na coluna do grid Piloto
                objFormConsultaContratos.vaSpread1.Col = 14
                .Cells(nLinha, 14) = objFormConsultaContratos.vaSpread1.Text
                
                'Captura na coluna do grid Descricao
                objFormConsultaContratos.vaSpread1.Col = 15
                .Cells(nLinha, 15) = objFormConsultaContratos.vaSpread1.Text
                
                'Captura na coluna do grid Descrição Pendência
                objFormConsultaContratos.vaSpread1.Col = 16
                .Cells(nLinha, 16) = objFormConsultaContratos.vaSpread1.Text
                
                'Captura na coluna do grid Status Pendência
                objFormConsultaContratos.vaSpread1.Col = 17
                .Cells(nLinha, 17) = objFormConsultaContratos.vaSpread1.Text
                
                'Captura na coluna do grid Data Inclusão Pendência
                objFormConsultaContratos.vaSpread1.Col = 18
                .Cells(nLinha, 18) = objFormConsultaContratos.vaSpread1.Text
                
                'Captura na coluna do grid Data de Retirada da Pendência
                objFormConsultaContratos.vaSpread1.Col = 19
                .Cells(nLinha, 19) = objFormConsultaContratos.vaSpread1.Text
                
                nLinha = nLinha + 1
            
            Next
            
            .Columns("A:R").EntireColumn.AutoFit
    
        Else
        
            For i = 1 To objFormConsultaContratos.vaSpread2.MaxRows
            
                'Posiciona na linha do Grid
                objFormConsultaContratos.vaSpread2.Row = i
                
                'Captura na coluna do grid Fase
                objFormConsultaContratos.vaSpread2.Col = 1
                .Cells(nLinha, 1) = objFormConsultaContratos.vaSpread2.Text
                
                'Captura na coluna do grid Andamento
                objFormConsultaContratos.vaSpread2.Col = 2
                .Cells(nLinha, 2) = objFormConsultaContratos.vaSpread2.Text
                
                'Captura na coluna do grid Atraso
                objFormConsultaContratos.vaSpread2.Col = 3
                .Cells(nLinha, 3) = objFormConsultaContratos.vaSpread2.Text
                
                'Captura na coluna do grid Pendente
                objFormConsultaContratos.vaSpread2.Col = 4
                .Cells(nLinha, 4) = objFormConsultaContratos.vaSpread2.Text
                
                'Captura na coluna do grid Total
                objFormConsultaContratos.vaSpread2.Col = 5
                .Cells(nLinha, 5) = objFormConsultaContratos.vaSpread2.Text
                
                nLinha = nLinha + 1
            
            Next
        
            .Rows(nLinha - 1).Font.Bold = True
            .Columns("E").Font.Bold = True
    
        End If
        
        .Cells(6, 2) = Format(Date, "dd/mm/yyyy")
        .Visible = True
        
    
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
                            "da classe clsGenOC_ConsultarContratos"
    GoTo Finalizar

End Function

Public Function ValidaProduto(ByVal pstrTipoProduto As String) As Boolean

On Error GoTo ErrorHandler:

    Dim sSql            As String
    Dim bRetorno        As Boolean
    Dim RsDados         As ADODB.Recordset
    Dim objDadosProduto As clsGenProduto
    
    ValidaProduto = False
    
    Set objDadosProduto = New clsGenProduto
    Set objMensagem = New clsgenMensagem
    
    objDadosProduto.InformaBanco objBanco
    objDadosProduto.InformaMensagem objMensagem
        
    bRetorno = objDadosProduto.ObtemID(pstrTipoProduto)
    
    If bRetorno Then
       
       ValidaProduto = True
    
    End If
    
Finalizar:

    Screen.MousePointer = vbDefault
    
    On Error GoTo 0
    
    If objMensagem.TipoMsg <> "" Then
    
        objMensagem.ExibeMensagem
        
    End If
    
    Set objMensagem = Nothing
    
    Exit Function
    
ErrorHandler:

    If objBanco.objErro Is Nothing Then
       
       objMensagem.TipoMsg = "E"
       objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & _
                               vbCrLf & vbCrLf & vbCrLf & _
                               "Houve um erro de execução no método: " & _
                               vbCrLf & vbCrLf & "ValidaProduto da classe " & _
                               "clsgenOC_ConsultarContrato"
    Else
       objMensagem.TipoMsg = "E"
       objMensagem.Descricao = "Erro: " & objBanco.objErro(0).Number & " - " & _
                               objBanco.objErro(0).Description & vbCrLf & vbCrLf & _
                               vbCrLf & _
                               "Houve um erro de execução no método: " & _
                               vbCrLf & vbCrLf & "ValidaProduto da classe " & _
                               "clsgenOC_ConsultarContrato"
    End If

    GoTo Finalizar
   
End Function


