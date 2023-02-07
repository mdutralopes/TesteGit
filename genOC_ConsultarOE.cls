VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsgenOC_ConsultarOE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Declara as classes utilizadas
Private objMensagem          As clsgenMensagem
Private objServico           As clsGenServico

'============================================================
'CALBAT - 12/02/2008
'Comentado para atender as adaptações
'============================================================
'Private objListaOE           As New clsgenListaOE
Private objListaOTS       As New clsgenListaOTS
'============================================================
'============================================================

Private objOE                As New clsgenOrderEntry

Private objFormConsultarOE   As frmgenConsultaOE
Private lProduto_ID          As Long
Private lGrupo_ID            As Long

' Declara variáveis
Private sFamilia_Produto     As String

Public Property Get Familia_Produto() As String

    Familia_Produto = sFamilia_Produto

End Property

Public Property Let Familia_Produto(ByVal sNewValue As String)

    sFamilia_Produto = sNewValue
    
End Property

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

Public Function TelaConsultarOE(ByVal psTitulo As String) As Boolean

On Error GoTo ErrorHandler:

    Dim bRetorno As Boolean
     
    Screen.MousePointer = vbArrowHourglass
     
    ' Intância a classe de mensagem
    Set objMensagem = New clsgenMensagem
     
    '#################
    'adeil RN-SGAD-14971ZKH 28/10/2003
    'se o form passado por parâmetro já estiver aberto, fecha-lo.
    'inicio
    If Familia_Produto = "I" Then
        If Verifica_Janela_Aberta("frmgenConsultaOE", "Produtos de Internet - Consultar OE") Then
            GoTo Finalizar
        End If
    Else
        If Verifica_Janela_Aberta("frmgenConsultaOE", "Produtos de Dados - Consultar OE") Then
            GoTo Finalizar
        End If
    End If
    'fim
    '#################
     
    Set objFormConsultarOE = New frmgenConsultaOE
    
    If Familia_Produto = "I" Then
        objFormConsultarOE.Caption = "Produtos de Internet - Consultar OE"
    Else
        objFormConsultarOE.Caption = "Produtos de Dados - Consultar OE"
    End If
    
    ' Informar o objeto controlador
    objFormConsultarOE.informa_OC Me
    
    ' Executar o método de exibição dos produtos
    bRetorno = ExibirListaProduto(objFormConsultarOE.CboServico, Familia_Produto)
    
    If Not bRetorno Then
    
        GoTo Finalizar
        
    End If
    
    objFormConsultarOE.Show
     
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
                            "Houve um erro de execução no método: TelaConsultarOE " & _
                            "da classe clsGenOC_ConsultarOE"
    GoTo Finalizar

End Function


Public Function ExibirListaOE() As Boolean

On Error GoTo ErrorHandler:

    Dim bRetorno            As Boolean
    Dim sSiglaProdutoSOE    As String
    Dim vNrOE               As Long
    Dim sTabela             As String
    
    'CALBAT - 19/11/2007 - Acerto para consultar OE ==================
    Dim objProduto        As clsGenProduto
    '==============================================================
    
    ExibirListaOE = False
    
    Set objOE = New clsgenOrderEntry
    objOE.InformaBanco frmgenMDI_SGAPlus.objBanco
               
    Set objMensagem = New clsgenMensagem
    objOE.InformaMensagem objMensagem
    
    'CALBAT - 19/11/2007 - Acerto para consultar OE ==================
    Set objProduto = New clsGenProduto
    objProduto.InformaBanco frmgenMDI_SGAPlus.objBanco
    objProduto.InformaMensagem objMensagem
    '==============================================================
    

    
    If objFormConsultarOE.CboServico.ListIndex > -1 Then
        lProduto_ID = objFormConsultarOE.CboServico.ItemData(objFormConsultarOE.CboServico.ListIndex)
    Else
        lProduto_ID = 0
    End If
    
    'CALBAT - 19/11/2007 - Acerto para consultar OE ==================
    lGrupo_ID = objProduto.ObtemIDGrupo(lProduto_ID)
    '==============================================================
               
    If (lProduto_ID = GbIDProdutoWID) Or _
       (lProduto_ID = GbIDProdutoVPN_RA) Or _
       (lProduto_ID = GbIDProdutoVPN_DIRECT) Then
       
         sSiglaProdutoSOE = "WNN"
         
    ElseIf (lProduto_ID = GbIDProdutoBLF) Then
    
        sSiglaProdutoSOE = "BFN"
        
    'ElseIf (lProduto_ID = GbIDProdutoBLD Or _
                lProduto_ID = GbIDProdutoVOIP) Then
     ElseIf (lGrupo_ID = GbIDGrupoBLD Or _
                 lProduto_ID = GbIDProdutoVOIP) Then
                
        sSiglaProdutoSOE = "INN"
        
    ElseIf (lProduto_ID = GbIDProdutoBIV) Then
        
        sSiglaProdutoSOE = "IVN"
        
    ElseIf (lProduto_ID = GbIDProdutoJAH) Then
        
        sSiglaProdutoSOE = "JAH"
        sTabela = "ATEN"
        
    ElseIf (lProduto_ID = GbIDProdutoDiginet) Then
        
        sSiglaProdutoSOE = "DNI"
        
    End If
    

    If objFormConsultarOE.TxtOE.Text = "" Then
        
        vNrOE = 0
        
    Else
    
        vNrOE = objFormConsultarOE.TxtOE.Text
    
    End If
               
    ' Executar o método de montagem da lista na classe de serviço
    bRetorno = objOE.MontarConsultaOE(vNrOE, _
                                      CStr(objFormConsultarOE.TxtAnoOE.Text), _
                                      sSiglaProdutoSOE, _
                                      lProduto_ID, _
                                      lGrupo_ID, _
                                      sTabela)
                                            
   ' Se ocorreu algum erro na DLL exibi-lo
    If Not bRetorno Then
        GoTo Finalizar
    End If
    
    'Preenche as informações da tela de consulta de OE
    Call PreencheTela
    
    
    'Executa a pesquisa por OTS
    Set objServico = New clsGenServico
    objServico.InformaBanco frmgenMDI_SGAPlus.objBanco
    objServico.InformaMensagem objMensagem
    
    
    bRetorno = objServico.DevolveDesignacao_OTS(lProduto_ID, objOE.NumOE, objOE.AnoOE)
        
    If Not bRetorno Then
        
        GoTo Finalizar
    
    End If
    
    ' Preenche a grid
    bRetorno = PreencheGrid()
    
    If Not bRetorno Then
        
        GoTo Finalizar
    
    End If
    
    ExibirListaOE = True
    
    GoTo Finalizar
       
Finalizar:

    Screen.MousePointer = vbDefault
    
    If objMensagem.TipoMsg <> "" Then
    
        objMensagem.ExibeMensagem
        
    End If
    
    'Set objServico = Nothing
    Set objMensagem = Nothing
    
    Exit Function
        
ErrorHandler:

Resume '    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: ExibirListaOE " & _
                            "da classe clsgenOC_ConsultarOE"
    
    GoTo Finalizar

End Function

'
Public Sub PreencheTela()

    
    objFormConsultarOE.txtOEAno.Text = objOE.NumOE & " / " & objOE.AnoOE
    objFormConsultarOE.txtStatusOE.Text = objOE.StatusOE
    objFormConsultarOE.txtAcaoOE.Text = objOE.Acao
    objFormConsultarOE.txtCoordEncaminhado.Text = objOE.GIC_N
    objFormConsultarOE.TxtContrato.Text = objOE.Contrato
    objFormConsultarOE.txtCliente = objOE.Cliente
    
End Sub

Public Function PreencheGrid() As Boolean

    Dim iLinha As Integer
    Dim sStatus As String
    Dim sOrdem  As String
    Dim i As Integer
    
    iLinha = 0
    
    PreencheGrid = False
    
    'A o objeto de coleção da lista de OE's
     'For Each objListaOE In objOE.ListaOE
    For Each objListaOTS In objServico.ListaOTS

        iLinha = iLinha + 1
        objFormConsultarOE.spdCircuitos.MaxRows = iLinha
        objFormConsultarOE.spdCircuitos.MaxCols = 3

        With objFormConsultarOE.spdCircuitos

            ' Atribui o valor da linha e define um range para cores
            .BlockMode = False
            '.Redraw = False

            
            .Row = iLinha
            .Col = 1
            .Text = objListaOTS.Ots
            
            .Row = iLinha
            .Col = 2
            .Text = objListaOTS.Designacao
            
            .Row = iLinha
            .Col = 3
            .Text = objListaOTS.StatusOTS

        End With

    Next
    
    PreencheGrid = True

End Function


Private Function CriticarCamposTela(ByVal psTipoConsulta As String) As Boolean

On Error GoTo ErrorHandler:

    Dim objMensagem As clsgenMensagem

    ' Intância a classe de mensagem
    Set objMensagem = New clsgenMensagem

    CriticarCamposTela = False
    
    ' Produto não selecionado
    If objFormConsultarOE.CboServico.ListIndex = -1 Then
    
        objMensagem.TipoMsg = "A"
        objMensagem.Descricao = "Produto não selecionado."
        objFormConsultarOE.CboServico.SetFocus
        
        GoTo Finalizar
    
    End If
        
    CriticarCamposTela = True

Finalizar:

    If objMensagem.TipoMsg <> "" Then
    
        Screen.MousePointer = vbDefault
        
        objMensagem.ExibeMensagem
        
    End If
    
    On Error GoTo 0
    
    Exit Function
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: CriticarCamposTela " & _
                            "da classe clsgenOC_ConsultarOE"
    
    GoTo Finalizar

End Function

Private Sub Class_Terminate()

    Set objMensagem = Nothing
    Set objOE = Nothing
    Set objServico = Nothing
    'Set objListaOE = Nothing
    'Set objProduto = Nothing
    
End Sub
