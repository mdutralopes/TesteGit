VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsgenOC_CadastrarCF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Declara as classes utilizadas
Private objMensagem As clsgenMensagem
Private objBanco    As SGAPLUSBanco.clsBanco
Private objCadastrarCentroFuncional As clsgenCentroFuncional
Private objListaCentroFuncional As clsgenListaCentroFuncionais

'=============================================================
'Author: Carlos Soares
'Data:  30/11/2005
'Description: Incluso para atende a solicitação de inclusão
' do Objeto Combo de UF
'=============================================================

'Declara a classe de UF
Private objListaUF As clsGenListaUF

'Declara a classe de UF
Private objListaCNL As clsGenListaCNL

'Declara a classe de CEP
Private objListaCEP As clsgenCentroFuncional

Private aListaCNL()                   As Variant
'=============================================================

' Declara os formulários usados nesta classe
Private objFormCadastrarCentroFuncional As frmgenCadastroCentroFuncional

' Declara variáveis comuns
Private lID_CENTRO_FUNCIONAL As Long
Private sTipoOperacao As String
Private lQuantidadeRegistros As Long

Public Property Get TipoOperacao() As String

    TipoOperacao = sTipoOperacao
    
End Property

Public Property Let TipoOperacao(ByVal sNewValue As String)

    sTipoOperacao = sNewValue
    
    ' Chama  método que controla a habilitação e desabilitação dos objetos,
    ' conforme o tipo de operação
    Call HabilitaDesabilitaObjetos
    
    If objFormCadastrarCentroFuncional.gdCEP.MaxRows > 0 Then
        objFormCadastrarCentroFuncional.CmdExcCEP.Enabled = True
    End If
    
    If objFormCadastrarCentroFuncional.gdCNL2.MaxRows > 0 Then
        objFormCadastrarCentroFuncional.cmdExcGrd.Enabled = True
    End If
    
End Property

Public Sub TelaCadastrarCentroFuncional()

On Error GoTo ErrorHandler:

    ' Instância o formulário de pré ativação e chama em seguida
    Set objFormCadastrarCentroFuncional = New frmgenCadastroCentroFuncional
    
    ' Intância a classe de mensagem
    Set objMensagem = New clsgenMensagem
    
    ' Informar o objeto controlador
    objFormCadastrarCentroFuncional.InformaClasseOC Me
    
    ' Instância a classe de dados para cadastrar centros funcionais
    Set objCadastrarCentroFuncional = New clsgenCentroFuncional
    
    ' Informar o nome do banco desejado
    objCadastrarCentroFuncional.InformaBanco frmgenMDI_SGAPlus.objBanco
            
    ' Informar o nome do objeto de mensagem
    objCadastrarCentroFuncional.InformaMensagem objMensagem
    
    ' Chama o método de exibição do GRID
    Call ExibirLista
    
    
Finalizar:

    Screen.MousePointer = vbDefault
    
    If objMensagem.TipoMsg <> "" Then
    
        objMensagem.ExibeMensagem
        
    End If
    
    Set objMensagem = Nothing
    
    On Error GoTo 0
    
    Exit Sub
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: TelaCadastrarCentroFuncional " & _
                            "da classe clsgenOC_CadastrarCF"
    GoTo Finalizar

End Sub
Public Sub CarregaUF()
'=============================================================
'Author: Carlos Soares
'Data:  30/11/2005
'Description: Incluso para atende a solicitação de inclusão
' do Objeto Combo de UF
'=============================================================
On Error GoTo ErrorHandler:

    Dim bRetorno As Boolean
    
    ' Intância a classe da lista de UF's
    Set objListaUF = Nothing
    Set objListaUF = New clsGenListaUF
    
    
    ' Intância a classe de banco
    Set objBanco = New clsBanco
    
    ' Informar o nome do objeto de banco
    objListaUF.InformaBanco frmgenMDI_SGAPlus.objBanco

    
    ' Intância a classe de mensagem
    Set objMensagem = New clsgenMensagem
    
    ' Informar o nome do objeto de mensagem
    objListaUF.InformaMensagem objMensagem
    
    '=============================================================
    'Author: Carlos Soares
    'Data:  30/11/2005
    'Description: Incluso para atende a solicitação de inclusão
    ' do Objeto Combo de UF
    '=============================================================
    '=============================================================
    
    ' Limpa combo
    objFormCadastrarCentroFuncional.CboUF.Clear
    
    ' Executar o método de captura das UF's
    bRetorno = objListaUF.MontarLista(True)
    
    If bRetorno = True Then
        GoTo Finalizar
    End If
    
    objFormCadastrarCentroFuncional.CboUF.ListIndex = 0
    '=============================================================
    '=============================================================
    
    
    If Not bRetorno Then
        
        GoTo Finalizar
    
    End If

    
Finalizar:

    Screen.MousePointer = vbDefault
    
    If objMensagem.TipoMsg <> "" Then
    
        objMensagem.ExibeMensagem
        
    End If
    
    objMensagem.TipoMsg = ""
    
    On Error GoTo 0
    
    Exit Sub


ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: CarregaUF " & _
                            "da classe clsgenOC_CadastrarCF"

    GoTo Finalizar

End Sub
Public Sub ExibeCNL()

On Error GoTo ErrorHandler:

    Dim bRetorno As Boolean
    
    Screen.MousePointer = vbArrowHourglass
    
    'Setando a classe
    Set objListaCNL = New clsGenListaCNL
    
    objListaCNL.InformaBanco frmgenMDI_SGAPlus.objBanco
    
    
    ' Intância a classe da lista de CNL's
    Set objListaCNL = Nothing
    Set objListaCNL = New clsGenListaCNL
    
    
    ' Intância a classe de banco
    'Set objBanco = New clsBanco
    
    ' Informar o nome do objeto de banco
    objListaCNL.InformaBanco frmgenMDI_SGAPlus.objBanco

    
    ' Intância a classe de mensagem
    Set objMensagem = New clsgenMensagem
    
    ' Informar o nome do objeto de mensagem
    objListaCNL.InformaMensagem objMensagem
    
    '=============================================================
    'Author: Carlos Soares
    'Data:  30/11/2005
    'Description: Incluso para atende a solicitação de inclusão
    ' do Objeto Combo de UF
    '=============================================================
    '=============================================================
    
    ' Limpa combo
    'objFormCadastrarCentroFuncional.gdCNL
    
    
    If objFormCadastrarCentroFuncional.CboUFCNL.Text <> "" Then
    
        ' Executar o método de captura das UF's
        bRetorno = objListaCNL.MontarLista(objFormCadastrarCentroFuncional.CboUF, True)
        
        'If bRetorno = False Then
        '   GoTo Finalizar
        'End If
        
    
        '=============================================================
        '=============================================================
        
        objFormCadastrarCentroFuncional.gdCNL.Col = 1
        objFormCadastrarCentroFuncional.gdCNL.ColHidden = False
        objFormCadastrarCentroFuncional.gdCNL.Col = 2
        objFormCadastrarCentroFuncional.gdCNL.ColHidden = False
        objFormCadastrarCentroFuncional.gdCNL.Col = 3
        objFormCadastrarCentroFuncional.gdCNL.ColHidden = True
        
        '=====================================================================
        objFormCadastrarCentroFuncional.cmdIncGrd.Enabled = True
        objFormCadastrarCentroFuncional.CmdIncCEP.Enabled = True
        objFormCadastrarCentroFuncional.TxtCEPIni.BackColor = vbWhite
        objFormCadastrarCentroFuncional.TxtCEPFim.BackColor = vbWhite
        '=====================================================================
    
        Call PreencheGrid
        
        If Not bRetorno Then
            
            GoTo Finalizar
        
        End If
    
    End If

Finalizar:

    Screen.MousePointer = vbDefault
    
    If objMensagem.TipoMsg <> "" Then
    
        objMensagem.ExibeMensagem
        
    End If
    
    objMensagem.TipoMsg = ""
    
    On Error GoTo 0
    
    Exit Sub
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: ExibirLista " & _
                            "da classe clsgenOC_CadastrarCF"

    GoTo Finalizar

End Sub
Public Function ExibirListaGrid(ByVal pOrigem As Integer, ByVal plID_USU_CF As Long) As Boolean

On Error GoTo ErrorHandler:

'    Dim objMensagem As clsgenMensagem
'    Dim bRetorno As Boolean
'
'    ExibirListaGrid = False
'
'    ' Intância a classe de mensagem
'    Set objMensagem = New clsgenMensagem
'
'    ' Informar o nome do banco desejado
'    objCadastrarUsuario_CF.InformaBanco frmgenMDI_SGAPlus.objBanco
'
'    ' Informar o nome do objeto de mensagem
'    objCadastrarUsuario_CF.InformaMensagem objMensagem
'
'    ' Executar o método de montagem da lista de centro funcional relacionados ao usuário
'    bRetorno = objCadastrarUsuario_CF.MontaListaGrid(plID_USU_CF, aListaCF, pOrigem)
'
'   ' Se ocorreu algum erro na DLL exibi-lo
'    If Not bRetorno Then
'
'        GoTo Finalizar
'
'    End If
'
'    ' Preenche a lista de usuários
'    Call PreencheGrid
'
'    ExibirListaGrid = True
    
Finalizar:

'    Screen.MousePointer = vbDefault
'
'    If objMensagem.TipoMsg <> "" Then
'
'        objMensagem.ExibeMensagem
'
'    End If
'
'    Set objCadastrarUsuario_CF = Nothing
'    Set objMensagem = Nothing
    
    Exit Function
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: ExibirListaGrid " & _
                            "da classe clsGenOC_CadastrarUsuario_CF"
    
    GoTo Finalizar

End Function
Public Sub ExcluirNoGridCEP(ByVal sCEPIni As Long, _
                           ByVal sCEPFim As Long, _
                           ByVal gGrid As Object)
                           

    Dim iContador As Integer
    Dim sCampo1, sCampo2 As String
    Dim iLinha As Integer
    
    'Verifica se grid esta vazio
    If gGrid.MaxRows = 0 Then
    
        Exit Sub
        
    End If

    'Seta variavel numero de linhas do grid
    'Pode haver linhas escondidas, que marcam os regs para exclusão
    gGrid.Row = gGrid.ActiveRow
    gGrid.Action = 5
    gGrid.MaxRows = gGrid.MaxRows - 1
    
    
    ' Atualiza o array
    ' Atribui os valores do campos nas propriedades
    
    'objListaCNL.Nome_Localidade = sDescCNL
    'objListaCNL.Sigla_Localidade = sCNL
    
    'objListaCNL.ListaCNL.Remove objListaCNL
    
    
    
    If gGrid.MaxRows = 0 Then
                           
         objFormCadastrarCentroFuncional.CmdExcCEP.Enabled = False
                           
    End If
    
    
Finalizar:

    Screen.MousePointer = vbDefault

    If objMensagem.TipoMsg <> "" Then

        objMensagem.ExibeMensagem

    End If

End Sub
Public Sub ExcluirNoGrid(gGrid As Object, cBotao As Object)
                         
    'Verifica se grid esta vazio
    If gGrid.MaxRows = 0 Then
        Exit Sub
    End If

    'Seta variavel numero de linhas do grid
    'Pode haver linhas escondidas, que marcam os regs para exclusão
    gGrid.Row = gGrid.ActiveRow
    gGrid.Action = 5
    gGrid.MaxRows = gGrid.MaxRows - 1
    
    'Habilita Controles ====================================================
    If gGrid.MaxRows = 0 Then
        cBotao.Enabled = False
        'objFormCadastrarCentroFuncional.gdCNL2.Refresh
        gGrid.Refresh
    End If
    '=======================================================================

End Sub
Public Sub IncluiNoGridCep(ByVal sCEPIni As Long, _
                           ByVal sCEPFim As Long, _
                           ByVal sCodCNL As Long, _
                           ByVal gGrid As Object)
    Dim LcRow  As Integer
    Dim cColecao As New Collection
    Dim lRetorno As Long
    Dim lmax As Long
    Dim iColuna As Integer
    Dim iLinha As Integer
    
  
    
    'Inclui info no grid e no array
    gGrid.Row = gGrid.MaxRows
    
    If gGrid.Row = 0 Then
    
        gGrid.MaxRows = 1
        
    Else
    
        gGrid.MaxRows = gGrid.MaxRows + 1
    
    End If
    
        'Inclui na grid do
        gGrid.Row = gGrid.MaxRows

        'Coluna 1 CEP 1
        gGrid.Col = 1
        gGrid.ColHidden = False
        gGrid.Text = sCEPIni

        'Coluna 2 CEP 2
        gGrid.Col = 2
        gGrid.ColHidden = False
        gGrid.Text = sCEPFim
        
        'Coluna 2 CodCNL
        gGrid.Col = 3
        gGrid.ColHidden = True
        gGrid.Text = sCodCNL
        
        
        
        objFormCadastrarCentroFuncional.CmdExcCEP.Enabled = True

    
Finalizar:

    Screen.MousePointer = vbDefault

'    If objMensagem.TipoMsg <> "" Then
'
'        objMensagem.ExibeMensagem
'
'    End If
End Sub
Private Sub LimpaFrameCNL()

    'objFormCadastrarCentroFuncional.CboUF.ListIndex = -1
    
    objFormCadastrarCentroFuncional.gdCNL2.MaxRows = 0
    objFormCadastrarCentroFuncional.gdCNL.MaxRows = 0
    
    objFormCadastrarCentroFuncional.cmdExcGrd.Enabled = False
    objFormCadastrarCentroFuncional.cmdIncGrd.Enabled = False

End Sub
Private Sub LimpaFrameCEP()

    objFormCadastrarCentroFuncional.TxtCEPIni.Text = ""
    objFormCadastrarCentroFuncional.TxtCEPFim.Text = ""
    
    objFormCadastrarCentroFuncional.TxtCEPIni.BackColor = &H80000004
    objFormCadastrarCentroFuncional.TxtCEPFim.BackColor = &H80000004
    
    objFormCadastrarCentroFuncional.gdCEP.MaxRows = 0
    
    objFormCadastrarCentroFuncional.CmdExcCEP.Enabled = False
    objFormCadastrarCentroFuncional.CmdIncCEP.Enabled = False


End Sub
Public Sub IncluirNoGrid2(ByVal sCNL As String, _
                         ByVal sDescCNL As String, _
                         ByVal sCodCNL As String, _
                         gGrid As Object)


    Dim LcRow  As Integer
    Dim cColecao As New Collection
    Dim lRetorno As Long
    Dim lmax As Long
    Dim iColuna As Integer
    Dim iLinha As Integer
    
    ' Intância a classe de mensagem
    Set objMensagem = New clsgenMensagem
       
    ' Intância a classe de mensagem
    Set objListaCNL = New clsGenListaCNL
       
    'Habilita objetos
    objFormCadastrarCentroFuncional.cmdExcGrd.Enabled = True
       
       
    'Verifca se grid esta vazio
    If gGrid.DataRowCnt <> 0 Then
              
        'Verifica se item do combo já existe no grid
        For LcRow = 1 To gGrid.MaxRows
            
            gGrid.Col = 1
            gGrid.Row = LcRow
            
            If gGrid.Text = sCNL Then
            
                If gGrid.RowHidden = False Then
                    ' Atualiza a mensagem de retorno
                    objMensagem.TipoMsg = "A"
                    objMensagem.Descricao = "Informação já cadastrada."
                
                    GoTo Finalizar
                Else
                    gGrid.RowHidden = True
                End If
                
            End If
        Next
        'gGrid.MaxRows = LcRow + 1
    Else
        'Não ha dados no grid
        lmax = 1000
        ReDim aListaCF(4, lmax)
        gGrid.MaxRows = iLinha
    End If
    
    'Inclui info no grid e no array
    gGrid.Row = gGrid.MaxRows
    
    If gGrid.Row = 0 Then
    
        gGrid.MaxRows = 1
        
    Else
    
        gGrid.MaxRows = gGrid.MaxRows + 1
    
    End If
    
        'Inclui na grid do
        gGrid.Row = gGrid.MaxRows

        'Coluna 1 CNL
        gGrid.Col = 1
        gGrid.ColHidden = False
        gGrid.Text = sCNL

        'Coluna 2 Descrição
        gGrid.Col = 2
        gGrid.ColHidden = False
        gGrid.Text = sDescCNL
        
        'Coluna 2 Cod_CNL
        gGrid.Col = 3
        gGrid.ColHidden = True
        gGrid.Text = sCodCNL
        
        ' Atualiza o array
        ' Atribui os valores do campos nas propriedades
        objListaCNL.COD_CNL = sCodCNL
        objListaCNL.Nome_Localidade = sDescCNL
        objListaCNL.Sigla_Localidade = sCNL
        
        objListaCNL.ListaCNL.Add objListaCNL


        objFormCadastrarCentroFuncional.gdCNL.Row = objFormCadastrarCentroFuncional.gdCNL.ActiveRow
        
        'Call objCadastrarCentroFuncional.ExcluirNoGrid(objFormCadastrarCentroFuncional.gdCNL, objFormCadastrarCentroFuncional.cmdExcGrd)
        Call ExcluirNoGrid(objFormCadastrarCentroFuncional.gdCNL, objFormCadastrarCentroFuncional.cmdExcGrd)
        
'        gGrid2.Row = gGrid2.ActiveRow
'        Call objCadastrarCentroFuncional.ExcluirNoGrid(gGrid2, cmdExcGrd)

    
Finalizar:

    Screen.MousePointer = vbDefault

    If objMensagem.TipoMsg <> "" Then

        objMensagem.ExibeMensagem

    End If
     
End Sub

Private Sub PreencheGrid()

    Dim iRow As Integer
    
    iRow = 1
    
    objFormCadastrarCentroFuncional.gdCNL.MaxRows = objListaCNL.ListaCNL.Count
    
    If TipoOperacao = "I" Then
        'limpa grid 2
        objFormCadastrarCentroFuncional.gdCNL2.MaxRows = 0
        
        'limpa grid CEP
        objFormCadastrarCentroFuncional.gdCEP.MaxRows = 0
    
    End If
    
    ' Ler o array para preencher o GRID
   For Each objListaCNL In objListaCNL.ListaCNL
        'CNL
        objFormCadastrarCentroFuncional.gdCNL.Col = 1
        objFormCadastrarCentroFuncional.gdCNL.Row = iRow
        objFormCadastrarCentroFuncional.gdCNL.Text = objListaCNL.Sigla_Localidade
        objFormCadastrarCentroFuncional.gdCNL.Col = 2
        objFormCadastrarCentroFuncional.gdCNL.Text = objListaCNL.Nome_Localidade
        objFormCadastrarCentroFuncional.gdCNL.Col = 3
        objFormCadastrarCentroFuncional.gdCNL.Text = objListaCNL.COD_CNL
        
        iRow = iRow + 1
    Next
    
End Sub
Public Sub ExibirLista()

On Error GoTo ErrorHandler:

    Dim bRetorno As Boolean
    
    Screen.MousePointer = vbArrowHourglass
    
    ' Intância a classe de mensagem
    Set objMensagem = New clsgenMensagem
    
    ' Informar o nome do objeto de mensagem
    objCadastrarCentroFuncional.InformaMensagem objMensagem
    
    ' Intância a classe da lista de centros funcionais
    Set objListaCentroFuncional = Nothing
    Set objListaCentroFuncional = New clsgenListaCentroFuncionais
    
    ' Limpa a lista
    objFormCadastrarCentroFuncional.lstCentrosFuncionais.Clear
    objFormCadastrarCentroFuncional.gdCEP.MaxRows = 0
    objFormCadastrarCentroFuncional.gdCNL.MaxRows = 0
    objFormCadastrarCentroFuncional.gdCNL2.MaxRows = 0

    ' Executar o método de montagem da lista na classe de analisar order entry
    bRetorno = objCadastrarCentroFuncional.MontarLista(objFormCadastrarCentroFuncional.optSigla.Value, _
                                                       objFormCadastrarCentroFuncional.txtBusca.Text)
    
    ' Se ocorreu algum erro na DLL exibi-lo
    If Not bRetorno Then
        
        GoTo Finalizar
    
    End If
    
    ' Preenche a lista de centros funcionais
    bRetorno = PreencheLista()
    
    If Not bRetorno Then
        
        GoTo Finalizar
    
    End If
    
    '=============================================================
    'Author: Carlos Soares
    'Data:  30/11/2005
    'Description: Incluso para atende a solicitação de inclusão
    ' do Objeto Combo de UF
    '=============================================================
    
    Set objListaUF = New clsGenListaUF
    
    objListaUF.InformaBanco frmgenMDI_SGAPlus.objBanco
    Call PreencheCombo(objFormCadastrarCentroFuncional.CboUF)
        
    '=============================================================
    '=============================================================
    
    ' Posiciona o na primeira lista caso encontre algum registro
    objFormCadastrarCentroFuncional.lstCentrosFuncionais.ListIndex = 0
    
Finalizar:

    Screen.MousePointer = vbDefault
    
    If objMensagem.TipoMsg <> "" Then
    
        objMensagem.ExibeMensagem
        
    End If
    
    objMensagem.TipoMsg = ""
    
    On Error GoTo 0
    
    Exit Sub
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: ExibirLista " & _
                            "da classe clsgenOC_CadastrarCF"

    GoTo Finalizar

End Sub

Public Function GravarDados() As Boolean

On Error GoTo ErrorHandler:

    Dim bRetorno As Boolean
    
    Dim iLinha As Long
    Dim lCodCNL As Long
    Dim lCepIni As Long
    Dim lCepFim As Long
    Dim sCNL As Long
    
    Dim lCodCF As Long

    GravarDados = False

    ' Intância a classe de mensagem
    Set objMensagem = New clsgenMensagem


    'Informa Banco
    objCadastrarCentroFuncional.InformaBanco frmgenMDI_SGAPlus.objBanco

    ' Informar o nome do objeto de mensagem
    objCadastrarCentroFuncional.InformaMensagem objMensagem

    ' Inicia a transação
    bRetorno = frmgenMDI_SGAPlus.objBanco.Iniciar_Transacao()

        ' Se ocorreu algum erro na DLL
        If Not bRetorno Then

            GoTo Finalizar

        End If

        ' Criticar os campos da tela
        bRetorno = CriticarCamposTela
        
        If Not bRetorno Then
        
            GoTo Finalizar
            
        End If
    
        ' Atribui os campos da tela nas propriedades
        Call AtribuirCamposDaTelaNasPropriedades
        
        ' Excutar o método para gravar os dados do centro funcional
        bRetorno = objCadastrarCentroFuncional.GravarDados(sTipoOperacao)
            
        ' Se ocorreu algum erro na DLL
        If Not bRetorno Then

            GoTo Finalizar

        End If
        
        'Grava CEP do CNL, caso tenha ==========================================
        '=======================================================================
        
        bRetorno = objCadastrarCentroFuncional.ExcluiCEP(objCadastrarCentroFuncional.ID_CodigoCentro_Funcional)
        
        If objFormCadastrarCentroFuncional.gdCEP.MaxRows > 0 Then
        
            
            For iLinha = 1 To objFormCadastrarCentroFuncional.gdCEP.MaxRows
            
                objFormCadastrarCentroFuncional.gdCEP.Row = iLinha
                
                'Captura Coluna CEP Inicial
                objFormCadastrarCentroFuncional.gdCEP.Col = 1
                lCepIni = objFormCadastrarCentroFuncional.gdCEP.Text
                
                'Captura Coluna CEP Final
                objFormCadastrarCentroFuncional.gdCEP.Col = 2
                lCepFim = objFormCadastrarCentroFuncional.gdCEP.Text
                
                sTipoOperacao = objCadastrarCentroFuncional.ValidaCEP(objCadastrarCentroFuncional.ID_CodigoCentro_Funcional, _
                                                                      lCepIni, _
                                                                      lCepFim)
                
                'Grava na Tabela de CEP
                bRetorno = objCadastrarCentroFuncional.GravaCEP(objCadastrarCentroFuncional.ID_CodigoCentro_Funcional, _
                                                                lCepIni, _
                                                                lCepFim)
                ' Se ocorreu algum erro na DLL
                If Not bRetorno Then
        
                    GoTo Finalizar
        
                End If

            Next iLinha
        
        End If
        '=======================================================================
        '=======================================================================
        
        'Grava CNL para CF =====================================================
        '=======================================================================

        bRetorno = objCadastrarCentroFuncional.ExcluiCF_CNL(objCadastrarCentroFuncional.ID_CodigoCentro_Funcional)
        
        If objFormCadastrarCentroFuncional.gdCNL2.MaxRows > 0 Then
        
            For iLinha = 1 To objFormCadastrarCentroFuncional.gdCNL2.MaxRows
            
                'Coluna 1 CNL
                objFormCadastrarCentroFuncional.gdCNL2.Row = iLinha
                objFormCadastrarCentroFuncional.gdCNL2.Col = 3
                objFormCadastrarCentroFuncional.gdCNL2.ColHidden = False
                lCodCNL = objFormCadastrarCentroFuncional.gdCNL2.Text
        
                'sTipoOperacao = objCadastrarCentroFuncional.ValidaCNL(objCadastrarCentroFuncional.ID_CodigoCentro_Funcional, lCodCNL)
                
                bRetorno = objCadastrarCentroFuncional.GravaCF_CNL(objCadastrarCentroFuncional.ID_CodigoCentro_Funcional, _
                                                                   lCodCNL)
            Next iLinha
        
        End If
        '=======================================================================
        '=======================================================================

    ' Finaliza a transação
    frmgenMDI_SGAPlus.objBanco.Finalizar_Transacao
            
    ' Atualiza a mensagem de retorno como gravação realizada com sucesso
    objMensagem.TipoMsg = "A"
    objMensagem.Descricao = "Gravação realizada com sucesso"

    GravarDados = True

Finalizar:

    Screen.MousePointer = vbDefault

    If objMensagem.TipoMsg <> "" Then

        objMensagem.ExibeMensagem

    End If
    
    If Not GravarDados Then
    
        frmgenMDI_SGAPlus.objBanco.Cancelar_Transacao

    End If

    Set objMensagem = Nothing
    
    On Error GoTo 0

    Exit Function

ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: GravarDados " & _
                            "da classe clsgenOC_CadastrarCF"

    GoTo Finalizar

End Function

Public Sub ExcluirDados()

On Error GoTo ErrorHandler:

    Dim bRetorno As Boolean

    ' Intância a classe de mensagem
    Set objMensagem = New clsgenMensagem
    
    objMensagem.TipoMsg = "Q"
    objMensagem.Descricao = "Confirma a exclusão do centro funcional selecionado?"
    If objMensagem.ExibeMensagem = "S" Then

        ' Informar o nome do objeto de mensagem
        objCadastrarCentroFuncional.InformaMensagem objMensagem
    
        ' Inicia a transação
        frmgenMDI_SGAPlus.objBanco.Iniciar_Transacao
          
           ' Excutar o método de exclusão dos dados do centro funcional
            bRetorno = objCadastrarCentroFuncional.ExcluirDados(objCadastrarCentroFuncional.ID_CodigoCentro_Funcional)
    
            ' Se ocorreu algum erro na DLL
            If Not bRetorno Then
    
                frmgenMDI_SGAPlus.objBanco.Cancelar_Transacao
    
                GoTo Finalizar
    
            End If
    
        ' Finaliza a transação
        frmgenMDI_SGAPlus.objBanco.Finalizar_Transacao
        
        ' Se existia somente um registro limpar o objeto de entrada para buscar
        If lQuantidadeRegistros = 1 Then
        
            ' Limpa o campo de busca
            objFormCadastrarCentroFuncional.txtBusca.Text = ""
            
        End If
        
        ' Chama o método de exibição do GRID
        Call ExibirLista
        
    Else
    
        objMensagem.TipoMsg = ""
        
    End If
    
Finalizar:

    Screen.MousePointer = vbDefault

    If objMensagem.TipoMsg <> "" Then

        objMensagem.ExibeMensagem

    End If

    Set objMensagem = Nothing
    
    On Error GoTo 0

    Exit Sub

ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: ExcluirDados " & _
                            "da classe clsgenOC_CadastrarCF"

    Exit Sub

End Sub

Private Sub AtribuirCamposDaTelaNasPropriedades()

    objCadastrarCentroFuncional.Codigo_Centro_Funcional = TrataPlick(Trim(objFormCadastrarCentroFuncional.txtSigla1)) & "-" & _
                                                          TrataPlick(Trim(objFormCadastrarCentroFuncional.txtSigla2)) & "-" & _
                                                          TrataPlick(Trim(objFormCadastrarCentroFuncional.txtSigla3)) & "-" & _
                                                          TrataPlick(Trim(objFormCadastrarCentroFuncional.txtSigla4))
    
    objCadastrarCentroFuncional.Nome_Centro_Funcional = TrataPlick(objFormCadastrarCentroFuncional.TxtDescricao)
    
    objCadastrarCentroFuncional.UF = objFormCadastrarCentroFuncional.CboUF
    
    objCadastrarCentroFuncional.ID_CodigoCentro_Funcional = objFormCadastrarCentroFuncional.LblIdCF
    
End Sub

Private Sub Class_Terminate()

    Set objCadastrarCentroFuncional = Nothing
    Set objMensagem = Nothing
    Set objListaCentroFuncional = Nothing
    
End Sub

Private Function CriticarCamposTela() As Boolean

On Error GoTo ErrorHandler:

    CriticarCamposTela = False
    
    If Len(objFormCadastrarCentroFuncional.txtSigla1.Text) = 0 Then
    
        objMensagem.TipoMsg = "A"
        objMensagem.Descricao = "Sigla incompleta."
        objFormCadastrarCentroFuncional.txtSigla1.SetFocus
        
        GoTo Finalizar
    
    ElseIf Len(objFormCadastrarCentroFuncional.txtSigla2.Text) = 0 Then
    
        objMensagem.TipoMsg = "A"
        objMensagem.Descricao = "Sigla incompleta."
        objFormCadastrarCentroFuncional.txtSigla2.SetFocus
        
        GoTo Finalizar
    
    ElseIf Len(objFormCadastrarCentroFuncional.txtSigla3.Text) = 0 Then
    
        objMensagem.TipoMsg = "A"
        objMensagem.Descricao = "Sigla incompleta."
        objFormCadastrarCentroFuncional.txtSigla3.SetFocus
        
        GoTo Finalizar
        
    ElseIf Len(objFormCadastrarCentroFuncional.TxtDescricao.Text) = 0 Then
    
        objMensagem.TipoMsg = "A"
        objMensagem.Descricao = "Descrição obrigatória."
        objFormCadastrarCentroFuncional.TxtDescricao.SetFocus
        
        GoTo Finalizar
        
    ElseIf Len(objFormCadastrarCentroFuncional.CboUF.Text) = 0 Then
    
        objMensagem.TipoMsg = "A"
        objMensagem.Descricao = "UF obrigatória."
        objFormCadastrarCentroFuncional.CboUF.SetFocus
        
        GoTo Finalizar
    
    End If
    
    CriticarCamposTela = True

Finalizar:

    If objMensagem.TipoMsg <> "" Then
    
        Screen.MousePointer = vbDefault
        
        objMensagem.ExibeMensagem
        
    End If
    
    Exit Function
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: CriticarCamposTela " & _
                            "da classe clsgenOC_CadastrarCF"
    
    GoTo Finalizar

End Function

Public Function SaiTela() As Boolean

On Error GoTo ErrorHandler:

    Dim bAlterou As Boolean

    SaiTela = False

    ' Chama o método para verificar se foram alterados alguns dados
    If AlterouDadosTela Then
       
       objMensagem.TipoMsg = "Q"
       objMensagem.Descricao = "Foram feitas alterações nos dados da tela que " & _
                               "ainda não foram Salvas! Deseja sair assim mesmo ?"
       
       
    End If
    
    SaiTela = True
    
Finalizar:

    If objMensagem.TipoMsg <> "" Then
    
        Screen.MousePointer = vbDefault
    
        SaiTela = objMensagem.ExibeMensagem = "S"
    
    End If
    
    Exit Function
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: SaiTela " & _
                            "da classe clsgenOC_CadastrarCFs"
    
    GoTo Finalizar

End Function

Public Function AlterouDadosTela() As Boolean

On Error GoTo ErrorHandler:

    AlterouDadosTela = True
    
    Set objMensagem = New clsgenMensagem
    
    If SeparaCentroFuncional(objCadastrarCentroFuncional.Codigo_Centro_Funcional, 1) <> objFormCadastrarCentroFuncional.txtSigla1 Then
    
        GoTo Finalizar
        
    End If

    If SeparaCentroFuncional(objCadastrarCentroFuncional.Codigo_Centro_Funcional, 2) <> objFormCadastrarCentroFuncional.txtSigla2 Then
    
        GoTo Finalizar
        
    End If

    If SeparaCentroFuncional(objCadastrarCentroFuncional.Codigo_Centro_Funcional, 3) <> objFormCadastrarCentroFuncional.txtSigla3 Then
    
        GoTo Finalizar
        
    End If

    If SeparaCentroFuncional(objCadastrarCentroFuncional.Codigo_Centro_Funcional, 4) <> objFormCadastrarCentroFuncional.txtSigla4 Then
    
        GoTo Finalizar
        
    End If

    If objCadastrarCentroFuncional.Nome_Centro_Funcional <> objFormCadastrarCentroFuncional.TxtDescricao Then
    
        GoTo Finalizar
        
    End If
    
    AlterouDadosTela = False
    
Finalizar:

    Exit Function
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: AlterouDadosTela " & _
                            "da classe clsgenOC_CadastrarCFs"
    
    GoTo Finalizar

End Function
Private Function PreencheCombo(ByRef cbo As ComboBox) As Boolean

'=============================================================
'=============================================================
'Author: Carlos Soares
'Data:  30/11/2005
'Description: Incluso para atende a solicitação de inclusão
' do Objeto Combo de UF
'=============================================================
'=============================================================

    'Call CarregaUF
On Error GoTo ErrorHandler:

    Dim objMensagem As clsgenMensagem
    Dim bRetorno As Boolean
    Dim iLinha As Integer
    
    PreencheCombo = False
    
    ' Limpa a combo
    cbo.Clear
    
    ' Intância a classe de mensagem
    Set objMensagem = New clsgenMensagem
    
    ' Informar o nome do banco desejado
    objListaUF.InformaBanco frmgenMDI_SGAPlus.objBanco
                
    ' Informar o nome do objeto de mensagem
    objListaUF.InformaMensagem objMensagem
    
    ' Executar o método de montagem da lista de centro funcional na classe de analisar order entry
    bRetorno = objListaUF.MontarLista(True)
    
    ' Se ocorreu algum erro na DLL
    If Not bRetorno Then
        
        Exit Function
    
    End If
    
    If objMensagem.TipoMsg = "" Then
    
        ' Preencher a combo UF
        For Each objListaUF In objListaUF.ListaUF
        
            'Combo de UF do Centro Funcional
            cbo.AddItem objListaUF.UF, iLinha
            'Combo de UF do CNL
            objFormCadastrarCentroFuncional.CboUFCNL.AddItem objListaUF.UF, iLinha
            
            iLinha = iLinha + 1
        Next
        
        
    End If
    
    PreencheCombo = True
    
Finalizar:

    Screen.MousePointer = vbDefault
    
    If objMensagem.TipoMsg <> "" Then
    
        objMensagem.ExibeMensagem
        
    End If
    
    Set objMensagem = Nothing
    
    Exit Function
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: ExibirListaCentroFuncional " & _
                            "da classe clswidOC_AtividadesAbertas"
    
    GoTo Finalizar

'=============================================================
'=============================================================

End Function
Private Function PreencheLista() As Boolean

    lQuantidadeRegistros = 0
    
    PreencheLista = False
    
    ' Ler o objeto de coleção da lista de usuários
    For Each objListaCentroFuncional In objCadastrarCentroFuncional.ListaCentroFuncional

        With objFormCadastrarCentroFuncional

            If .optSigla.Value = True Then

                ' Sigla do centro funcional
                .lstCentrosFuncionais.AddItem (objListaCentroFuncional.Sigla)

            ElseIf objFormCadastrarCentroFuncional.optDescricao.Value = True Then

                ' Descrição do centro funcional
                .lstCentrosFuncionais.AddItem (objListaCentroFuncional.Descricao)

            End If
            
        End With

        lQuantidadeRegistros = lQuantidadeRegistros + 1
    
    Next
    
    PreencheLista = True
    
End Function
Public Sub RefreshDados()


    Screen.MousePointer = vbArrowHourglass

    ' Atualiza os dados de usuário na tela
    With objFormCadastrarCentroFuncional
    
        .txtSigla1 = SeparaCentroFuncional(objCadastrarCentroFuncional.ListaCentroFuncional(objFormCadastrarCentroFuncional.lstCentrosFuncionais.ListIndex + 1).Sigla, _
                                           1)
        .txtSigla2 = SeparaCentroFuncional(objCadastrarCentroFuncional.ListaCentroFuncional(objFormCadastrarCentroFuncional.lstCentrosFuncionais.ListIndex + 1).Sigla, _
                                           2)
        .txtSigla3 = SeparaCentroFuncional(objCadastrarCentroFuncional.ListaCentroFuncional(objFormCadastrarCentroFuncional.lstCentrosFuncionais.ListIndex + 1).Sigla, _
                                           3)
        .txtSigla4 = SeparaCentroFuncional(objCadastrarCentroFuncional.ListaCentroFuncional(objFormCadastrarCentroFuncional.lstCentrosFuncionais.ListIndex + 1).Sigla, _
                                           4)
        .TxtDescricao = objCadastrarCentroFuncional.ListaCentroFuncional(objFormCadastrarCentroFuncional.lstCentrosFuncionais.ListIndex + 1).Descricao
        
        .LblIdCF = objCadastrarCentroFuncional.ListaCentroFuncional(objFormCadastrarCentroFuncional.lstCentrosFuncionais.ListIndex + 1).ID_CENTRO_FUNCIONAL
    
        objCadastrarCentroFuncional.ID_CodigoCentro_Funcional = objCadastrarCentroFuncional.ListaCentroFuncional(objFormCadastrarCentroFuncional.lstCentrosFuncionais.ListIndex + 1).ID_CENTRO_FUNCIONAL
        objCadastrarCentroFuncional.Codigo_Centro_Funcional = objCadastrarCentroFuncional.ListaCentroFuncional(objFormCadastrarCentroFuncional.lstCentrosFuncionais.ListIndex + 1).Sigla
        objCadastrarCentroFuncional.Nome_Centro_Funcional = objCadastrarCentroFuncional.ListaCentroFuncional(objFormCadastrarCentroFuncional.lstCentrosFuncionais.ListIndex + 1).Descricao
        objCadastrarCentroFuncional.UF = objCadastrarCentroFuncional.ListaCentroFuncional(objFormCadastrarCentroFuncional.lstCentrosFuncionais.ListIndex + 1).UF
        
        Call Captura_UF_CF_CNL
    
    End With
    
    Screen.MousePointer = vbDefault

End Sub
Private Sub Captura_UF_CF_CNL()

Dim sUF As String


    'Seta o Objeto da classe de UF ===============================================
'    Set objListaUF = New clsGenListaUF
'
'    objListaUF.InformaBanco frmgenMDI_SGAPlus.objBanco
'    objListaUF.InformaMensagem objMensagem
'
'    sUF = objListaUF.RecuperaUF(objCadastrarCentroFuncional.ID_CodigoCentro_Funcional)
    
    If objCadastrarCentroFuncional.UF <> "" Then
    
        'Procura na combo a UF locaizada
        Call ProcuraTextoNaCombo(objFormCadastrarCentroFuncional.CboUF, objCadastrarCentroFuncional.UF)
        
        'Habilita o botão de inclusão
        objFormCadastrarCentroFuncional.cmdIncGrd.Enabled = False
        objFormCadastrarCentroFuncional.CmdIncCEP.Enabled = False
        
        objFormCadastrarCentroFuncional.TxtCEPIni.BackColor = &H80000004
        objFormCadastrarCentroFuncional.TxtCEPFim.BackColor = &H80000004
        
        objFormCadastrarCentroFuncional.CboUFCNL.Enabled = False

        Recupera_CNL
        Recupera_CEP
    Else
    
        objFormCadastrarCentroFuncional.CboUF.ListIndex = -1
    
        LimpaFrameCEP
        LimpaFrameCNL
    End If
    '=============================================================================
        
        'Set objListaCNL = New clsGenListaCNL
        
        'objListaCNL.RecuperaCNL (objCadastrarCentroFuncional.ID_CodigoCentro_Funcional)
        
        
End Sub
Private Sub Recupera_CEP()

Dim iRow As Long

    Set objListaCEP = New clsgenCentroFuncional

    objListaCEP.InformaBanco frmgenMDI_SGAPlus.objBanco
    objListaCEP.InformaMensagem objMensagem
    
    objListaCEP.RecuperaCEP (objCadastrarCentroFuncional.ID_CodigoCentro_Funcional)
    
    LimpaFrameCEP
    
    iRow = 1
    
    For Each objListaCEP In objListaCEP.ListaCEP
    
        objFormCadastrarCentroFuncional.gdCEP.MaxRows = iRow
        
        objFormCadastrarCentroFuncional.gdCEP.Row = iRow
        objFormCadastrarCentroFuncional.gdCEP.Col = 1
        objFormCadastrarCentroFuncional.gdCEP.Text = objListaCEP.Cep_Inicial
        objFormCadastrarCentroFuncional.gdCEP.Col = 2
        objFormCadastrarCentroFuncional.gdCEP.Text = objListaCEP.Cep_Final
        objFormCadastrarCentroFuncional.gdCEP.Col = 3
        objFormCadastrarCentroFuncional.gdCEP.Text = objListaCEP.Codigo_Centro_Funcional
    
        iRow = iRow + 1
    
    Next

'    Set objListaCEP = New clsgenListaCentroFuncionais
'    objListaCEP.InformaBanco frmgenMDI_SGAPlus.objBanco
'    objListaCEP.InformaMensagem objMensagem
'
'
'
'    iRow = 1
'
'    For Each objListaCNL In objListaCNL.RecuperarCNL
'
'        objFormCadastrarCentroFuncional.gdCNL2.MaxRows = iRow
'
'        'CNL
'        objFormCadastrarCentroFuncional.gdCNL2.Col = 1
'        objFormCadastrarCentroFuncional.gdCNL2.Row = iRow
'        objFormCadastrarCentroFuncional.gdCNL2.Text = objListaCNL.Sigla_Localidade
'        objFormCadastrarCentroFuncional.gdCNL2.Col = 2
'        objFormCadastrarCentroFuncional.gdCNL2.Text = objListaCNL.Nome_Localidade
'        objFormCadastrarCentroFuncional.gdCNL2.Col = 3
'        objFormCadastrarCentroFuncional.gdCNL2.Text = objListaCNL.COD_CNL
'
'        iRow = iRow + 1
'
'    Next

End Sub
Private Sub Recupera_CNL()
Dim iRow As Long
        
    Set objListaCNL = New clsGenListaCNL
    objListaCNL.InformaBanco frmgenMDI_SGAPlus.objBanco
    objListaCNL.InformaMensagem objMensagem
    
    objListaCNL.RecuperaCNL (objCadastrarCentroFuncional.ID_CodigoCentro_Funcional)
    
    iRow = 1
    
    LimpaFrameCNL
    
    For Each objListaCNL In objListaCNL.RecuperarCNL
        
        objFormCadastrarCentroFuncional.gdCNL2.MaxRows = iRow
        
        'CNL
        objFormCadastrarCentroFuncional.gdCNL2.Col = 1
        objFormCadastrarCentroFuncional.gdCNL2.Row = iRow
        objFormCadastrarCentroFuncional.gdCNL2.Text = objListaCNL.Sigla_Localidade
        objFormCadastrarCentroFuncional.gdCNL2.Col = 2
        objFormCadastrarCentroFuncional.gdCNL2.Text = objListaCNL.Nome_Localidade
        objFormCadastrarCentroFuncional.gdCNL2.Col = 3
        objFormCadastrarCentroFuncional.gdCNL2.Text = objListaCNL.COD_CNL
        
        iRow = iRow + 1
        
    Next

End Sub
Public Sub Carrega_info_CNL()
    
    objListaUF.MontarLista (True)
    
    
End Sub

Private Sub HabilitaDesabilitaObjetos()

    ' Coloca cor de fundo como default para habilitação
    If (sTipoOperacao = "I") Or (sTipoOperacao = "A") Then ' Inclusão ou alteração
        
        ' Muda cor de fundo dos objetos de entrada de dados para branco
        objFormCadastrarCentroFuncional.txtSigla1.BackColor = &H80000005
        objFormCadastrarCentroFuncional.txtSigla2.BackColor = &H80000005
        objFormCadastrarCentroFuncional.txtSigla3.BackColor = &H80000005
        objFormCadastrarCentroFuncional.txtSigla4.BackColor = &H80000005
        objFormCadastrarCentroFuncional.TxtDescricao.BackColor = &H80000005
        objFormCadastrarCentroFuncional.CboUF.BackColor = &H80000005
        objFormCadastrarCentroFuncional.CboUFCNL.BackColor = &H80000005
        
        ' Muda cor de escrita dos objetos de entrada de dados para preto
        objFormCadastrarCentroFuncional.txtSigla1.ForeColor = &H80000008
        objFormCadastrarCentroFuncional.txtSigla2.ForeColor = &H80000008
        objFormCadastrarCentroFuncional.txtSigla3.ForeColor = &H80000008
        objFormCadastrarCentroFuncional.txtSigla4.ForeColor = &H80000008
        objFormCadastrarCentroFuncional.TxtDescricao.ForeColor = &H80000008
        objFormCadastrarCentroFuncional.CboUF.ForeColor = &H80000008
        objFormCadastrarCentroFuncional.CboUFCNL.ForeColor = &H80000008
        
        ' Habilita o frame dos objetos de dados
        objFormCadastrarCentroFuncional.frameDados.Enabled = True
        
        ' Desabilita botões de inclusão, alteração, exlusão e frames de busca e lista
        objFormCadastrarCentroFuncional.cmdIncluir.Enabled = False
        objFormCadastrarCentroFuncional.CmdAlterar.Enabled = False
        objFormCadastrarCentroFuncional.cmdExcluir.Enabled = False
        objFormCadastrarCentroFuncional.cmdUsuarioCF.Enabled = False
        objFormCadastrarCentroFuncional.cmdPerfilCF.Enabled = False
        
        objFormCadastrarCentroFuncional.frameBusca.Enabled = False
        objFormCadastrarCentroFuncional.frameLista.Enabled = False
        
        ' Habilita os botões gravar, limpar e cancelar
        objFormCadastrarCentroFuncional.CmdSalvar.Enabled = True
        objFormCadastrarCentroFuncional.CmdLimpar.Enabled = True
        objFormCadastrarCentroFuncional.cmdCancelar.Enabled = True
        
        
        objFormCadastrarCentroFuncional.CmdIncCEP.Enabled = True
        
        objFormCadastrarCentroFuncional.TxtCEPIni.BackColor = vbWhite
        objFormCadastrarCentroFuncional.TxtCEPFim.BackColor = vbWhite
        
        objFormCadastrarCentroFuncional.CboUFCNL.Enabled = True
        
    Else
    
        ' Muda cor de fundo dos objetos de entrada de dados para cinza
        objFormCadastrarCentroFuncional.txtSigla1.BackColor = &H80000004
        objFormCadastrarCentroFuncional.txtSigla2.BackColor = &H80000004
        objFormCadastrarCentroFuncional.txtSigla3.BackColor = &H80000004
        objFormCadastrarCentroFuncional.txtSigla4.BackColor = &H80000004
        objFormCadastrarCentroFuncional.TxtDescricao.BackColor = &H80000004
        
        objFormCadastrarCentroFuncional.CboUF.BackColor = &H80000004
        objFormCadastrarCentroFuncional.CboUFCNL.BackColor = &H80000004
        objFormCadastrarCentroFuncional.CboUFCNL.ListIndex = -1
        objFormCadastrarCentroFuncional.TxtCEPIni.BackColor = &H80000004
        objFormCadastrarCentroFuncional.TxtCEPFim.BackColor = &H80000004
        
    
        ' Muda cor de escrita dos objetos de entrada de dados azul
        objFormCadastrarCentroFuncional.txtSigla1.ForeColor = &HC00000
        objFormCadastrarCentroFuncional.txtSigla2.ForeColor = &HC00000
        objFormCadastrarCentroFuncional.txtSigla3.ForeColor = &HC00000
        objFormCadastrarCentroFuncional.txtSigla4.ForeColor = &HC00000
        objFormCadastrarCentroFuncional.TxtDescricao.ForeColor = &HC00000
        objFormCadastrarCentroFuncional.CboUF.ForeColor = &HC00000
        objFormCadastrarCentroFuncional.CboUFCNL.ForeColor = &HC00000
        objFormCadastrarCentroFuncional.CboUFCNL.ListIndex = -1
                
        
        ' Desabilita o frame dos objetos de dados
        objFormCadastrarCentroFuncional.frameDados.Enabled = False
    
        ' Habilita botões de inclusão, alteração, exlusão e frames de busca e lista
        objFormCadastrarCentroFuncional.cmdIncluir.Enabled = True
        objFormCadastrarCentroFuncional.CmdAlterar.Enabled = True
        objFormCadastrarCentroFuncional.cmdExcluir.Enabled = True
        objFormCadastrarCentroFuncional.cmdUsuarioCF.Enabled = True
        objFormCadastrarCentroFuncional.cmdPerfilCF.Enabled = True
        objFormCadastrarCentroFuncional.frameBusca.Enabled = True
        objFormCadastrarCentroFuncional.frameLista.Enabled = True
        objFormCadastrarCentroFuncional.CmdIncCEP.Enabled = False
        objFormCadastrarCentroFuncional.CmdExcCEP.Enabled = False
        objFormCadastrarCentroFuncional.cmdIncGrd.Enabled = False
        
        ' Desabilita os botões gravar, limpar e cancelar
        objFormCadastrarCentroFuncional.CmdSalvar.Enabled = False
        objFormCadastrarCentroFuncional.CmdLimpar.Enabled = False
        objFormCadastrarCentroFuncional.cmdCancelar.Enabled = False
        'objFormCadastrarCentroFuncional.CboUFCNL.Enabled = False
        
        'Zera as linhas do Grid
        objFormCadastrarCentroFuncional.gdCNL.MaxRows = 0
        objFormCadastrarCentroFuncional.gdCNL2.MaxRows = 0
        objFormCadastrarCentroFuncional.gdCEP.MaxRows = 0
        
        objFormCadastrarCentroFuncional.CboUFCNL.Enabled = False
    
    End If

End Sub

Public Sub LimparDadosTela()

    objFormCadastrarCentroFuncional.txtSigla1.Text = ""
    objFormCadastrarCentroFuncional.txtSigla2.Text = ""
    objFormCadastrarCentroFuncional.txtSigla3.Text = ""
    objFormCadastrarCentroFuncional.txtSigla4.Text = ""
    objFormCadastrarCentroFuncional.TxtDescricao.Text = ""
    
'    objFormCadastrarCentroFuncional.cmdExcGrd.Enabled = False
    objFormCadastrarCentroFuncional.CmdIncCEP.Enabled = False
    objFormCadastrarCentroFuncional.CmdExcCEP.Enabled = False
    objFormCadastrarCentroFuncional.cmdExcGrd.Enabled = False
    
    objFormCadastrarCentroFuncional.gdCNL2.MaxRows = 0
    objFormCadastrarCentroFuncional.gdCEP.MaxRows = 0
    
    
    LimpaFrameCEP
    
    
    Exit Sub
        
End Sub

Public Sub Saida()
    
   Unload objFormCadastrarCentroFuncional
        
End Sub

Private Function SeparaCentroFuncional(ByVal psCentroFuncional As String, _
                                       ByVal iNumeroSigla As Integer) As String

On Error GoTo ErrorHandler:
    
    ' Declara outras variáveis
    Dim iContador As Integer
    Dim iControlador As Integer
    Dim sValorConcatenado As String
    Dim iTamanhoCampo As Integer
    Dim sSigla As String
    
    SeparaCentroFuncional = ""
    
    ' Inicia a variável de controle
    iControlador = 1
    
    ' Inicia a variável de concatenação
    sValorConcatenado = ""
    
    ' Define o tamanho do campo
    iTamanhoCampo = Len(psCentroFuncional)
    
    ' Se o tamanho for igual a ZERO abandonar função
    If iTamanhoCampo = 0 Then
    
        GoTo Finalizar
        
    End If
    
    ' Inicia a variável de sigla
    sSigla = ""
    
    For iContador = 1 To iTamanhoCampo
    
        If (Mid(psCentroFuncional, iContador, 1) <> "-") And (iContador <> iTamanhoCampo) Then
        
            sValorConcatenado = sValorConcatenado & Mid(psCentroFuncional, iContador, 1)
            
        Else
    
            ' Atribui o valor concatenado na variável correspondente
            sSigla = sValorConcatenado
            
            If iContador = iTamanhoCampo Then
            
                sSigla = sSigla & Right(psCentroFuncional, 1)
                
            End If
            
            ' Limpa a variável de concatenação
            sValorConcatenado = ""
            
            If iControlador = iNumeroSigla Then
            
                SeparaCentroFuncional = sSigla

                Exit For
                
            End If
            
            ' Incrementa +1 ao controlador
            iControlador = iControlador + 1
            
        End If
        
    Next
    
Finalizar:

    On Error GoTo 0
    
    Exit Function

ErrorHandler:

    If Err.Number <> 0 Then
       
       objMensagem.TipoMsg = "E"
       objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                           vbCrLf & "Houve um erro de execução no método: " & vbCrLf & vbCrLf & "SeparaCentroFuncional  " & _
                          "da classe clsgenOC_CadastrarCF"
    
    End If

    GoTo Finalizar

                                       
End Function


Public Sub Usuario_CF()
'
'Prepara objetos para chamada da tela Cadastro_Usuario_CF
'Passa parametro de origem da chamada
'Tela Usuarios parametro 1
'Tela CF parametro 2

Dim objCadastrarUsuario_CF  As clsgenOC_CadastrarUsuario_CF

On Error GoTo ErrorHandler

    Set objCadastrarUsuario_CF = New clsgenOC_CadastrarUsuario_CF
    
    Set objMensagem = New clsgenMensagem
    
    objCadastrarUsuario_CF.TelaCadastrarUsuarios_CF 2, objCadastrarCentroFuncional.Codigo_Centro_Funcional, objCadastrarCentroFuncional.ID_CodigoCentro_Funcional
    
    Set objCadastrarUsuario_CF = Nothing

Finalizar:

    Screen.MousePointer = vbDefault
    
    If objMensagem.TipoMsg <> "" Then
    
        objMensagem.ExibeMensagem
        
    End If
    
    Set objMensagem = Nothing
    
    On Error GoTo 0
    
    Exit Sub
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método:  Usuario_CF" & _
                            "da classe clsgenOC_CadastrarUsuario"
    GoTo Finalizar

End Sub

Public Sub Perfil_CF()
'
'Prepara objetos para chamada da tela Cadastro_Perfil_CF
'Passa parametro de origem da chamada

Dim objCadastrarPerfil_CF  As clsgenOC_CadastrarPerfilCF

On Error GoTo ErrorHandler

    Set objCadastrarPerfil_CF = New clsgenOC_CadastrarPerfilCF
    
    Set objMensagem = New clsgenMensagem
    
    objCadastrarPerfil_CF.TelaCadastrarPerfil_CF 2, objCadastrarCentroFuncional.Codigo_Centro_Funcional, objCadastrarCentroFuncional.ID_CodigoCentro_Funcional
    
    Set objCadastrarPerfil_CF = Nothing

Finalizar:

    Screen.MousePointer = vbDefault
    
    If objMensagem.TipoMsg <> "" Then
    
        objMensagem.ExibeMensagem
        
    End If
    
    Set objMensagem = Nothing
    
    On Error GoTo 0
    
    Exit Sub
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método:  Perifl_CF" & _
                            "da classe clsgenOC_CadastrarPerfil"
    GoTo Finalizar

End Sub




