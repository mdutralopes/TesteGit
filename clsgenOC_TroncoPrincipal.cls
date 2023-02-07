VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsgenOC_TroncoPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim objBanco As SGAPLUSBanco.clsBanco
Dim clsgenTronco As clsgenTronco
Dim clsCentroFuncional As clsgenCentroFuncional
Dim clsTroncoSecundario As clsgenOC_TroncoSecundario
Dim clsAutenticacao As clsgenAutenticacao

Dim Tronco() As Variant

Private Sub Carrega_Grid(ByRef Tronco() As Variant)
    Dim i As Integer

    With frmGenManterTroncoPrincipal
        .spdTroncos.MaxRows = 0
        
        For i = 1 To UBound(Tronco, 2)
            If IsEmpty(Tronco(0, i - 1)) Then Exit For
            
            .spdTroncos.MaxRows = .spdTroncos.MaxRows + 1
            Call .spdTroncos.SetText(1, i, Tronco(1, i - 1))
            Call .spdTroncos.SetText(2, i, Tronco(3, i - 1))
            Call .spdTroncos.SetText(3, i, Tronco(2, i - 1))
            Call .spdTroncos.SetText(4, i, CStr(Tronco(4, i - 1)))
            Call .spdTroncos.SetText(5, i, CStr(Tronco(0, i - 1)))
        Next
    End With
End Sub

Public Function CarregarForm() As Boolean
    On Error GoTo ErrorHandler
    
    CarregarForm = False
   
    Dim paPermissoes() As Variant
   
    Set clsAutenticacao = New clsgenAutenticacao
    
    clsAutenticacao.IDUsuario = frmgenMDI_SGAPlus.objAutenticacao.IDUsuario
    clsAutenticacao.Login = frmgenMDI_SGAPlus.objAutenticacao.Login
    Call clsAutenticacao.InformaBanco(frmgenMDI_SGAPlus.objBanco, objMsg)
    
    Call clsAutenticacao.RetornarPermissoesFase(clsAutenticacao.IDUsuario, 246, 4, 7, paPermissoes, objMsg)

    If UBound(paPermissoes, 2) = 0 Then
        objMsg.TipoMsg = "E"
        objMsg.TituloMsg = "clsgenOC_TroncoPrincipal.CarregarForm"
        objMsg.Descricao = "O usuário não tem direito de consulta"
        
        GoTo ErrorHandler
    End If
        
    With frmGenManterTroncoPrincipal
        .Left = 0
        .Height = 7342
        .Width = 11940
        .Top = 0
        .spdTroncos.Lock = True
        
        .CmdAlterar.Enabled = paPermissoes(4, 1) = 1
        .CmdCadastrar.Enabled = paPermissoes(4, 1) = 1
        .CmdExcluir.Enabled = paPermissoes(5, 1) = 1
    End With
   
    Set clsgenTronco = New clsgenTronco
    Set clsCentroFuncional = New clsgenCentroFuncional

    Call clsgenTronco.InformaBanco(frmgenMDI_SGAPlus.objBanco, objMsg)
    
    If objMsg.TipoMsg <> "" Then GoTo ErrorHandler
    
    If clsgenTronco.ListarDados(Tronco, objMsg) Then
        Call Carrega_Grid(Tronco)
        CarregarForm = True
    End If
    
    If objMsg.TipoMsg <> "" Then GoTo ErrorHandler

Exit Function

ErrorHandler:
    If objMsg.TipoMsg = "" Then
        objMsg.TipoMsg = "E"
        objMsg.TituloMsg = "clsgenOC_TroncoPrincipal.CarregarForm"
        objMsg.Descricao = Err.Description
    End If
    objMsg.ExibeMensagem
End Function

Public Function Cadastrar() As Boolean
    Cadastrar = False
    
    frmGenManterTroncoSecundario.psTipoGravacao = "I"
    frmGenManterTroncoSecundario.Show vbModal
    
    Cadastrar = True
End Function

Public Function Alterar( _
                    ByVal psID As String) As Boolean
    On Error GoTo ErrorHandler
    Alterar = False
    
    Set clsTroncoSecundario = New clsgenOC_TroncoSecundario
    
    Call clsTroncoSecundario.Carrega_Dados_Alteracao(psID)
    
    Alterar = True
Exit Function

ErrorHandler:
    If objMsg.TipoMsg = "" Then
        objMsg.TipoMsg = "E"
        objMsg.TituloMsg = "clsgenOC_CFPrincipal.Alterar"
        objMsg.Descricao = Err.Description
    End If
    objMsg.ExibeMensagem
End Function

Public Function Excluir( _
                    ByVal psID As String) As Boolean
    On Error GoTo ErrorHandler

    Excluir = False
    
    If MsgBox("Deseja realmente excluir este registro?", vbYesNo + vbDefaultButton2, "Mensagem") = vbYes Then
        Set clsgenTronco = New clsgenTronco
        
        Call clsgenTronco.InformaBanco(frmgenMDI_SGAPlus.objBanco, objMsg)
        
        If objMsg.TipoMsg <> "" Then GoTo ErrorHandler
        
        If clsgenTronco.ExcluirDados(psID, objMsg) Then
            Excluir = True
        End If
    
        If objMsg.TipoMsg <> "" Then GoTo ErrorHandler
        
        objMsg.TipoMsg = "A"
        objMsg.TituloMsg = "Mensagem"
        objMsg.Descricao = "Operação efetuada com sucesso!"
        objMsg.ExibeMensagem
    End If
Exit Function

ErrorHandler:
    If objMsg.TipoMsg = "" Then
        objMsg.TipoMsg = "E"
        objMsg.TituloMsg = "clsgenOC_TroncoPrincipal.Excluir"
        objMsg.Descricao = Err.Description
    End If
    objMsg.ExibeMensagem
End Function

Public Function Consultar( _
                    ByVal psID As String) As Boolean
    On Error GoTo ErrorHandler

    Consultar = False
    
    Set clsTroncoSecundario = New clsgenOC_TroncoSecundario
    
    Call clsTroncoSecundario.Carrega_Dados_Alteracao(psID, True)
    
    Consultar = True
Exit Function

ErrorHandler:
    If objMsg.TipoMsg = "" Then
        objMsg.TipoMsg = "E"
        objMsg.TituloMsg = "clsgenOC_TroncoPrincipal.Consultar"
        objMsg.Descricao = Err.Description
    End If
    objMsg.ExibeMensagem
End Function

Private Sub Class_Initialize()
    Set objMsg = New ClsgenMsg
End Sub

Private Sub Class_Terminate()
    Set objBanco = Nothing
    Set clsgenTronco = Nothing
    Set clsCentroFuncional = Nothing
    Set clsTroncoSecundario = Nothing
    Set clsAutenticacao = Nothing
    Set objMsg = Nothing
End Sub
