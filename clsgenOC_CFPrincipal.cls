VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsgenOC_CFPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim objBanco As SGAPLUSBanco.clsBanco
Dim clsCentroFuncional As clsgenCentroFuncional
Dim clsCFSecundario As clsgenOC_CFSecundario
Dim clsAutenticacao As clsgenAutenticacao

Dim CF() As Variant

Private Sub Carrega_Grid(ByRef CF() As Variant)
    Dim i As Integer
    
    With frmGenManterCFPrincipal
        .spdCF.MaxRows = 0
        
        For i = 1 To UBound(CF, 2)
            If IsEmpty(CF(0, i)) Then Exit For
            
            .spdCF.MaxRows = .spdCF.MaxRows + 1
            Call .spdCF.SetText(1, i, CF(1, i))
            Call .spdCF.SetText(2, i, CF(2, i))
            Call .spdCF.SetText(3, i, CStr(CF(0, i)))
        Next
    End With
End Sub

Public Function CarregarForm() As Boolean
    On Error GoTo ErrorHandler
    
    CarregarForm = False
   
    Set clsAutenticacao = New clsgenAutenticacao
    
    Dim paPermissoes() As Variant
    
    clsAutenticacao.IDUsuario = frmgenMDI_SGAPlus.objAutenticacao.IDUsuario
    clsAutenticacao.Login = frmgenMDI_SGAPlus.objAutenticacao.Login
    Call clsAutenticacao.InformaBanco(frmgenMDI_SGAPlus.objBanco, objMsg)
    
    Call clsAutenticacao.RetornarPermissoesFase(clsAutenticacao.IDUsuario, 241, 4, 7, paPermissoes, objMsg)

    If UBound(paPermissoes, 2) = 0 Then
        objMsg.TipoMsg = "E"
        objMsg.TituloMsg = "CarregarForm.CarregarForm"
        objMsg.Descricao = "O usu?rio n?o tem direito de consulta"
        
        GoTo ErrorHandler
    End If
        
    With frmGenManterCFPrincipal
        .Left = 0
        .Height = 7342
        .Width = 11940
        .Top = 0
        .spdCF.Lock = True
        
        .CmdAlterar.Enabled = paPermissoes(4, 1) = 1
        .CmdCadastrar.Enabled = paPermissoes(4, 1) = 1
        .CmdExcluir.Enabled = paPermissoes(5, 1) = 1
    End With
   
    Set clsCentroFuncional = New clsgenCentroFuncional

    Call clsCentroFuncional.InformaBanco(frmgenMDI_SGAPlus.objBanco, objMsg)
    
    If objMsg.TipoMsg <> "" Then GoTo ErrorHandler
    
    If clsCentroFuncional.ListarDados(CF, objMsg) Then
        Call Carrega_Grid(CF)
        CarregarForm = True
    End If
    
    If objMsg.TipoMsg <> "" Then GoTo ErrorHandler

Exit Function

ErrorHandler:
    If objMsg.TipoMsg = "" Then
        objMsg.TipoMsg = "E"
        objMsg.TituloMsg = "clsgenOC_CFPrincipal.CarregarForm"
        objMsg.Descricao = Err.Description
    End If
    objMsg.ExibeMensagem
End Function

Public Function Cadastrar() As Boolean
    Cadastrar = False
    
    frmGenManterCFSecundario.psTipoGravacao = "I"
    frmGenManterCFSecundario.Show vbModal
    
    Cadastrar = True
End Function

Public Function Alterar( _
                    ByVal psID As String) As Boolean
    On Error GoTo ErrorHandler
    Alterar = False
    
    Set clsCFSecundario = New clsgenOC_CFSecundario
    
    Call clsCFSecundario.Carrega_Dados_Alteracao(psID)
    
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
        Set clsCentroFuncional = New clsgenCentroFuncional
        
        Call clsCentroFuncional.InformaBanco(frmgenMDI_SGAPlus.objBanco, objMsg)
        
        If objMsg.TipoMsg <> "" Then GoTo ErrorHandler
        
        If clsCentroFuncional.ExcluirDados(psID, objMsg) Then
            Excluir = True
        End If
    
        If objMsg.TipoMsg <> "" Then GoTo ErrorHandler
        
        objMsg.TipoMsg = "A"
        objMsg.TituloMsg = "Mensagem"
        objMsg.Descricao = "Opera??o efetuada com sucesso!"
        objMsg.ExibeMensagem
    End If
Exit Function

ErrorHandler:
    If objMsg.TipoMsg = "" Then
        objMsg.TipoMsg = "E"
        objMsg.TituloMsg = "clsgenOC_CFPrincipal.CarregarForm"
        objMsg.Descricao = Err.Description
    End If
    objMsg.ExibeMensagem
End Function

Public Function Consultar( _
                    ByVal psID As String) As Boolean
    On Error GoTo ErrorHandler

    Consultar = False
    
    Set clsCentroFuncional = New clsgenOC_CFSecundario
    
    Call clsCFSecundario.Carrega_Dados_Alteracao(psID, True)
    
    Consultar = True
Exit Function

ErrorHandler:
    If objMsg.TipoMsg = "" Then
        objMsg.TipoMsg = "E"
        objMsg.TituloMsg = "clsgenOC_CFPrincipal.CarregarForm"
        objMsg.Descricao = Err.Description
    End If
    objMsg.ExibeMensagem
End Function

Private Sub Class_Initialize()
    Set objMsg = New ClsgenMsg
End Sub

Private Sub Class_Terminate()
    Set objBanco = Nothing
    Set clsCentroFuncional = Nothing
    Set clsCFSecundario = Nothing
    Set clsAutenticacao = Nothing
    Set objMsg = Nothing
End Sub

