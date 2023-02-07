VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsgenOC_UsuarioCFPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim objBanco As SGAPLUSBanco.clsBanco
Dim clsgenUsuarioCF As clsgenUsuarioXCentroFuncional
Dim clsCentroFuncional As clsgenCentroFuncional
Dim clsUsuarioCFSecundario As clsgenOC_UsuarioCFSecundario
Dim clsAutenticacao As clsgenAutenticacao

Dim UsuarioCF() As Variant

Private Sub Carrega_Grid(ByRef UsuarioCF() As Variant)
    Dim i As Integer
    
    With frmGenManterUsuarioCFPrincipal
        .spdUsuarioCF.MaxRows = 0
        
        For i = 1 To UBound(UsuarioCF, 2)
            If IsEmpty(UsuarioCF(0, i - 1)) Then Exit For
            
            .spdUsuarioCF.MaxRows = .spdUsuarioCF.MaxRows + 1
            Call .spdUsuarioCF.SetText(1, i, UsuarioCF(2, i - 1))
            Call .spdUsuarioCF.SetText(2, i, UsuarioCF(3, i - 1))
            Call .spdUsuarioCF.SetText(3, i, CStr(UsuarioCF(0, i - 1)))
            Call .spdUsuarioCF.SetText(4, i, CStr(UsuarioCF(1, i - 1)))
        Next
    End With
End Sub

Public Function CarregarForm() As Boolean
    On Error GoTo ErrorHandler
    Dim paPermissoes() As Variant

    CarregarForm = False

    Set clsAutenticacao = New clsgenAutenticacao

    clsAutenticacao.IDUsuario = frmgenMDI_SGAPlus.objAutenticacao.IDUsuario
    clsAutenticacao.Login = frmgenMDI_SGAPlus.objAutenticacao.Login
    Call clsAutenticacao.InformaBanco(frmgenMDI_SGAPlus.objBanco, objMsg)
    
    Call clsAutenticacao.RetornarPermissoesFase(clsAutenticacao.IDUsuario, 248, 4, 7, paPermissoes, objMsg)

    If UBound(paPermissoes, 2) = 0 Then
        objMsg.TipoMsg = "E"
        objMsg.TituloMsg = "CarregarForm.CarregarForm"
        objMsg.Descricao = "O usuário não tem direito de consulta"
        
        GoTo ErrorHandler
    End If
        
    With frmGenManterUsuarioCFPrincipal
        .Left = 0
        .Height = 7342
        .Width = 11940
        .Top = 0
        .spdUsuarioCF.Lock = True
        
        .CmdAlterar.Enabled = paPermissoes(4, 1) = 1
        .CmdCadastrar.Enabled = paPermissoes(4, 1) = 1
        .CmdExcluir.Enabled = paPermissoes(5, 1) = 1
    End With
   
    Set clsgenUsuarioCF = New clsgenUsuarioXCentroFuncional
    Set clsCentroFuncional = New clsgenCentroFuncional

    Call clsgenUsuarioCF.InformaBanco(frmgenMDI_SGAPlus.objBanco, objMsg)
    
    If objMsg.TipoMsg <> "" Then GoTo ErrorHandler
    
    If clsgenUsuarioCF.ListarDados(UsuarioCF, objMsg) Then
        Call Carrega_Grid(UsuarioCF)
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
    On Error GoTo ErrorHandler
    
    Cadastrar = False
    
    Dim i As Long
    
    Dim clsgenUsuario As clsgenUsuario
    Dim clsgenCF As clsgenCentroFuncional
    
    Set clsgenUsuario = New clsgenUsuario
    Set clsgenCF = New clsgenCentroFuncional
    
    Dim VetorUsuario() As Variant
    Dim VetorCF() As Variant
    Dim lVetorUsuario() As Variant
    Dim lVetorCF() As Variant
    
    frmGenManterUsuarioCFSecundario.psTipoGravacao = "I"
    
    Call PreencheCombosForm(frmGenManterUsuarioCFSecundario, clsgenUsuario, clsgenCF, lVetorUsuario, lVetorCF, VetorUsuario, VetorCF)
    If objMsg.TipoMsg <> "" Then GoTo ErrorHandler
    
    frmGenManterUsuarioCFSecundario.Show vbModal
    
    Cadastrar = True
Exit Function

ErrorHandler:
    If objMsg.TipoMsg = "" Then
        objMsg.TipoMsg = "E"
        objMsg.TituloMsg = "clsgenOC_TroncoPrincipal.CarregarForm"
        objMsg.Descricao = Err.Description
    End If
    objMsg.ExibeMensagem
End Function

Public Function Alterar( _
                    ByVal psIDUsuario As Long, ByVal psIDCF As Long) As Boolean

    On Error GoTo ErrorHandler
    Alterar = False
    
    Set clsUsuarioCFSecundario = New clsgenOC_UsuarioCFSecundario
    
    Dim clsgenUsuario As clsgenUsuario
    Dim clsgenCF As clsgenCentroFuncional
    
    Set clsgenUsuario = New clsgenUsuario
    Set clsgenCF = New clsgenCentroFuncional
    
    Dim VetorUsuario() As Variant
    Dim VetorCF() As Variant
    Dim lVetorUsuario() As Variant
    Dim lVetorCF() As Variant
    
    Call PreencheCombosForm(frmGenManterUsuarioCFSecundario, clsgenUsuario, clsgenCF, lVetorUsuario, lVetorCF, VetorUsuario, VetorCF)
    
    Call clsUsuarioCFSecundario.Carrega_Dados_Alteracao(psIDUsuario, psIDCF)
    
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
                    ByVal psIDUsuario As Long, ByVal psIDCF As Long) As Boolean
    On Error GoTo ErrorHandler

    Excluir = False
    
    If MsgBox("Deseja realmente excluir este registro?", vbYesNo + vbDefaultButton2, "Mensagem") = vbYes Then
        Set clsgenUsuarioCF = New clsgenUsuarioXCentroFuncional
        
        Call clsgenUsuarioCF.InformaBanco(frmgenMDI_SGAPlus.objBanco, objMsg)
        
        If objMsg.TipoMsg <> "" Then GoTo ErrorHandler
        
        If clsgenUsuarioCF.ExcluirDados(psIDUsuario, psIDCF, objMsg) Then
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
        objMsg.TituloMsg = "clsgenOC_TroncoPrincipal.CarregarForm"
        objMsg.Descricao = Err.Description
    End If
    objMsg.ExibeMensagem
End Function

Public Function Consultar( _
                    ByVal psID As String) As Boolean
    On Error GoTo ErrorHandler
'
'    Consultar = False
'
'    Set clsTroncoSecundario = New clsgenOC_TroncoSecundario
'
'    Call clsTroncoSecundario.Carrega_Dados_Alteracao(psID, True)
'
'    Consultar = True

Exit Function

ErrorHandler:
    If objMsg.TipoMsg = "" Then
        objMsg.TipoMsg = "E"
        objMsg.TituloMsg = "clsgenOC_TroncoPrincipal.CarregarForm"
        objMsg.Descricao = Err.Description
    End If
    objMsg.ExibeMensagem
End Function

Private Sub Class_Initialize()
    Set objMsg = New ClsgenMsg
End Sub

Private Sub Class_Terminate()
    Set objBanco = Nothing
    Set clsgenUsuarioCF = Nothing
    Set clsCentroFuncional = Nothing
    Set clsUsuarioCFSecundario = Nothing
    Set clsAutenticacao = Nothing
    Set objMsg = Nothing
End Sub

Private Sub PreencheCombosForm(ByRef frmLocal As Form, ByRef clsgenUsuario As clsgenUsuario, ByRef clsgenCF As clsgenCentroFuncional, ByRef lVetorUsuario() As Variant, ByRef lVetorCF() As Variant, ByRef VetorUsuario() As Variant, ByRef VetorCF() As Variant)
    Dim i As Integer
    On Error GoTo ErrorHandler
    
    'Recuperando Usuarios
    Call clsgenUsuario.InformaBanco(frmgenMDI_SGAPlus.objBanco, objMsg)
    If objMsg.TipoMsg <> "" Then GoTo ErrorHandler
    Call clsgenUsuario.ListarDados(VetorUsuario, objMsg)
    If objMsg.TipoMsg <> "" Then GoTo ErrorHandler
    
    ReDim lVetorUsuario(1, UBound(VetorUsuario, 2))
    For i = 0 To UBound(VetorUsuario, 2)
        lVetorUsuario(0, i) = VetorUsuario(0, i)
        lVetorUsuario(1, i) = VetorUsuario(1, i)
    Next
    Call PreencheCombo(frmLocal.cboLogin, lVetorUsuario)
    
    'Recuperando Centro Funcional
    Call clsgenCF.InformaBanco(frmgenMDI_SGAPlus.objBanco, objMsg)
    If objMsg.TipoMsg <> "" Then GoTo ErrorHandler
    Call clsgenCF.ListarDados(VetorCF, objMsg)
    If objMsg.TipoMsg <> "" Then GoTo ErrorHandler
        
    ReDim lVetorCF(1, UBound(VetorCF, 2))
    For i = 0 To UBound(VetorCF, 2)
        lVetorCF(0, i) = VetorCF(0, i)
        lVetorCF(1, i) = VetorCF(1, i)
    Next
    Call PreencheCombo(frmLocal.cboCodigo, lVetorCF)

Exit Sub

ErrorHandler:
    If objMsg.TipoMsg = "" Then
        objMsg.TipoMsg = "E"
        objMsg.TituloMsg = "clsgenOC_TroncoPrincipal.CarregarForm"
        objMsg.Descricao = Err.Description
    End If
    objMsg.ExibeMensagem
End Sub
