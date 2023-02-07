VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsgenOC_TipoFasePrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim objBanco As SGAPLUSBanco.clsBanco
Dim clsTipoFase As clsgenTipoFase
Dim clsTipoFaseSecundario As clsgenOC_TipoFaseSecundario
Dim clsAutenticacao As clsgenAutenticacao

Dim paDadosTipoFase() As Variant
Dim paPermissoes() As Variant

Private Sub Carrega_Grid(ByRef paDadosTipoFase() As Variant)
    Dim i As Integer
    
    With frmGenManterTiposFasesPrincipal
        .spdTipoFase.MaxRows = 0
        
        For i = 1 To UBound(paDadosTipoFase, 2)
            If IsEmpty(paDadosTipoFase(0, i)) Then Exit For
            
            .spdTipoFase.MaxRows = .spdTipoFase.MaxRows + 1
            Call .spdTipoFase.SetText(1, i, paDadosTipoFase(0, i))
            Call .spdTipoFase.SetText(2, i, paDadosTipoFase(1, i))
            Call .spdTipoFase.SetText(3, i, paDadosTipoFase(2, i))
            Call .spdTipoFase.SetText(4, i, CStr(paDadosTipoFase(3, i)))
            Call .spdTipoFase.SetText(5, i, CStr(paDadosTipoFase(4, i)))
            Call .spdTipoFase.SetText(6, i, CStr(paDadosTipoFase(5, i)))
        Next
    End With
End Sub

Public Function CarregarForm() As Boolean
    On Error GoTo ErrorHandler
    
    CarregarForm = False
  
    Set clsAutenticacao = New clsgenAutenticacao
    
    clsAutenticacao.IDUsuario = frmgenMDI_SGAPlus.objAutenticacao.IDUsuario
    clsAutenticacao.Login = frmgenMDI_SGAPlus.objAutenticacao.Login
    
    Call clsAutenticacao.InformaBanco(frmgenMDI_SGAPlus.objBanco, objMsg)
    
    Call clsAutenticacao.RetornarPermissoesFase(clsAutenticacao.IDUsuario, 245, 4, 7, paPermissoes, objMsg)

    If UBound(paPermissoes, 2) = 0 Then
        objMsg.TipoMsg = "E"
        objMsg.TituloMsg = "CarregarForm.CarregarForm"
        objMsg.Descricao = "O usuário não tem direito de consulta"
        'Unload frmGenManterTiposFasesPrincipal
        GoTo ErrorHandler
    End If
        
    With frmGenManterTiposFasesPrincipal
        .Left = 0
        .Height = 7342
        .Width = 11940
        .Top = 0
        .spdTipoFase.Lock = True
        
        '.CmdAlterar.Enabled = IIf(paPermissoes(4, 1) = 1, True, False)
        .CmdAlterar.Enabled = paPermissoes(4, 1) = 1
        '.CmdCadastrar.Enabled = paPermissoes(4, 1) = 1
        '.CmdExcluir.Enabled = paPermissoes(5, 1) = 1
    End With
    
    Set clsTipoFase = New clsgenTipoFase
    
    Set clsTipoFaseSecundario = New clsgenOC_TipoFaseSecundario

    Call clsTipoFase.InformaBanco(frmgenMDI_SGAPlus.objBanco, objMsg)
    
    If objMsg.TipoMsg <> "" Then GoTo ErrorHandler
    
    If clsTipoFase.ListarDados(paDadosTipoFase, objMsg) Then
        Call Carrega_Grid(paDadosTipoFase)
        CarregarForm = True
    End If
    
    If objMsg.TipoMsg <> "" Then GoTo ErrorHandler
Exit Function

ErrorHandler:
    If objMsg.TipoMsg = "" Then
        objMsg.TipoMsg = "E"
        objMsg.TituloMsg = "clsgenOC_TipoFasePrincipal.CarregarForm"
        objMsg.Descricao = Err.Description
    End If
    objMsg.ExibeMensagem
End Function
Public Function Cadastrar() As Boolean
    Cadastrar = False
    
    Call PreencheComboServico(frmGenManterTiposFasesSecundario.cboServico)
    
    frmGenManterTiposFasesSecundario.psTipoGravacao = "I"
    frmGenManterTiposFasesSecundario.Show vbModal
    
    Cadastrar = True
End Function

Public Function Alterar( _
                    ByVal psID As Double) As Boolean
    On Error GoTo ErrorHandler
    Alterar = False
    
    Set clsTipoFaseSecundario = New clsgenOC_TipoFaseSecundario
    
    Call clsTipoFaseSecundario.Carrega_Dados_Alteracao(psID)
        
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

Private Sub Class_Initialize()
    Set objMsg = New ClsgenMsg
End Sub

Private Sub Class_Terminate()
    Set objBanco = Nothing
    Set clsTipoFase = Nothing
    Set clsTipoFaseSecundario = Nothing
    Set clsAutenticacao = Nothing
    Set objMsg = Nothing
End Sub
