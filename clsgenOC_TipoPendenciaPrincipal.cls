VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsgenOC_TipoPendenciaPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim objBanco As SGAPLUSBanco.clsBanco
Dim clsgenTipoPendencia As clsgenTipoPendencia
Dim clsTipoPendenciaSecundario As clsgenOC_TipoPendenciaSecundario
Dim clsAutenticacao As clsgenAutenticacao

Dim paTipoPendencia() As Variant

Private Sub Carrega_Grid(ByRef paTipoPendencia() As Variant)
    Dim i As Integer
    
    With frmGenManterTipoPendenciasPrincipal
        .spdTipoPendencia.MaxRows = 0
        
        For i = 1 To UBound(paTipoPendencia, 2)
            If IsEmpty(paTipoPendencia(0, i)) Then Exit For
            
            .spdTipoPendencia.MaxRows = .spdTipoPendencia.MaxRows + 1
            Call .spdTipoPendencia.SetText(1, i, paTipoPendencia(0, i))
            Call .spdTipoPendencia.SetText(2, i, CStr(notisnull(paTipoPendencia(1, i))))
            Call .spdTipoPendencia.SetText(3, i, CStr(paTipoPendencia(2, i)))
        Next
    End With
End Sub

Public Function CarregarForm() As Boolean
    On Error GoTo ErrorHandler
    
    CarregarForm = False
   
    Set clsgenTipoPendencia = New clsgenTipoPendencia
    Set clsAutenticacao = New clsgenAutenticacao

    clsAutenticacao.IDUsuario = frmgenMDI_SGAPlus.objAutenticacao.IDUsuario
    clsAutenticacao.Login = frmgenMDI_SGAPlus.objAutenticacao.Login

    Call clsAutenticacao.InformaBanco(frmgenMDI_SGAPlus.objBanco, objMsg)

    Call clsAutenticacao.RetornarPermissoesFase(clsAutenticacao.IDUsuario, 244, 4, 7, paTipoPendencia, objMsg)
    
    If UBound(paTipoPendencia, 2) = 0 Then
        objMsg.TipoMsg = "E"
        objMsg.TituloMsg = "CarregarForm.CarregarForm"
        objMsg.Descricao = "O usuário não tem direito de consulta"
        
        GoTo ErrorHandler
    End If

    With frmGenManterTipoPendenciasPrincipal
        .Left = 0
        .Height = 7342
        .Width = 11940
        .Top = 0
        .spdTipoPendencia.Lock = True
        
        .CmdAlterar.Enabled = paTipoPendencia(4, 1) = 1
        .CmdCadastrar.Enabled = paTipoPendencia(4, 1) = 1
        .CmdExcluir.Enabled = paTipoPendencia(5, 1) = 1
    End With
   
    Set clsgenTipoPendencia = New clsgenTipoPendencia
    
    Set clsTipoPendenciaSecundario = New clsgenOC_TipoPendenciaSecundario
    
    Call clsgenTipoPendencia.InformaBanco(frmgenMDI_SGAPlus.objBanco, objMsg)
    
    If objMsg.TipoMsg <> "" Then GoTo ErrorHandler
    
    If clsgenTipoPendencia.ListarDados(paTipoPendencia, objMsg) Then
        Call Carrega_Grid(paTipoPendencia)
        CarregarForm = True
    End If
    
    If objMsg.TipoMsg <> "" Then GoTo ErrorHandler
Exit Function

ErrorHandler:
    If objMsg.TipoMsg = "" Then
        objMsg.TipoMsg = "E"
        objMsg.TituloMsg = "clsgenOC_TipoPendenciaPrincipal.CarregarForm"
        objMsg.Descricao = Err.Description
    End If
    objMsg.ExibeMensagem
End Function

Public Function Cadastrar() As Boolean
    Cadastrar = False
    frmGenManterTipoPendenciasSecundario.psTipoGravacao = "I"
    frmGenManterTipoPendenciasSecundario.Show vbModal
    Cadastrar = True
End Function

Public Function Alterar( _
                    ByVal psID As String) As Boolean
    On Error GoTo ErrorHandler
    Alterar = False
    
    Set clsTipoPendenciaSecundario = New clsgenOC_TipoPendenciaSecundario
    
    Call clsTipoPendenciaSecundario.Carrega_Dados_Alteracao(psID)
    
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
        Set clsgenTipoPendencia = New clsgenTipoPendencia
        
        Call clsgenTipoPendencia.InformaBanco(frmgenMDI_SGAPlus.objBanco, objMsg)
        
        If objMsg.TipoMsg <> "" Then GoTo ErrorHandler
        
        If clsgenTipoPendencia.ExcluirDados(psID, objMsg) Then
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
        objMsg.TituloMsg = "clsgenOC_TipoPendenciaPrincipal.CarregarForm"
        objMsg.Descricao = Err.Description
    End If
    objMsg.ExibeMensagem
End Function

Public Function Consultar( _
                    ByVal psID As String) As Boolean
    On Error GoTo ErrorHandler

    Consultar = False
    
    Set clsTipoPendenciaSecundario = New clsgenOC_TipoPendenciaSecundario
    
    Call clsTipoPendenciaSecundario.Carrega_Dados_Alteracao(psID, True)
    
    Consultar = True
Exit Function

ErrorHandler:
    If objMsg.TipoMsg = "" Then
        objMsg.TipoMsg = "E"
        objMsg.TituloMsg = "clsgenOC_TipoPendenciaPrincipal.CarregarForm"
        objMsg.Descricao = Err.Description
    End If
    objMsg.ExibeMensagem
End Function

Private Sub Class_Initialize()
    Set objMsg = New ClsgenMsg
End Sub

Private Sub Class_Terminate()
    Set objBanco = Nothing
    Set clsgenTipoPendencia = Nothing
    Set clsTipoPendenciaSecundario = Nothing
    Set clsAutenticacao = Nothing
    Set objMsg = Nothing
End Sub
