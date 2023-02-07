VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsgenOC_OperadoraPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim objBanco As SGAPLUSBanco.clsBanco
Dim clsgenOperadora As clsgenOperadora
Dim clsCentroFuncional As clsgenCentroFuncional
Dim clsOperadoraSecundario As clsGenOC_OperadoraSecundario
Dim clsAutenticacao As clsgenAutenticacao

Dim Operadora() As Variant

Private Sub Carrega_Grid(ByRef Operadora() As Variant)
    Dim i As Integer
    
    With frmGenManterOperadoraPrincipal
        .spdOperadora.MaxRows = 0
        
        For i = 1 To UBound(Operadora, 2)
            If IsEmpty(Operadora(0, i)) Then Exit For
            
            .spdOperadora.MaxRows = .spdOperadora.MaxRows + 1
            Call .spdOperadora.SetText(1, i, Operadora(1, i))
            Call .spdOperadora.SetText(2, i, Operadora(2, i))
            Call .spdOperadora.SetText(3, i, CStr(Operadora(3, i)))
            Call .spdOperadora.SetText(4, i, CStr(Operadora(4, i)))
            Call .spdOperadora.SetText(5, i, CStr(Operadora(0, i)))
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
    
    Call clsAutenticacao.RetornarPermissoesFase(clsAutenticacao.IDUsuario, 243, 4, 7, paPermissoes, objMsg)

    If UBound(paPermissoes, 2) = 0 Then
        objMsg.TipoMsg = "E"
        objMsg.TituloMsg = "CarregarForm.CarregarForm"
        objMsg.Descricao = "O usuário não tem direito de consulta"
        
        GoTo ErrorHandler
    End If
        
    With frmGenManterOperadoraPrincipal
        .Left = 0
        .Height = 7342
        .Width = 11940
        .Top = 0
        
        .CmdAlterar.Enabled = paPermissoes(4, 1) = 1
        .CmdCadastrar.Enabled = paPermissoes(4, 1) = 1
        .CmdExcluir.Enabled = paPermissoes(5, 1) = 1
        
        .spdOperadora.Lock = True
    End With
   
    Set clsgenOperadora = New clsgenOperadora
    Set clsCentroFuncional = New clsgenCentroFuncional

    Call clsgenOperadora.InformaBanco(frmgenMDI_SGAPlus.objBanco, objMsg)
    
    If objMsg.TipoMsg <> "" Then GoTo ErrorHandler
    
    If clsgenOperadora.ListarDados(Operadora, objMsg) Then
        Call Carrega_Grid(Operadora)
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
    Dim i As Long
    
    Cadastrar = False
    
    Dim clsGenRegiao As clsGenRegiao
    Dim clsGenAtuacao As clsGenAtuacao
    
    Set clsGenRegiao = New clsGenRegiao
    Set clsGenAtuacao = New clsGenAtuacao
    
    Dim VetorRegiao() As Variant
    Dim VetorAtuacao() As Variant
    Dim lVetorRegiao() As Variant
    Dim lVetorAtuacao() As Variant
    
    frmGenManterOperadoraSecundario.psTipoGravacao = "I"
    
    Call PreencheCombosForm(frmGenManterOperadoraSecundario, clsGenAtuacao, clsGenRegiao, lVetorAtuacao, lVetorRegiao, VetorAtuacao, VetorRegiao)
    
    If objMsg.TipoMsg <> "" Then GoTo ErrorHandler
    
    frmGenManterOperadoraSecundario.Show vbModal
    
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
                    ByVal psIDOperadora As Long) As Boolean
    On Error GoTo ErrorHandler
    Alterar = False
    
    Set clsOperadoraSecundario = New clsGenOC_OperadoraSecundario
    
    Dim clsGenRegiao As clsGenRegiao
    Dim clsGenAtuacao As clsGenAtuacao
    
    Set clsGenRegiao = New clsGenRegiao
    Set clsGenAtuacao = New clsGenAtuacao
    
    Dim VetorRegiao() As Variant
    Dim VetorAtuacao() As Variant
    Dim lVetorRegiao() As Variant
    Dim lVetorAtuacao() As Variant
    
    Call PreencheCombosForm(frmGenManterOperadoraSecundario, clsGenAtuacao, clsGenRegiao, lVetorAtuacao, lVetorRegiao, VetorAtuacao, VetorRegiao)
    
    Call clsOperadoraSecundario.Carrega_Dados_Alteracao(psIDOperadora)
    
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
                    ByVal psIDOperadora As Long) As Boolean
    On Error GoTo ErrorHandler

    Excluir = False
    
    If MsgBox("Deseja realmente excluir este registro?", vbYesNo + vbDefaultButton2, "Mensagem") = vbYes Then
        Set clsgenOperadora = New clsgenOperadora
        
        Call clsgenOperadora.InformaBanco(frmgenMDI_SGAPlus.objBanco, objMsg)
        
        If objMsg.TipoMsg <> "" Then GoTo ErrorHandler
        
        If clsgenOperadora.ExcluirDados(psIDOperadora, objMsg) Then
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


Private Sub Class_Initialize()
    Set objMsg = New clsgenMensagem
End Sub

Private Sub Class_Terminate()
    Set objBanco = Nothing
    Set clsgenOperadora = Nothing
    Set clsCentroFuncional = Nothing
    
    Set clsOperadoraSecundario = Nothing
    Set clsAutenticacao = Nothing
    Set objMsg = Nothing
End Sub

Private Sub PreencheCombosForm(ByRef frmLocal As Form, ByRef clsGenAtuacao As clsGenAtuacao, ByRef clsGenRegiao As clsGenRegiao, ByRef lVetorAtuacao() As Variant, ByRef lVetorRegiao() As Variant, ByRef VetorAtuacao() As Variant, ByRef VetorRegiao() As Variant)
    Dim i As Integer
    On Error GoTo ErrorHandler
    
    'Recuperando Usuarios
    Call clsGenAtuacao.InformaBanco(frmgenMDI_SGAPlus.objBanco, objMsg)
    
    If objMsg.TipoMsg <> "" Then GoTo ErrorHandler
    
    Call clsGenAtuacao.ListarDados(VetorAtuacao, objMsg)
    
    If objMsg.TipoMsg <> "" Then GoTo ErrorHandler
    
    ReDim lVetorAtuacao(1, UBound(VetorAtuacao, 2))
    
    For i = 0 To UBound(VetorAtuacao, 2)
        lVetorAtuacao(0, i) = VetorAtuacao(0, i)
        lVetorAtuacao(1, i) = VetorAtuacao(1, i)
    Next
    
    Call PreencheCombo(frmLocal.cboAtuacao, lVetorAtuacao)
        
    'Recuperando Centro Funcional
    Call clsGenRegiao.InformaBanco(frmgenMDI_SGAPlus.objBanco, objMsg)
    
    If objMsg.TipoMsg <> "" Then GoTo ErrorHandler
    
    Call clsGenRegiao.ListarDados(VetorRegiao, objMsg)
    
    If objMsg.TipoMsg <> "" Then GoTo ErrorHandler
        
    ReDim lVetorRegiao(1, UBound(VetorRegiao, 2))
    
    For i = 0 To UBound(VetorRegiao, 2)
        lVetorRegiao(0, i) = VetorRegiao(0, i)
        lVetorRegiao(1, i) = VetorRegiao(1, i)
    Next
    
    Call PreencheCombo(frmLocal.cboRegiao, lVetorRegiao)

Exit Sub

ErrorHandler:
    If objMsg.TipoMsg = "" Then
        objMsg.TipoMsg = "E"
        objMsg.TituloMsg = "clsgenOC_TroncoPrincipal.CarregarForm"
        objMsg.Descricao = Err.Description
    End If
    objMsg.ExibeMensagem
    
End Sub
