VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsgenOC_PerfilCF_Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim objBanco As SGAPLUSBanco.clsBanco

Private objMensagem                 As clsgenMensagem
Private objCadastrarUsuario_CF      As clsgenUsuario_CF

'Dim ClsPerfilCF         As ClsgenPerfilCF
Dim clsCentroFuncional  As clsgenCentroFuncional

Dim clsAutenticacao As clsgenAutenticacao
    
Dim paPermissoes() As Variant
Dim paDadosPerfilCF() As Variant

Private Sub Carrega_Grid(ByRef paDadosTipoFase() As Variant)
    Dim i As Integer
    
    With frmGenManterPefilCFPrincipal
        .SpdPerfil.MaxRows = 0
        
        For i = 1 To UBound(paDadosTipoFase, 2)
            If IsEmpty(paDadosTipoFase(0, i)) Then Exit For
            
            .SpdPerfil.MaxRows = .SpdPerfil.MaxRows + 1
            Call .SpdPerfil.SetText(1, i, paDadosTipoFase(0, i))
            Call .SpdPerfil.SetText(2, i, paDadosTipoFase(1, i))
            Call .SpdPerfil.SetText(3, i, paDadosTipoFase(2, i))
            Call .SpdPerfil.SetText(4, i, CStr(paDadosTipoFase(3, i)))
            Call .SpdPerfil.SetText(5, i, CStr(paDadosTipoFase(4, i)))
            Call .SpdPerfil.SetText(6, i, CStr(paDadosTipoFase(5, i)))
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
    
    Call clsAutenticacao.RetornarPermissoesFase(clsAutenticacao.IDUsuario, 221, paPermissoes, objMsg)

    If UBound(paPermissoes, 2) = 0 Then
        objMsg.TipoMsg = "E"
        objMsg.TituloMsg = "ClsgenOC_PerfilCF_Principal.CarregarForm"
        objMsg.Descricao = "O usuário não tem direito de consulta"
        GoTo ErrorHandler
    End If
        
    With frmGenManterPefilCFPrincipal
        .Left = 0
        .Height = 7342
        .Width = 11940
        .Top = 0
        '.SpdPerfil.Refresh
        '.SpdPerfil.Lock = True

        .CmdAlterar.Enabled = paPermissoes(4, 1) = 1
        .CmdCadastrar.Enabled = paPermissoes(4, 1) = 1
        .CmdExcluir.Enabled = paPermissoes(5, 1) = 1
    End With
    

    Set ClsPerfilCF = New ClsgenPerfilCF

    Call ClsPerfilCF.InformaBanco(frmgenMDI_SGAPlus.objBanco, objMsg)

    If objMsg.TipoMsg <> "" Then GoTo ErrorHandler

    If ClsPerfilCF.ListarDados(paDadosPerfilCF, objMsg) Then
        Call Carrega_Grid(paDadosPerfilCF)
        CarregarForm = True
    End If

    If objMsg.TipoMsg <> "" Then GoTo ErrorHandler

    Call Carrega_Grid(paDadosPerfilCF)
    
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

    objMsg.TipoMsg = "A"
    objMsg.TituloMsg = "Mensagem"
    objMsg.Descricao = "Alteração efetuada com sucesso!"
    objMsg.ExibeMensagem

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
        Set clsTipoFase = New clsgenTipoFase
        
        Call clsTipoFase.InformaBanco(frmgenMDI_SGAPlus.objBanco, objMsg)
        
        If objMsg.TipoMsg <> "" Then GoTo ErrorHandler
        
        If clsTipoFase.ExcluirDados(psID, objMsg) Then
            Excluir = True
        End If
    
        If objMsg.TipoMsg <> "" Then GoTo ErrorHandler
    End If
Exit Function

ErrorHandler:
    If objMsg.TipoMsg = "" Then
        objMsg.TipoMsg = "E"
        objMsg.TituloMsg = "clsgenOC_TipoFasePrincipal.Excluir"
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
