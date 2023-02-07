VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsgenOC_MotivoRecusaPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim objBanco As SGAPLUSBanco.clsBanco
Dim ClsGenMotivoRecusa As ClsGenMotivoRecusa
Dim clsCentroFuncional As clsgenCentroFuncional
Dim clsMotivoRecusaSecundario As ClsGenOC_MotivoRecusaSecundario
Dim clsAutenticacao As clsgenAutenticacao

Dim Motivo() As Variant
Dim paPermissoes() As Variant

Public Function CarregarForm() As Boolean
    On Error GoTo ErrorHandler
    
    CarregarForm = False
   
    Set ClsGenMotivoRecusa = New ClsGenMotivoRecusa
    Set clsCentroFuncional = New clsgenCentroFuncional
    Set clsAutenticacao = New clsgenAutenticacao
    
    Call ClsGenMotivoRecusa.InformaBanco(frmgenMDI_SGAPlus.objBanco, objMsg)
    
    clsAutenticacao.IDUsuario = frmgenMDI_SGAPlus.objAutenticacao.IDUsuario
    clsAutenticacao.Login = frmgenMDI_SGAPlus.objAutenticacao.Login
    Call clsAutenticacao.InformaBanco(frmgenMDI_SGAPlus.objBanco, objMsg)
    
    Call clsAutenticacao.RetornarPermissoesFase(clsAutenticacao.IDUsuario, 242, 4, 7, paPermissoes, objMsg)

    If UBound(paPermissoes, 2) = 0 Then
        objMsg.TipoMsg = "E"
        objMsg.TituloMsg = "CarregarForm.CarregarForm"
        objMsg.Descricao = "O usuário não tem direito de consulta"
        
        Set ClsGenMotivoRecusa = Nothing
        
        GoTo ErrorHandler
    End If
        
     With frmGenManterMotivoRecusaPrincipal
        .CmdAlterar.Enabled = paPermissoes(4, 1) = 1
        .CmdCadastrar.Enabled = paPermissoes(4, 1) = 1
        .CmdExcluir.Enabled = paPermissoes(5, 1) = 1
        .spdMotivo.Lock = True
    End With
    
    If ClsGenMotivoRecusa.ListarDados(Motivo()) Then
        Call Carrega_Grid(Motivo())
        CarregarForm = True
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
    Set objMsg = New ClsgenMsg
End Sub

Private Sub Class_Terminate()
    Set objBanco = Nothing
    Set ClsGenMotivoRecusa = Nothing
    Set clsCentroFuncional = Nothing
    Set clsMotivoRecusaSecundario = Nothing
    Set clsAutenticacao = Nothing
    Set objMsg = Nothing
End Sub

Public Function Cadastrar() As Boolean
    Cadastrar = False
    
    With frmGenManterMotivoRecusaSecundario
          .psTipoGravacao = "I"
          .Show vbModal
     End With
    Cadastrar = True
End Function

Public Function Alterar( _
                    ByVal psID As String) As Boolean
    On Error GoTo ErrorHandler
    Alterar = False
    
    Set clsMotivoRecusaSecundario = New ClsGenOC_MotivoRecusaSecundario
    
    Call clsMotivoRecusaSecundario.Carrega_Dados_Alteracao(psID)
    
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
    Excluir = False
    
    If MsgBox("Deseja realmente excluir este registro?", vbYesNo + vbDefaultButton2, "Mensagem") = vbYes Then
        Set ClsGenMotivoRecusa = New ClsGenMotivoRecusa
        
        Call ClsGenMotivoRecusa.InformaBanco(frmgenMDI_SGAPlus.objBanco, objMsg)
        If ClsGenMotivoRecusa.ExcluirDados(psID) Then
            Excluir = True
            
            objMsg.TipoMsg = "A"
            objMsg.TituloMsg = "Mensagem"
            objMsg.Descricao = "Operação efetuada com sucesso!"
            objMsg.ExibeMensagem
        End If
        
    End If
End Function

Private Sub Carrega_Grid(ByRef pMotivo() As Variant)
    Dim i As Integer
    
    With frmGenManterMotivoRecusaPrincipal
          .spdMotivo.MaxRows = 0
          .spdMotivo.MaxCols = 2
    
          For i = 1 To UBound(pMotivo, 2)
               If IsEmpty(pMotivo(0, i - 1)) Then Exit For
                    .spdMotivo.MaxRows = frmGenManterMotivoRecusaPrincipal.spdMotivo.MaxRows + 1
               
                    Call .spdMotivo.SetText(1, i, pMotivo(1, i - 1))            'Descrição do Motivo.
                    Call .spdMotivo.SetText(2, i, CStr(pMotivo(0, i - 1)))            'ID do Motivo.
          Next
     End With
End Sub

