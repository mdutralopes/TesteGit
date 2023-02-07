VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsGenOC_MotivoRecusaSecundario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim objBanco As SGAPLUSBanco.clsBanco
Dim ClsGenMotivoRecusa As ClsGenMotivoRecusa

Private lid As Long
Private sDescricao As String

Public Property Let ID(ByVal plID As Long)
    lid = plID
End Property

Public Property Get ID() As Long
    ID = lid
End Property
Public Property Let Descricao(ByVal psDescricao As String)
    sDescricao = psDescricao
End Property

Public Property Get Descricao() As String
    Descricao = sDescricao
End Property

Public Function GravarDados( _
                            ByVal psTipoGravacao As String) As Boolean

    On Error GoTo ErrorHandler
    
    GravarDados = False
    Set ClsGenMotivoRecusa = New ClsGenMotivoRecusa
    
    If ClsGenMotivoRecusa.InformaBanco(frmgenMDI_SGAPlus.objBanco, objMsg) Then
        If ClsGenMotivoRecusa.GravarDados(psTipoGravacao, Descricao, lid) Then
            GravarDados = True
        End If
    Else
        If objMsg.TipoMsg <> "" Then GoTo ErrorHandler
    End If
    
    If objMsg.TipoMsg <> "" Then GoTo ErrorHandler
    
    objMsg.TipoMsg = "A"
    objMsg.TituloMsg = "Mensagem"
    objMsg.Descricao = "Operação efetuada com sucesso!"
    objMsg.ExibeMensagem

Exit Function

ErrorHandler:
    If objMsg.TipoMsg = "" Then
        objMsg.TipoMsg = "E"
        objMsg.TituloMsg = "clsgenMotivoRecusa.CarregarForm"
        objMsg.Descricao = Err.Description
    End If
    objMsg.ExibeMensagem
End Function


Public Function Carrega_Dados_Alteracao( _
                            ByVal psID As String, Optional blConsulta As Boolean) As Boolean

    On Error GoTo ErrorHandler
    
    Dim LocalMotivo() As Variant
    
    Set ClsGenMotivoRecusa = New ClsGenMotivoRecusa
    
    Call ClsGenMotivoRecusa.InformaBanco(frmgenMDI_SGAPlus.objBanco, objMsg)
    
    If objMsg.TipoMsg <> "" Then GoTo ErrorHandler
    
    Call ClsGenMotivoRecusa.RecuperarDados(psID, LocalMotivo())
    
    If objMsg.TipoMsg <> "" Then GoTo ErrorHandler
    
    If Not IsEmpty(LocalMotivo(0, 0)) Then
        frmGenManterMotivoRecusaSecundario.psTipoGravacao = "A"
        frmGenManterMotivoRecusaSecundario.psID = psID
        frmGenManterMotivoRecusaSecundario.txtMotivo = LocalMotivo(0, 0)
    End If
    
    If blConsulta Then frmGenManterMotivoRecusaSecundario.CmdSalvar.Enabled = False
    
    frmGenManterMotivoRecusaSecundario.Show vbModal

Exit Function

ErrorHandler:
    If objMsg.TipoMsg = "" Then
        objMsg.TipoMsg = "E"
        objMsg.TituloMsg = "clsgenOC_MotivoRecusaPrincipal.CarregarForm"
        objMsg.Descricao = Err.Description
    End If
    objMsg.ExibeMensagem
End Function

Private Sub Class_Terminate()
    Set objBanco = Nothing
    Set ClsGenMotivoRecusa = Nothing
End Sub
