VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsgenOC_TipoPendenciaSecundario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim objBanco As SGAPLUSBanco.clsBanco
Dim clsgenTipoPendencia As clsgenTipoPendencia

Private lid As Long
Private sDescricao As String
Private sCodigoCOntrole As String

Public Function GravarDados(ByVal psTipoGravacao As String) As Boolean

    On Error GoTo ErrorHandler
    
    GravarDados = False
    Set clsgenTipoPendencia = New clsgenTipoPendencia
    
    If clsgenTipoPendencia.InformaBanco(frmgenMDI_SGAPlus.objBanco, objMsg) Then
        If clsgenTipoPendencia.GravarDados(sDescricao, sCodigoCOntrole, psTipoGravacao, objMsg, ID) Then
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
        objMsg.TituloMsg = "clsgenOC_TroncoPrincipal.CarregarForm"
        objMsg.Descricao = Err.Description
    End If
    objMsg.ExibeMensagem
End Function

Public Function Carrega_Dados_Alteracao( _
                            ByVal plID As Long, Optional blConsulta As Boolean) As Boolean

    On Error GoTo ErrorHandler
    
    Dim aTipoPendencias() As Variant
    
    Set clsgenTipoPendencia = New clsgenTipoPendencia
    
    Call clsgenTipoPendencia.InformaBanco(frmgenMDI_SGAPlus.objBanco, objMsg)
    
    If objMsg.TipoMsg <> "" Then GoTo ErrorHandler
    
    Call clsgenTipoPendencia.RecuperarDados(plID, aTipoPendencias, objMsg)
    
    If objMsg.TipoMsg <> "" Then GoTo ErrorHandler
    
    If Not IsEmpty(aTipoPendencias(1, 0)) Then
        With frmGenManterTipoPendenciasSecundario
            .psTipoGravacao = "A"
            .txtDescricao = aTipoPendencias(0, 0)
            .TxtControle = notisnull(aTipoPendencias(1, 0))
            .plID = plID
        End With
    End If
    
    If blConsulta Then frmGenManterTipoPendenciasSecundario.CmdSalvar.Enabled = False
    
    frmGenManterTipoPendenciasSecundario.Show vbModal

Exit Function

ErrorHandler:
    If objMsg.TipoMsg = "" Then
        objMsg.TipoMsg = "E"
        objMsg.TituloMsg = "clsgenOC_TipoPendenciasPrincipal.CarregarForm"
        objMsg.Descricao = Err.Description
    End If
    objMsg.ExibeMensagem
End Function

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


Public Property Let CodigoControle(ByVal psCodControle As String)
    sCodigoCOntrole = psCodControle
End Property

Public Property Get CodigoControle() As String
    CodigoControle = sCodigoCOntrole
End Property

Private Sub Class_Terminate()
    Set objBanco = Nothing
    Set clsgenTipoPendencia = Nothing
End Sub

