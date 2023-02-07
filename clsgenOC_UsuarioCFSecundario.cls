VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsgenOC_UsuarioCFSecundario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim objBanco As SGAPLUSBanco.clsBanco
Dim clsgenUsuarioCF As clsgenUsuarioXCentroFuncional

Dim lIDUsuario As Long
Dim lIDCF As Long
Dim sUsuario As String
Dim sCF As String
Dim sCF_OLD  As Long

Public Function GravarDados( _
                            ByVal psTipoGravacao As String) As Boolean

    On Error GoTo ErrorHandler
    
    GravarDados = False
    Set clsgenUsuarioCF = New clsgenUsuarioXCentroFuncional
        
    Dim vVerifica() As Variant
    
    If clsgenUsuarioCF.InformaBanco(frmgenMDI_SGAPlus.objBanco, objMsg) Then
        If clsgenUsuarioCF.RecuperarDados(IDUsuario, IDCF, vVerifica, objMsg) Then
            If UBound(vVerifica) = 0 Then
                If clsgenUsuarioCF.GravarDados(psTipoGravacao, IDUsuario, IDCF, CF_OLD, objMsg) Then
                    GravarDados = True
                End If
            Else
                objMsg.TipoMsg = "E"
                objMsg.TituloMsg = "clsgenOC_UsuarioCFPrincipal.CarregarForm"
                objMsg.Descricao = "Centro Funcional já cadastrado!"
                objMsg.ExibeMensagem
            End If
        Else
            If objMsg.TipoMsg <> "" Then GoTo ErrorHandler
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
                            ByVal psIDUsuario As Long, ByVal psIDCF As Long, Optional blConsulta As Boolean) As Boolean

    On Error GoTo ErrorHandler
    
    Dim LocalUsuarioCF() As Variant
    Dim i As Integer
    
    Set clsgenUsuarioCF = New clsgenUsuarioXCentroFuncional
    
    Call clsgenUsuarioCF.InformaBanco(frmgenMDI_SGAPlus.objBanco, objMsg)
    
    If objMsg.TipoMsg <> "" Then GoTo ErrorHandler
    
    Call clsgenUsuarioCF.RecuperarDados(psIDUsuario, psIDCF, LocalUsuarioCF, objMsg)
    
    If objMsg.TipoMsg <> "" Then GoTo ErrorHandler
    
    If Not IsEmpty(LocalUsuarioCF(1, 0)) Then
        frmGenManterUsuarioCFSecundario.psTipoGravacao = "A"
        
        'Posicionando no registro selecionado
        For i = 0 To frmGenManterUsuarioCFSecundario.cboLogin.ListCount - 1
            If frmGenManterUsuarioCFSecundario.cboLogin.ItemData(i) = psIDUsuario Then
                frmGenManterUsuarioCFSecundario.cboLogin.ListIndex = i
                Exit For
            End If
        Next
        For i = 0 To frmGenManterUsuarioCFSecundario.cboCodigo.ListCount - 1
            If frmGenManterUsuarioCFSecundario.cboCodigo.ItemData(i) = psIDCF Then
                frmGenManterUsuarioCFSecundario.cboCodigo.ListIndex = i
                Exit For
            End If
        Next
        
        frmGenManterUsuarioCFSecundario.psIDUsuario = psIDUsuario
        frmGenManterUsuarioCFSecundario.psIDCF = psIDCF
    End If
    
    If blConsulta Then frmGenManterUsuarioCFSecundario.CmdSalvar.Enabled = False
    
    frmGenManterUsuarioCFSecundario.cboLogin.Enabled = False
    
    frmGenManterUsuarioCFSecundario.Show vbModal

Exit Function

ErrorHandler:
    If objMsg.TipoMsg = "" Then
        objMsg.TipoMsg = "E"
        objMsg.TituloMsg = "clsgenOC_UsuarioXCentroFuncional.CarregarForm"
        objMsg.Descricao = Err.Description
    End If
    objMsg.ExibeMensagem
End Function


Public Property Let IDUsuario(ByVal plIDUsuario As Long)
    lIDUsuario = plIDUsuario
End Property

Public Property Get IDUsuario() As Long
    IDUsuario = lIDUsuario
End Property

Public Property Let IDCF(ByVal pLIDCF As Long)
    lIDCF = pLIDCF
End Property

Public Property Get IDCF() As Long
    IDCF = lIDCF
End Property

Public Property Let Usuario(ByVal psUsuario As String)
    sUsuario = psUsuario
End Property

Public Property Get Usuario() As String
    Usuario = sUsuario
End Property

Public Property Let CF(ByVal psCF As String)
    sCF = psCF
End Property

Public Property Get CF() As String
    CF = sCF
End Property

Public Property Let CF_OLD(ByVal psCF_OLD As Long)
    sCF_OLD = psCF_OLD
End Property

Public Property Get CF_OLD() As Long
    CF_OLD = sCF_OLD
End Property

Private Sub Class_Terminate()
    Set objBanco = Nothing
    Set clsgenUsuarioCF = Nothing
End Sub
