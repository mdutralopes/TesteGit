VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsgenOC_CFSecundario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim objBanco As SGAPLUSBanco.clsBanco
Dim clsgenCF As clsgenCentroFuncional

Private lid As Long
Private sCodigo As String
Private sDescricao As String

Public Function GravarDados( _
                            ByVal psTipoGravacao As String) As Boolean

    On Error GoTo ErrorHandler
    
    GravarDados = False
    Set clsgenCF = New clsgenCentroFuncional
    
    If clsgenCF.InformaBanco(frmgenMDI_SGAPlus.objBanco, objMsg) Then
        If clsgenCF.GravarDados(Codigo, Descricao, psTipoGravacao, objMsg, ID) Then
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
                            ByVal psID As String, Optional blConsulta As Boolean) As Boolean

    On Error GoTo ErrorHandler
    
    Dim LocalCF() As Variant
    
    Set clsgenCF = New clsgenCentroFuncional
    
    Call clsgenCF.InformaBanco(frmgenMDI_SGAPlus.objBanco, objMsg)
    
    If objMsg.TipoMsg <> "" Then GoTo ErrorHandler
    
    Call clsgenCF.RecuperarDados(psID, LocalCF, objMsg)
    
    If objMsg.TipoMsg <> "" Then GoTo ErrorHandler
    
    If Not IsEmpty(LocalCF(1, 1)) Then
        frmGenManterCFSecundario.psTipoGravacao = "A"
        frmGenManterCFSecundario.psID = psID
        frmGenManterCFSecundario.txtCodigo = LocalCF(1, 1)
        frmGenManterCFSecundario.txtDescricao = LocalCF(2, 1)
    End If
    
    If blConsulta Then frmGenManterCFSecundario.CmdSalvar.Enabled = False
    
    frmGenManterCFSecundario.Show vbModal

Exit Function

ErrorHandler:
    If objMsg.TipoMsg = "" Then
        objMsg.TipoMsg = "E"
        objMsg.TituloMsg = "clsgenOC_TroncoPrincipal.CarregarForm"
        objMsg.Descricao = Err.Description
    End If
    objMsg.ExibeMensagem
End Function

'sSQL = sSQL & " ID,"
Public Property Let ID(ByVal plID As Long)
    lid = plID
End Property

Public Property Get ID() As Long
    ID = lid
End Property

Public Property Let Codigo(ByVal psCodigo As String)
    sCodigo = psCodigo
End Property

Public Property Get Codigo() As String
    Codigo = sCodigo
End Property

Public Property Let Descricao(ByVal psDescricao As String)
    sDescricao = psDescricao
End Property

Public Property Get Descricao() As String
    Descricao = sDescricao
End Property

Private Sub Class_Terminate()
    Set objBanco = Nothing
    Set clsgenCF = Nothing
End Sub
