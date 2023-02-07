VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsgenOC_UsuarioSecundario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim objBanco As SGAPLUSBanco.clsBanco
Dim clsgenUsuario As clsgenUsuario

Private lid As Long
Private sLogin As String
Private sNome As String
Private sMAIL As String
Private sRAMAL As String
Private sSENHA As String
Private sMATRICULA As String


Public Property Let ID(ByVal plID As Long)
    lid = plID
End Property

Public Property Get ID() As Long
    ID = lid
End Property

Public Property Let Login(ByVal psLOGIN As String)
    sLogin = psLOGIN
End Property

Public Property Get Login() As String
    Login = sLogin
End Property

Public Property Let Nome(ByVal psNome As String)
    sNome = psNome
End Property

Public Property Get Nome() As String
    Nome = sNome
End Property

Public Property Let MAIL(ByVal psMAIL As String)
    sMAIL = psMAIL
End Property

Public Property Get MAIL() As String
    MAIL = sMAIL
End Property

Public Property Let RAMAL(ByVal psRAMAL As String)
    sRAMAL = psRAMAL
End Property

Public Property Get RAMAL() As String
    RAMAL = sRAMAL
End Property

Public Property Let SENHA(ByVal psSENHA As String)
    sSENHA = psSENHA
End Property

Public Property Get SENHA() As String
    SENHA = sSENHA
End Property

Public Property Let MATRICULA(ByVal psMATRICULA As String)
    sMATRICULA = psMATRICULA
End Property

Public Property Get MATRICULA() As String
    MATRICULA = sMATRICULA
End Property

Public Function GravarDados( _
                            ByVal psTipoGravacao As String) As Boolean

    On Error GoTo ErrorHandler
    
    GravarDados = False
    Set clsgenUsuario = New clsgenUsuario
    
    If clsgenUsuario.InformaBanco(frmgenMDI_SGAPlus.objBanco, objMsg) Then
        If clsgenUsuario.GravarDados(Login, Nome, MAIL, RAMAL, SENHA, MATRICULA, psTipoGravacao, objMsg, ID) Then
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
        objMsg.TituloMsg = "clsgenOC_UsuarioPrincipal.GravarDados"
        objMsg.Descricao = Err.Description
    End If
    objMsg.ExibeMensagem
End Function

Public Function Carrega_Dados_Alteracao( _
                            ByVal psID As String, Optional blConsulta As Boolean) As Boolean
    On Error GoTo ErrorHandler
    
    Dim LocalUsuario() As Variant
    
    Set clsgenUsuario = New clsgenUsuario
    
    Call clsgenUsuario.InformaBanco(frmgenMDI_SGAPlus.objBanco, objMsg)
    
    If objMsg.TipoMsg <> "" Then GoTo ErrorHandler
    
    Call clsgenUsuario.RecuperarDados(psID, LocalUsuario, objMsg)
    
    If objMsg.TipoMsg <> "" Then GoTo ErrorHandler
    
    If Not IsEmpty(LocalUsuario(1, 0)) Then
        frmGenManterUsuarioSecundario.psTipoGravacao = "A"
        frmGenManterUsuarioSecundario.psID = psID
        frmGenManterUsuarioSecundario.txtLogin = notisnull(LocalUsuario(1, 0))
        frmGenManterUsuarioSecundario.TxtNome = notisnull(LocalUsuario(2, 0))
        frmGenManterUsuarioSecundario.txtMail = notisnull(LocalUsuario(3, 0))
        frmGenManterUsuarioSecundario.txtRamal = notisnull(LocalUsuario(4, 0))
        frmGenManterUsuarioSecundario.txtSenha = notisnull(LocalUsuario(5, 0))
        frmGenManterUsuarioSecundario.txtMatricula = notisnull(LocalUsuario(6, 0))
    End If
    
    If blConsulta Then frmGenManterUsuarioSecundario.CmdSalvar.Enabled = False
    
    frmGenManterUsuarioSecundario.Show vbModal

Exit Function

ErrorHandler:
    If objMsg.TipoMsg = "" Then
        objMsg.TipoMsg = "E"
        objMsg.TituloMsg = "clsgenOC_UsuarioSecundario.Carrega_dados_alteracao"
        objMsg.Descricao = Err.Description
    End If
    objMsg.ExibeMensagem
End Function
Private Sub Class_Terminate()
    Set objBanco = Nothing
    Set clsgenUsuario = Nothing
End Sub

Public Function VERIFICAR_USUARIO_EXISTENTE(ByVal psLOGIN As String) As Boolean
    
On Error GoTo ErrorHandler:
    
    Dim objBanco As clsBanco
    Dim bRetorno As Boolean
    
    
Exit Function

ErrorHandler:
    If objMsg.TipoMsg = "" Then
        objMsg.TipoMsg = "E"
        objMsg.TituloMsg = "clsgenOC_UsuarioPrincipal.Excluir"
        objMsg.Descricao = Err.Description
    End If
    objMsg.ExibeMensagem
    
End Function

