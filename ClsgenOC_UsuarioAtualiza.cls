VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsgenOC_UsuarioAtualiza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim objBanco As SGAPLUSBanco.clsBanco
Dim clsUsuario As clsgenUsuario
Public Function CarregarForm() As Boolean
    On Error GoTo ErrorHandler
    
    Dim aUsuario() As Variant
    
    CarregarForm = False
  
    Set clsUsuario = New clsgenUsuario

    Call clsUsuario.InformaBanco(frmgenMDI_SGAPlus.objBanco, objMsg)

    If objMsg.TipoMsg <> "" Then GoTo ErrorHandler

    If clsUsuario.RecuperarDados(frmgenMDI_SGAPlus.objAutenticacao.IDUsuario, aUsuario, objMsg) Then
        With frmGenManterUsuarioAtualiza
            frmGenManterUsuarioAtualiza.txtLogin = frmgenMDI_SGAPlus.objAutenticacao.Login
            frmGenManterUsuarioAtualiza.TxtNome.Text = notisnull(aUsuario(2, 0))
            frmGenManterUsuarioAtualiza.txtMail.Text = notisnull(aUsuario(3, 0))
            frmGenManterUsuarioAtualiza.txtRamal.Text = notisnull(aUsuario(4, 0))
            frmGenManterUsuarioAtualiza.txtMatricula.Text = notisnull(aUsuario(6, 0))
            
            .Left = 0
            .Height = 7342
            .Width = 11940
            .Top = 0
        End With
    End If

    If objMsg.TipoMsg <> "" Then GoTo ErrorHandler

    Set clsUsuario = Nothing
    
    CarregarForm = True
Exit Function

ErrorHandler:
    If objMsg.TipoMsg = "" Then
        objMsg.TipoMsg = "E"
        objMsg.TituloMsg = "clsgenOC_UsuarioAtualiza.CarregarForm"
        objMsg.Descricao = Err.Description
    End If
    objMsg.ExibeMensagem
End Function
Public Function Alterar(ByVal psLOGIN As String, _
                                          ByVal psNome As String, _
                                          ByVal psMAIL As String, _
                                          ByVal psRAMAL As String, _
                                          ByVal psSENHANOVA As String, _
                                          ByVal psSENHAConfirma As String, _
                                          ByVal psMatr As String, _
                                          Optional ByVal piID_USUARIO As Long) As Boolean
    On Error GoTo ErrorHandler

    Alterar = False

    If Trim(psSENHANOVA) = "" Then
        objMsg.TipoMsg = "E"
        objMsg.TituloMsg = "clsgenOC_UsuarioAtualiza.Alterar"
        objMsg.Descricao = "Senha inválida"
        GoTo ErrorHandler
    End If

    If psSENHANOVA <> psSENHAConfirma Then
        objMsg.TipoMsg = "E"
        objMsg.TituloMsg = "clsgenOC_UsuarioAtualiza.Alterar"
        objMsg.Descricao = "Confirmação de senha diferente da nova senha"
        GoTo ErrorHandler
    End If
    
    Set clsUsuario = New clsgenUsuario
    
    Call clsUsuario.InformaBanco(frmgenMDI_SGAPlus.objBanco, objMsg)
    
    If objMsg.TipoMsg <> "" Then GoTo ErrorHandler

    Call clsUsuario.GravarDados(frmgenMDI_SGAPlus.objAutenticacao.Login, psNome, psMAIL, psRAMAL, psSENHANOVA, psMatr, "A", _
        objMsg, frmgenMDI_SGAPlus.objAutenticacao.IDUsuario)

    If objMsg.TipoMsg <> "" Then GoTo ErrorHandler

    objMsg.TipoMsg = "A"
    objMsg.TituloMsg = "Mensagem"
    objMsg.Descricao = "Operação efetuada com sucesso!"
    objMsg.ExibeMensagem

    Alterar = True

    Exit Function
ErrorHandler:
    If objMsg.TipoMsg = "" Then
        objMsg.TipoMsg = "E"
        objMsg.TituloMsg = "clsgenOC_UsuarioAtualiza.Alterar"
        objMsg.Descricao = Err.Description & objBanco.objErro(0).Description
    End If
    objMsg.ExibeMensagem
End Function
Private Sub Class_Initialize()
    Set objMsg = New ClsgenMsg
End Sub

Private Sub Class_Terminate()
    Set objBanco = Nothing
    Set clsUsuario = Nothing
    Set objMsg = Nothing
End Sub
