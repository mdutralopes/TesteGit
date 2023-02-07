VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsgenOC_UsuarioPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim objBanco As SGAPLUSBanco.clsBanco
Dim clsgenUsuario As clsgenUsuario
Dim clsUsuarioSecundario As clsgenOC_UsuarioSecundario
Dim clsgenAutenticacao As clsgenAutenticacao

Dim Usuario() As Variant

Public Function CarregarForm() As Boolean
    On Error GoTo ErrorHandler
    
    Dim paPermissoes() As Variant
        
    CarregarForm = False
  
    Set clsgenAutenticacao = New clsgenAutenticacao
    
    clsgenAutenticacao.IDUsuario = frmgenMDI_SGAPlus.objAutenticacao.IDUsuario
    clsgenAutenticacao.Login = frmgenMDI_SGAPlus.objAutenticacao.Login
    Call clsgenAutenticacao.InformaBanco(frmgenMDI_SGAPlus.objBanco, objMsg)
    
    Call clsgenAutenticacao.RetornarPermissoesFase(clsgenAutenticacao.IDUsuario, 247, 4, 7, paPermissoes, objMsg)
'    Call clsgenAutenticacao.RetornarPermissoesFase(clsgenAutenticacao.IDUsuario, 62, paPermissoes, objMsg)

    If UBound(paPermissoes, 2) = 0 Then
        objMsg.TipoMsg = "E"
        objMsg.TituloMsg = "CarregarForm.CarregarForm"
        objMsg.Descricao = "O usuário não tem direito de consulta"
        
        GoTo ErrorHandler
    End If
            
    With frmGenManterUsuarioPrincipal
        .Left = 0
        .Height = 7342
        .Width = 11940
        .Top = 0
        .spdUsuario.Lock = True
       
        .CmdAlterar.Enabled = paPermissoes(4, 1) = 1
        .CmdCadastrar.Enabled = paPermissoes(4, 1) = 1
        .CmdExcluir.Enabled = paPermissoes(5, 1) = 1
    End With

    Set clsgenUsuario = New clsgenUsuario
    Set clsgenAutenticacao = New clsgenAutenticacao
        
    Call clsgenUsuario.InformaBanco(frmgenMDI_SGAPlus.objBanco, objMsg)
    
    If objMsg.TipoMsg <> "" Then GoTo ErrorHandler
    
    If clsgenUsuario.ListarDados(Usuario(), objMsg) Then
        Call Carrega_Grid(Usuario())
        CarregarForm = True
    End If
    
    If objMsg.TipoMsg <> "" Then GoTo ErrorHandler

Exit Function

ErrorHandler:
    If objMsg.TipoMsg = "" Then
        objMsg.TipoMsg = "E"
        objMsg.TituloMsg = "clsgenOC_UsuarioPrincipal.CarregarForm"
        objMsg.Descricao = Err.Description
    End If
    objMsg.ExibeMensagem
End Function

Private Sub Carrega_Grid(ByRef Usuario() As Variant)
On Error GoTo ErrorHandler:
    Dim i As Integer
    
    With frmGenManterUsuarioPrincipal
        .spdUsuario.MaxRows = 0
        
        For i = 1 To UBound(Usuario, 2)
            If IsEmpty(Usuario(0, i)) Then Exit For
            
            .spdUsuario.MaxRows = .spdUsuario.MaxRows + 1
            Call .spdUsuario.SetText(1, i, CStr(Usuario(0, i)))
            Call .spdUsuario.SetText(2, i, notisnull(Usuario(1, i)))
            Call .spdUsuario.SetText(3, i, notisnull(Usuario(2, i)))
            Call .spdUsuario.SetText(4, i, notisnull(Usuario(3, i)))
            Call .spdUsuario.SetText(5, i, notisnull(Usuario(4, i)))
            Call .spdUsuario.SetText(6, i, notisnull(Usuario(6, i)))
        Next
    End With
    
    Exit Sub
ErrorHandler:
    If objMsg.TipoMsg = "" Then
        objMsg.TipoMsg = "E"
        objMsg.TituloMsg = "clsgenOC_UsuarioPrincipal.CarregarGrid"
        objMsg.Descricao = Err.Description
    End If
    objMsg.ExibeMensagem
End Sub

Public Function Cadastrar() As Boolean
On Error GoTo ErrorHandler:
    Cadastrar = False
    frmGenManterUsuarioSecundario.psTipoGravacao = "I"
    frmGenManterUsuarioSecundario.Show vbModal
    Cadastrar = True
Exit Function

ErrorHandler:
    If objMsg.TipoMsg = "" Then
        objMsg.TipoMsg = "E"
        objMsg.TituloMsg = "clsgenOC_UsuarioPrincipal.Cadastrar"
        objMsg.Descricao = Err.Description
    End If
    objMsg.ExibeMensagem
End Function

Private Sub Class_Initialize()
    Set objMsg = New ClsgenMsg
End Sub

Private Sub Class_Terminate()
    Set objBanco = Nothing
    Set clsgenUsuario = Nothing
    Set clsUsuarioSecundario = Nothing
    Set clsgenAutenticacao = Nothing
    Set objMsg = Nothing
End Sub

Public Function Alterar( _
                    ByVal psID As String) As Boolean
On Error GoTo ErrorHandler:

    Alterar = False
    
    Set clsUsuarioSecundario = New clsgenOC_UsuarioSecundario
        
    Call clsUsuarioSecundario.Carrega_Dados_Alteracao(psID)
    
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
        Set clsgenUsuario = New clsgenUsuario
        
        Call clsgenUsuario.InformaBanco(frmgenMDI_SGAPlus.objBanco, objMsg)
        
        If objMsg.TipoMsg <> "" Then GoTo ErrorHandler
        
        If clsgenUsuario.ExcluirDados(psID, objMsg) Then
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
        objMsg.TituloMsg = "clsgenOC_UsuarioPrincipal.Excluir"
        objMsg.Descricao = Err.Description
    End If
    objMsg.ExibeMensagem
End Function
