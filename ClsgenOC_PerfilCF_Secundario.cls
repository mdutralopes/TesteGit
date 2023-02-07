VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsgenOC_CadastrarPerfil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim objBanco As SGAPLUSBanco.clsBanco
Dim clsTipoFase As clsgenTipoFase

Private lid As Long
Private psServico As String
Private psAcao As String
Private psNome As String
Private piOrdem As Integer
Private piDuracao As Integer

Public Function GravarDados( _
                            ByVal psTipoGravacao As String) As Boolean
    On Error GoTo ErrorHandler
    
    GravarDados = False
    
    Set clsTipoFase = New clsgenTipoFase
    
    If clsTipoFase.InformaBanco(frmgenMDI_SGAPlus.objBanco, objMsg) Then
        If clsTipoFase.GravarDados(psTipoGravacao, psServico, psAcao, psNome, piOrdem, piDuracao, objMsg, ID) Then
            GravarDados = True
        End If
    Else
        If objMsg.TipoMsg <> "" Then GoTo ErrorHandler
    End If
    
    If objMsg.TipoMsg <> "" Then GoTo ErrorHandler

Exit Function

ErrorHandler:
    If objMsg.TipoMsg = "" Then
        objMsg.TipoMsg = "E"
        objMsg.TituloMsg = "clsgenOC_PerfilCFPrincipal.GravarDados"
        objMsg.Descricao = Err.Description & objBanco.objErro(0).Description
    End If
    objMsg.ExibeMensagem
End Function

Public Function Carrega_Dados_Alteracao( _
                            ByVal psID As Long, Optional blConsulta As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim LocalTipoFase() As Variant
    Dim iServico As Integer
    
    Set ClsPerfilCF = New clsgenTipoFase
    
    Call ClsPerfilCF.InformaBanco(frmgenMDI_SGAPlus.objBanco, objMsg)
    
    If objMsg.TipoMsg <> "" Then GoTo ErrorHandler

    Call PreencheComboServico(frmGenManterPefilCFSecundario.cboServico)
    
    Call PreencheComboAcao(frmGenManterPefilCFSecundario.cboAcao)
    
    Call PreencheLstFonte(frmGenManterPefilCFSecundario.lstFonte)
    
    Call ClsPerfilCF.RecuperarDados(psID, LocalTipoFase, objMsg)
    
    If objMsg.TipoMsg <> "" Then GoTo ErrorHandler
    
    If Not IsEmpty(LocalTipoFase(1, 0)) Then
        With frmGenManterPerfilCFSecundario
            .psTipoGravacao = "A"
            
            If LocalTipoFase(0, 0) = "DIGINET" Then
                iServico = 0
            ElseIf LocalTipoFase(0, 0) = "WID" Then
                iServico = 1
            ElseIf LocalTipoFase(0, 0) = "BLFLEX" Then
                iServico = 2
            End If
            '.cboServic.ListIndex(.cboServico.ItemData) = iServico
            frmGenManterPerfilCFSecundario.cboServico.ListIndex = (frmGenManterPerfilCFSecundario.cboServico.ItemData(iServico))
            .TxtAcao = LocalTipoFase(1, 0)
            .txtFase = LocalTipoFase(2, 0)
            .TxtOrdem = LocalTipoFase(3, 0)
            .txtDuracao = LocalTipoFase(4, 0)
            .psID = psID
            
        End With
    End If
    frmGenManterPerfilCFSecundario.Show vbModal
Exit Function

ErrorHandler:
    If objMsg.TipoMsg = "" Then
        objMsg.TipoMsg = "E"
        objMsg.TituloMsg = "clsgenOC_PerfilCFPrincipal.Carrega_Dados_Alteracao"
    End If
    objMsg.ExibeMensagem
End Function

Public Function PreencheComboAcao(combo As ComboBox)
        'Exemplo de chamada
        'Call PreencheComboServico(frmGenManterTiposFasesSecundario.cboServico)
        
    Dim paAcao As Variant
    Dim sSQL As String
    Dim iIndice As Integer
    
    sSQL = "select  descricao, id " & Chr(13)
    sSQL = sSQL & "from    acao " & Chr(13)
    
    bRetorno = objBanco.Obter_Selecao(sSQL, paAcao, 0, bExcedentes, False, "", "")
    
    If Not bRetorno Then GoTo ErrorHandler
    
    For iIndice = 0 To UBound(paAcao)
        combo.AddItem paAcao(iIndice, 0), iIndice
    Next iIndice
    
    PreencheComboAcao = True
        
ErrorHandler:
    If objMsg.TipoMsg = "" Then
        objMsg.TipoMsg = "E"
        objMsg.TituloMsg = "clsgenOC_PerfilCFPrincipal.Carrega_Dados_Alteracao"
    End If
    objMsg.ExibeMensagem
'        Call PreencheCombo(combo, aServico)
End Function

Public Function PreencheLstFonte(List As ListBox)
    
    Dim sSQL As String
    Dim paCFunc As Variant
    
    sSQL = "       select  codigo, id " & Chr(13)
    sSQL = sSQL & "from    centro_funcional " & Chr(13)
    
    bRetorno = objBanco.Obter_Selecao(sSQL, paCFunc, 0, bExcedentes, False, "", "")
    
    If Not bRetorno Then GoTo ErrorHandler
    
    For iIndice = 0 To UBound(paCFunc)
        List.AddItem paCFunc(iIndice, 0), iIndice
    Next iIndice
    
    PreencheLstFonte = True
    
ErrorHandler:
    If objMsg.TipoMsg = "" Then
        objMsg.TipoMsg = "E"
        objMsg.TituloMsg = "clsgenOC_PerfilCFPrincipal.Carrega_Dados_Alteracao"
    End If
    objMsg.ExibeMensagem
End Function

Public Property Let ID(ByVal vlid As Long)
    lid = vlid
End Property

Public Property Get ID() As Long
    ID = lid
End Property
Public Property Let Servico(ByVal sServico As String)
    psServico = sServico
End Property

Public Property Get Servico() As String
    Servico = psServico
End Property

Public Property Let Acao(ByVal sAcao As String)
    psAcao = sAcao
End Property

Public Property Get Acao() As String
    Acao = psAcao
End Property

Public Property Get Nome() As String
    Nome = psNome
End Property

Public Property Let Nome(ByVal sNome As String)
    psNome = sNome
End Property

Public Property Get Ordem() As Integer
    Ordem = piOrdem
End Property

Public Property Let Ordem(ByVal iOrdem As Integer)
    piOrdem = iOrdem
End Property

Public Property Get Duracao() As Integer
    Duracao = piDuracao
End Property

Public Property Let Duracao(ByVal iDuracao As Integer)
    piDuracao = iDuracao
End Property

Private Sub Class_Terminate()
    Set objBanco = Nothing
    Set ClsPerfilCF = Nothing
End Sub
