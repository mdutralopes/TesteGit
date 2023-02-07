VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGenOC_OperadoraSecundario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
Option Explicit

' Classes usados neste controle
Private objBanco As clsBanco                    'As SGAPLUSBanco.clsBanco
Private objMsg As New clsgenMensagem
Private clsOperadora As clsgenOperadora
Private clsRegiao As clsGenRegiao
Private clsAtuacao As clsGenAtuacao

' Arrays para controle
Private paRegiao() As Variant
Private paAtuacao() As Variant
Private paOperadora() As Variant

' Variaveis da classe
Private lid As Long
Private sPontaIP As String
Private sVelocidade As String
Private sPontaE1 As String
Private lSequencial As Long
Private sRoteador As String
Private sSlot As String
Private sVipSlot As String
Private sPorta As String
Private sNomeNo As String
Private sNumNo As String
Private sSlotE1 As String
Private sPortaE1 As String
'
'
Public Function GravarDados( _
    ByVal psTipoGravacao As String _
    ) As Boolean

' Inicializa o trap de erros
On Error GoTo ErrorHandler
    
    GravarDados = False
'    Set objMsg = New clsgenMensagem
'    Set clsgenTronco = New clsgenTronco
'
'    If clsgenTronco.InformaBanco(frmgenMDI_SGAPlus.objBanco, objMsg) Then
'        If clsgenTronco.GravarDados(ID, psTipoGravacao, PontaIP, Velocidade, PontaE1, Sequencial, Roteador, SlotE1, VipSlot, Porta, NomeNo, NumNo, SlotE1, PortaE1, objMsg) Then
'            GravarDados = True
'        End If
'    Else
'        If objMsg.TipoMsg <> "" Then GoTo ErrorHandler
'    End If
'    If objMsg.TipoMsg <> "" Then GoTo ErrorHandler

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
    ByVal psID As String, _
    Optional blConsulta As Boolean _
    ) As Boolean

On Error GoTo ErrorHandler
    
    Dim LocalTronco() As Variant
    
    Set objMsg = New clsgenMensagem
'    Set clsgenTronco = New clsgenTronco
'
'    Call clsgenTronco.InformaBanco(frmgenMDI_SGAPlus.objBanco, objMsg)
'
'    If objMsg.TipoMsg <> "" Then GoTo ErrorHandler
'
'    Call clsgenTronco.RecuperarDados(psID, LocalTronco, objMsg)
'
'    If objMsg.TipoMsg <> "" Then GoTo ErrorHandler
'
'    If Not IsEmpty(LocalTronco(1, 0)) Then
'        frmGenManterTroncoSecundario.psTipoGravacao = "A"
'        frmGenManterTroncoSecundario.txtPontaIP = LocalTronco(1, 0)
'        frmGenManterTroncoSecundario.psID = psID
'        frmGenManterTroncoSecundario.txtVelocidade = LocalTronco(2, 0)
'        frmGenManterTroncoSecundario.txtPontaE1 = LocalTronco(3, 0)
'        frmGenManterTroncoSecundario.TxtSequencial = LocalTronco(4, 0)
'        frmGenManterTroncoSecundario.txtRoteador = LocalTronco(5, 0)
'        frmGenManterTroncoSecundario.TxtSlotMipVip = LocalTronco(6, 0)
'        frmGenManterTroncoSecundario.TxtVipSlot = IIf(IsNull(LocalTronco(7, 0)), "", LocalTronco(7, 0))
'        frmGenManterTroncoSecundario.txtNomeNo = LocalTronco(9, 0)
'        frmGenManterTroncoSecundario.TxtNumeroNo = LocalTronco(10, 0)
'        frmGenManterTroncoSecundario.TxtSlotPlacaE1 = LocalTronco(11, 0)
'        frmGenManterTroncoSecundario.txtPorta = LocalTronco(8, 0)
'        frmGenManterTroncoSecundario.TxtPortaE1 = LocalTronco(12, 0)
'    End If
'
'    If blConsulta Then frmGenManterTroncoSecundario.CmdSalvar.Enabled = False
'
'    frmGenManterTroncoSecundario.Show vbModal

Exit Function

ErrorHandler:
    If objMsg.TipoMsg = "" Then
        objMsg.TipoMsg = "E"
        objMsg.TituloMsg = "clsgenOC_TroncoPrincipal.CarregarForm"
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

Public Property Let PontaIP(ByVal psPontaIP As String)
    sPontaIP = psPontaIP
End Property

Public Property Get PontaIP() As String
    PontaIP = sPontaIP
End Property

Public Property Let Velocidade(ByVal psVelocidade As String)
    sVelocidade = psVelocidade
End Property

Public Property Get Velocidade() As String
    Velocidade = sVelocidade
End Property

'sSQL = sSQL & " PONTA_E1,"
Public Property Let PontaE1(ByVal psPontaE1 As String)
    sPontaE1 = psPontaE1
End Property

Public Property Get PontaE1() As String
    PontaE1 = sPontaE1
End Property

'sSQL = sSQL & " SEQUENCIAL,"
Public Property Let Sequencial(ByVal plSequencial As Long)
    lSequencial = plSequencial
End Property

Public Property Get Sequencial() As Long
    Sequencial = lSequencial
End Property

'sSQL = sSQL & " ROTEADOR,"
Public Property Let Roteador(ByVal psRoteador As String)
    sRoteador = psRoteador
End Property

Public Property Get Roteador() As String
    Roteador = sRoteador
End Property

'sSQL = sSQL & " SLOT,"
Public Property Let Slot(ByVal psSlot As String)
    sSlot = psSlot
End Property

Public Property Get Slot() As String
    Slot = sSlot
End Property

'sSQL = sSQL & " VIPSLOT,"
Public Property Let VipSlot(ByVal psVipSlot As String)
    sVipSlot = psVipSlot
End Property

Public Property Get VipSlot() As String
    VipSlot = sVipSlot
End Property

'sSQL = sSQL & " PORTA,"
Public Property Let Porta(ByVal psPorta As String)
    sPorta = psPorta
End Property

Public Property Get Porta() As String
    Porta = sPorta
End Property

'sSQL = sSQL & " NOME_NO,"
Public Property Let NomeNo(ByVal psNomeNo As String)
    sNomeNo = psNomeNo
End Property

Public Property Get NomeNo() As String
    NomeNo = sNomeNo
End Property

'sSQL = sSQL & " NUM_NO,"
Public Property Let NumNo(ByVal psNumNo As String)
    sNumNo = psNumNo
End Property

Public Property Get NumNo() As String
    NumNo = sNumNo
End Property

'sSQL = sSQL & " SLOT_E1,"
Public Property Let SlotE1(ByVal psSlotE1 As String)
    sSlotE1 = psSlotE1
End Property

Public Property Get SlotE1() As String
    SlotE1 = sSlotE1
End Property

'sSQL = sSQL & " PORTA_E1"
Public Property Let PortaE1(ByVal psPortaE1 As String)
    sPortaE1 = psPortaE1
End Property

Public Property Get PortaE1() As String
    PortaE1 = sPortaE1
End Property

Private Sub Class_Terminate()
    'Set objBanco = Nothing
    'Set clsgenTronco = Nothing
    Set objMsg = Nothing
End Sub
