VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsgenListaCentroFuncionais"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Declara as variáveis de atribuição das propriedades
Private lID_CENTRO_FUNCIONAL As Long
Private sSigla As String
Private sDescricao As String

Public Property Get ID_CENTRO_FUNCIONAL() As Long

    ID_CENTRO_FUNCIONAL = lID_CENTRO_FUNCIONAL
    
End Property

Public Property Let ID_CENTRO_FUNCIONAL(ByVal lNewValue As Long)

    lID_CENTRO_FUNCIONAL = lNewValue
    
End Property

Public Property Get Sigla() As String

    Sigla = sSigla

End Property

Public Property Let Sigla(ByVal sNewValue As String)

    sSigla = sNewValue

End Property

Public Property Get Descricao() As String

    Descricao = sDescricao

End Property

Public Property Let Descricao(ByVal sNewValue As String)

    sDescricao = sNewValue

End Property

