VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsGenMotivoRecusa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim objMsg   As clsgenMensagem
Dim objBanco As SGAPLUSBanco.clsBanco


Public Function RecuperarDados(ByVal piID_MOTIVO As Integer, _
                                               ByRef paDadosMotivo() As Variant) As Boolean

On Error GoTo ErrorHandler:

    Dim sSQL As String
    Dim bRetorno As Boolean
    Dim bExcedentes As Boolean

    
    RecuperarDados = False
    
    sSQL = ""
    sSQL = sSQL & " SELECT  DESCRICAO FROM MOTIVO_RECUSA"
    sSQL = sSQL & " WHERE ID = " & piID_MOTIVO
    
    bRetorno = objBanco.Obter_Selecao(sSQL, paDadosMotivo(), 0, bExcedentes, False, "", "")
 
    If Not bRetorno Then
        MsgBox "MOTIVO INEXISTENTE"
        GoTo Finalizar
        
    End If

    RecuperarDados = True
    
    GoTo Finalizar

Finalizar:

    Exit Function
        
ErrorHandler:

    RecuperarDados = False
    
    GoTo Finalizar

End Function


Public Function ListarDados(ByRef paMotivos() As Variant) As Boolean

On Error GoTo ErrorHandler:

    Dim sSQL As String
    Dim bRetorno As Boolean
    Dim bExcedentes As Boolean

    ListarDados = False
    
    sSQL = ""
    sSQL = sSQL & " SELECT ID, DESCRICAO FROM SGAPLUS_ADM.MOTIVO_RECUSA"
        
    bRetorno = objBanco.Obter_Selecao(sSQL, paMotivos(), 0, bExcedentes, False, "", "")
 
    If Not bRetorno Then
        MsgBox "MOTIVO INEXISTENTE"
        GoTo Finalizar
        
    End If

    ListarDados = True
    
    GoTo Finalizar

Finalizar:

    Exit Function
        
ErrorHandler:

    ListarDados = False
    
    GoTo Finalizar

End Function

Public Function GravarDados(ByVal psTipoGravacao As String, _
                                          ByVal psMotivo As String, _
                                          Optional ByVal piID_MOTIVO As Integer) As Boolean

On Error GoTo ErrorHandler:

    Dim sSQL As String
    Dim bRetorno As Boolean
    
    GravarDados = False
    
    bRetorno = objBanco.Iniciar_Transacao()
    
    Select Case psTipoGravacao
        
        Case "I"
                
            sSQL = ""
            sSQL = sSQL & " INSERT INTO MOTIVO_RECUSA "
            sSQL = sSQL & " (ID, DESCRICAO) "
            sSQL = sSQL & " VALUES (SQ_MOTIVO_RECUSA.NEXTVAL, '" & psMotivo & "')"
            bRetorno = objBanco.Executar_SQL(sSQL)
                       
        Case "A"
            
            sSQL = ""
            sSQL = sSQL & " UPDATE MOTIVO_RECUSA SET "
            sSQL = sSQL & " DESCRICAO = '" & psMotivo & "'"
            sSQL = sSQL & " WHERE ID = " & piID_MOTIVO

            bRetorno = objBanco.Executar_SQL(sSQL)
        
    End Select

    If Not bRetorno Then
                
        objBanco.Cancelar_Transacao
                
        GoTo Finalizar
            
    End If
            
    bRetorno = objBanco.Finalizar_Transacao
    
    GravarDados = True
    
    GoTo Finalizar
    
Finalizar:

    Exit Function
        
ErrorHandler:

    objBanco.Cancelar_Transacao
    
    GoTo Finalizar

End Function


Public Function ExcluirDados(ByVal piID_MOTIVO As Integer) As Boolean

On Error GoTo ErrorHandler:

    Dim sSQL As String
    Dim bRetorno As Boolean

    ExcluirDados = False
    
    bRetorno = objBanco.Iniciar_Transacao()
    
        sSQL = ""
        sSQL = sSQL & " DELETE FROM MOTIVO_RECUSA"
        sSQL = sSQL & " WHERE ID = " & piID_MOTIVO
        
        bRetorno = objBanco.Executar_SQL(sSQL)
        
        If Not bRetorno Then
            
            bRetorno = objBanco.Cancelar_Transacao()
            
            GoTo Finalizar
        
        End If
    
    bRetorno = objBanco.Finalizar_Transacao()
    
    ExcluirDados = True
    
    GoTo Finalizar
    
Finalizar:

    Exit Function
        
ErrorHandler:

    objBanco.Cancelar_Transacao
    
    ExcluirDados = False
    
    GoTo Finalizar

End Function

Public Function InformaBanco(ByVal objNomeBanco As SGAPLUSBanco.clsBanco) As Boolean

On Error GoTo ErrorHandler:

    InformaBanco = False
    Set objBanco = objNomeBanco
    InformaBanco = True

Finalizar:

    Exit Function
        
ErrorHandler:

    objMsg.TipoMsg = "E"
    objMsg.TituloMsg = "clsgenTronco.RecuperarDados"
    objMsg.Descricao = Err.Description
    InformaBanco = False
    
    GoTo Finalizar

End Function
Public Function InformaMensagem(ByVal objNomeMensagem As clsgenMensagem)

On Error GoTo ErrorHandler:

    InformaMensagem = False
    
    Set objMensagem = objNomeMensagem
    
    InformaMensagem = True
    
    GoTo Finalizar
    
Finalizar:

    Exit Function
    
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Houve um erro de execução no Método:" & _
                            "InformaMensagem da classe clsgenMotivoRecusa"
    
    GoTo Finalizar

End Function

Private Sub Class_Terminate()
    Set objBanco = Nothing
End Sub
