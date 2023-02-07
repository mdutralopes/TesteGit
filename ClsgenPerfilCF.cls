VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsgenPerfilCF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim objBanco As SGAPLUSBanco.clsBanco

Public Function RecuperarDados(ByVal plID_TipoFase As Long, _
                                ByVal plID_CF As Long, _
                                ByVal piID_Produto As Integer, _
                                               ByRef paDadosPerfilCF() As Variant, _
                                               ByRef objMsg As ClsgenMsg) As Boolean
On Error GoTo ErrorHandler:
    Dim sSQL As String
    Dim bRetorno As Boolean
    Dim bExcedentes As Boolean
    
    RecuperarDados = False

    sSQL = ""
    sSQL = sSQL & "     SELECT  TF.DESCRICAO,"
    sSQL = sSQL & "         PCF.ATUALIZA, "
    sSQL = sSQL & "         PCF.EXCLUI, "
    sSQL = sSQL & "         PCF.CONSULTA"
    sSQL = sSQL & "     FROM    PERFIL_CENTRO_FUNCIONAL PCF,"
    sSQL = sSQL & "             TIPO_FASE TF,"
    sSQL = sSQL & "             DURACAO_FASE DF"
    sSQL = sSQL & "     Where   "
    sSQL = sSQL & "             PCF.TIPO_FASE_ID =" & plID_TipoFase
    sSQL = sSQL & "     AND     PCF.CENTRO_FUNCIONAL_ID =" & plID_CF
    sSQL = sSQL & "     AND     DF.PRODUTO_ID= " & piID_Produto
    'TIPO_FASE
    sSQL = sSQL & "     AND     PCF.TIPO_FASE_ID = TF.ID"
    'DURACAO_FASE
    sSQL = sSQL & "     AND     PCF.TIPO_FASE_ID = DF.ID"
    
    sSQL = sSQL & "     ORDER BY DF.PRODUTO_ID, "
    sSQL = sSQL & "         DF.ACAO_ID, "
    sSQL = sSQL & "         DF.SEQUENCIA_ACAO"
    
    bRetorno = objBanco.Obter_Selecao(sSQL, paDadosPerfilCF(), 0, bExcedentes, False, "", "")
 
    If Not bRetorno Then GoTo ErrorHandler
    
    RecuperarDados = True
    
    GoTo Finalizar
Finalizar:
    Exit Function
ErrorHandler:
    objMsg.TipoMsg = "E"
    objMsg.TituloMsg = "ClsgenPerfilCF.RecuperarDados"
    objMsg.Descricao = Err.Description & objBanco.objErro(0).Description & objBanco.objErro(0).Number
    
    RecuperarDados = False
    GoTo Finalizar
End Function
Public Function ListarDados(ByRef paPerfilCF() As Variant, _
                                        ByRef objMsg As ClsgenMsg, _
                                        Optional ByVal piID_Produto As Integer) As Boolean
On Error GoTo ErrorHandler:
    Dim sSQL As String
    Dim bRetorno As Boolean
    Dim bExcedentes As Boolean
    
    ListarDados = False
    
    sSQL = ""
    sSQL = sSQL & "     SELECT  TF.DESCRICAO,"
    sSQL = sSQL & "         PCF.ATUALIZA, "
    sSQL = sSQL & "         PCF.EXCLUI, "
    sSQL = sSQL & "         PCF.CONSULTA"
    sSQL = sSQL & "     FROM    PERFIL_CENTRO_FUNCIONAL PCF,"
    sSQL = sSQL & "             TIPO_FASE TF,"
    sSQL = sSQL & "             DURACAO_FASE DF"
    sSQL = sSQL & "     Where   "
    'TIPO_FASE
    sSQL = sSQL & "             PCF.TIPO_FASE_ID = TF.ID"
    'DURACAO_FASE
    sSQL = sSQL & "     AND     PCF.TIPO_FASE_ID = DF.ID"
    'Parâmetro Produto opcional
    If piID_Produto <> 0 Then
        sSQL = sSQL & "     AND     DF.PRODUTO_ID= " & piID_Produto
    End If
    sSQL = sSQL & "     ORDER BY DF.PRODUTO_ID, "
    sSQL = sSQL & "         DF.ACAO_ID, "
    sSQL = sSQL & "         DF.SEQUENCIA_ACAO"
    
    bRetorno = objBanco.Obter_Selecao(sSQL, paPerfilCF(), 0, bExcedentes, True, "", "")
 
    If Not bRetorno Then GoTo ErrorHandler
    
    ListarDados = True
    
    GoTo Finalizar

Finalizar:
    Exit Function
        
ErrorHandler:
    ListarDados = False
    
    objMsg.TipoMsg = "E"
    objMsg.TituloMsg = "ClsgenPerfilCF.ListarDados"
    objMsg.Descricao = Err.Description & objBanco.objErro(0).Description & objBanco.objErro(0).Number
    
    GoTo Finalizar
End Function
Public Function InformaBanco(ByVal objNomeBanco As SGAPLUSBanco.clsBanco, _
                                    ByRef objMsg As ClsgenMsg) As Boolean
On Error GoTo ErrorHandler:
    InformaBanco = False
    Set objBanco = objNomeBanco
    InformaBanco = True

Finalizar:
    Exit Function
        
ErrorHandler:
    objMsg.TipoMsg = "E"
    objMsg.TituloMsg = "clsgenTipoFase.InformaBanco"
    objMsg.Descricao = Err.Description
    
    InformaBanco = False
    
    GoTo Finalizar
End Function
Public Function GravarDados(ByVal psTipoGravacao As String, _
                                          ByRef objMsg As ClsgenMsg, _
                                          ByVal plID_TipoFase As Long, _
                                          ByVal plID_CF As Long, _
                                          ByVal piConsulta As Integer, _
                                          ByVal piAtualiza As Integer, _
                                          ByVal piExclui As Integer) As Boolean
On Error GoTo ErrorHandler:
    Dim sSQL As String
    Dim bRetorno As Boolean
    
    GravarDados = False
    
    bRetorno = objBanco.Iniciar_Transacao()
    
    If Not bRetorno Then GoTo ErrorHandler
    
    Select Case psTipoGravacao
        Case "I"
            sSQL = ""
            sSQL = sSQL & " INSERT INTO PERFIL_CENTRO_FUNCIONAL"
            sSQL = sSQL & " (CENTRO_FUNCIONAL_ID, TIPO_FASE_ID, ATUALIZA, EXCLUI, CONSULTA) "
            sSQL = sSQL & " VALUES (plID_CF, "
            sSQL = sSQL & plID_TipoFase & ","
            sSQL = sSQL & piAtualiza & ","
            sSQL = sSQL & piExclui & ","
            sSQL = sSQL & piConsulta & ")"
            
            bRetorno = objBanco.Executar_SQL(sSQL)
        Case "A"
            sSQL = ""
            sSQL = sSQL & " UPDATE PERFIL_CENTRO_FUNCIONAL SET "
            sSQL = sSQL & " ATUALIZA =" & piAtualiza & ","
            sSQL = sSQL & " EXCLUI =" & piExclui & ","
            sSQL = sSQL & " CONSULTA =" & piConsulta & ","
            sSQL = sSQL & " WHERE CENTRO_FUNCIONAL_ID = " & plID_CF
            sSQL = sSQL & " AND TIPO_FASE_ID = " & plID_TipoFase

            bRetorno = objBanco.Executar_SQL(sSQL)
    End Select
    
    If Not bRetorno Then GoTo ErrorHandler
            
    bRetorno = objBanco.Finalizar_Transacao
    
    If Not bRetorno Then GoTo ErrorHandler
    
    GravarDados = True
    
    GoTo Finalizar
Finalizar:
    Exit Function
    
ErrorHandler:
    objBanco.Cancelar_Transacao
    
    objMsg.TipoMsg = "E"
    objMsg.TituloMsg = "ClsgenPerfilCF.GravarDados"
    objMsg.Descricao = Err.Description & objBanco.objErro(0).Description & objBanco.objErro(0).Number

    GravarDados = False

    GoTo Finalizar
End Function
Public Function ExcluirDados(ByVal plID_TipoFase As Long, ByVal piID_CF As Long, ByRef objMsg As ClsgenMsg) As Boolean
On Error GoTo ErrorHandler:
    Dim sSQL As String
    Dim bRetorno As Boolean

    ExcluirDados = False

    bRetorno = objBanco.Iniciar_Transacao()
    If Not bRetorno Then GoTo ErrorHandler

    sSQL = ""
    sSQL = sSQL & " DELETE  FROM PERFIL_CENTRO_FUNCIONAL"
    sSQL = sSQL & " WHERE   CENTRO_FUNCIONAL_ID = " & plID_CF
    sSQL = sSQL & "         AND TIPO_FASE_ID = " & plID_TipoFase

    bRetorno = objBanco.Executar_SQL(sSQL)

    If Not bRetorno Then GoTo ErrorHandler

    bRetorno = objBanco.Finalizar_Transacao()
    If Not bRetorno Then GoTo ErrorHandler
    
    ExcluirDados = True

    GoTo Finalizar
Finalizar:
    Exit Function

ErrorHandler:
    objBanco.Cancelar_Transacao

    objMsg.TipoMsg = "E"
    objMsg.TituloMsg = "ClsgenPerfilCF.ExcluirDados"
    objMsg.Descricao = Err.Description & objBanco.objErro(0).Description & objBanco.objErro(0).Number
    ExcluirDados = False

    GoTo Finalizar
End Function
Private Sub Class_Terminate()
    Set objBanco = Nothing
End Sub


