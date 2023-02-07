VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGenDadosControleCorrecao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private objBanco                As SGAPLUSBanco.clsBanco
Private objMensagem             As clsgenMensagem

Private lID                     As Long
Private lFaseID                 As Long
Private sTipoCorrecao           As String
Private sDataSolicitacao        As String
Private sDataExecucao           As String
Private sDataInicioPrevAnt      As String
Private sDataFimPrevAnt         As String
Private lUsuarioIDSolic         As Long
Private lUsuarioIDExec          As Long
Private sJustificativa          As String

Public Property Get ID() As Long

    ID = lID

End Property

Public Property Let ID(ByVal lNewValue As Long)

    lID = lNewValue

End Property

Public Property Get FaseID() As Long

    FaseID = lFaseID

End Property

Public Property Let FaseID(ByVal lNewValue As Long)

    lFaseID = lNewValue

End Property

Public Property Get UsuarioIDSolic() As Long

    UsuarioIDSolic = lUsuarioIDSolic

End Property

Public Property Let UsuarioIDSolic(ByVal lNewValue As Long)

    lUsuarioIDSolic = lNewValue

End Property

Public Property Get UsuarioIDExec() As Long

    UsuarioIDExec = lUsuarioIDExec

End Property

Public Property Let UsuarioIDExec(ByVal lNewValue As Long)

    lUsuarioIDExec = lNewValue

End Property

Public Property Get TipoCorrecao() As String

    TipoCorrecao = sTipoCorrecao

End Property

Public Property Let TipoCorrecao(ByVal sNewValue As String)

    sTipoCorrecao = sNewValue

End Property

Public Property Get DataExecucao() As String

    DataExecucao = sDataExecucao

End Property

Public Property Let DataExecucao(ByVal sNewValue As String)

    sDataExecucao = sNewValue

End Property

Public Property Get DataSolicitacao() As String

    DataSolicitacao = sDataSolicitacao

End Property

Public Property Let DataSolicitacao(ByVal sNewValue As String)

    sDataSolicitacao = sNewValue

End Property

Public Property Get DataInicioPrevAnt() As String

    DataInicioPrevAnt = sDataInicioPrevAnt

End Property

Public Property Let DataInicioPrevAnt(ByVal sNewValue As String)

    sDataInicioPrevAnt = sNewValue

End Property

Public Property Get DataFimPrevAnt() As String

    DataFimPrevAnt = sDataFimPrevAnt

End Property

Public Property Let DataFimPrevAnt(ByVal sNewValue As String)

    sDataFimPrevAnt = sNewValue

End Property

Public Property Get Justificativa() As String

    Justificativa = sJustificativa

End Property

Public Property Let Justificativa(ByVal sNewValue As String)

    sJustificativa = sNewValue

End Property

Public Function InformaBanco(ByVal objNomeBanco As clsBanco) As Boolean

On Error GoTo ErrorHandler:

    InformaBanco = False
    
    Set objBanco = objNomeBanco
    
    InformaBanco = True

Finalizar:

    Exit Function
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.TituloMsg = "clsGenAnotacao.InformaBanco"
    objMensagem.Descricao = Err.Description & objBanco.objErro(0).Description & objBanco.objErro(0).Number
    
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
                            "InformaMensagem da classe ClsGenAnotacao"
    
    GoTo Finalizar

End Function

Public Function RecuperarDados(plFaseID As Long, psTipoCorrecao As String) As Boolean

On Error GoTo ErrorHandler:

    Dim sSQL            As String ' Guardar a concatenação do comando SQL
    Dim bRetorno        As Boolean ' Controla o retorno dos métodos do banco
    Dim RsDados         As ADODB.Recordset
    Dim lIDMax          As Long
    Dim FlagRecupera    As Boolean
    
    ' Preparar o retorno do método como verdadeiro
    RecuperarDados = False
    
    FlagRecupera = True
    
    lIDMax = 0
    
    If psTipoCorrecao = "D" Then
    
        sSQL = "SELECT "
        sSQL = sSQL & "MAX(ID) AS ID "
        sSQL = sSQL & "FROM CONTROLE_CORRECAO "
        sSQL = sSQL & " WHERE FASE_ID = " & plFaseID
        
        Set RsDados = objBanco.Obter_Recordset(sSQL, "", "", 30)
                
        If Not RsDados.EOF Then
    
            lIDMax = IIf(Not IsNull(RsDados!ID), RsDados!ID, 0)
        
        End If
        
        RsDados.Close
        Set RsDados = Nothing
    
    End If
    
    sSQL = "SELECT "
    sSQL = sSQL & "ID, "
    sSQL = sSQL & "FASE_ID, "
    sSQL = sSQL & "TIPO_CORRECAO, "
    sSQL = sSQL & "DATA_SOLICITACAO, "
    sSQL = sSQL & "DATA_EXECUCAO, "
    sSQL = sSQL & "DATA_INICIO_PREV_ANTERIOR, "
    sSQL = sSQL & "DATA_FIM_PREV_ANTERIOR, "
    sSQL = sSQL & "JUSTIFICATIVA, "
    sSQL = sSQL & "USUARIO_ID_SOLICITACAO, "
    sSQL = sSQL & "USUARIO_ID_EXECUCAO "
    sSQL = sSQL & "FROM CONTROLE_CORRECAO "
    sSQL = sSQL & " WHERE FASE_ID         = " & plFaseID
    sSQL = sSQL & "   AND TIPO_CORRECAO   = '" & psTipoCorrecao & "'"
    
    If psTipoCorrecao <> "D" Then
       
       sSQL = sSQL & "   AND ULTIMO_REGISTRO = 'S'"
    
    End If
    
    sSQL = sSQL & " ORDER BY ID DESC"
    
    Set RsDados = objBanco.Obter_Recordset(sSQL, "", "", 30)
            
    If Not RsDados.EOF Then
       'Caso esteja recuperando um registro de Devolucao, verificar se o último registro
       'inserido na tabela é de Devolucao, caso não seja, o sistema não precisa exibi-lo.
       
       If lIDMax <> 0 And (RsDados!ID < lIDMax) Then
       
          FlagRecupera = False
          
       Else
       
          FlagRecupera = True
       
       End If
        
       If FlagRecupera Then
          
           ID = "0" & RsDados!ID
            
           FaseID = "0" & RsDados!Fase_ID
            
           UsuarioIDSolic = "0" & RsDados!USUARIO_ID_SOLICITACAO
            
           UsuarioIDExec = "0" & RsDados!USUARIO_ID_EXECUCAO
            
           TipoCorrecao = "" & RsDados!TIPO_CORRECAO
            
           DataSolicitacao = IIf(Not IsNull(RsDados!DATA_SOLICITACAO), _
                                            Format(RsDados!DATA_SOLICITACAO, "DD/MM/YYYY"), "")
           
           DataExecucao = IIf(Not IsNull(RsDados!DATA_EXECUCAO), _
                                            Format(RsDados!DATA_EXECUCAO, "DD/MM/YYYY"), "")
            
           DataInicioPrevAnt = IIf(Not IsNull(RsDados!DATA_INICIO_PREV_ANTERIOR), _
                                   Format(RsDados!DATA_INICIO_PREV_ANTERIOR, "DD/MM/YYYY"), "")
            
           DataFimPrevAnt = IIf(Not IsNull(RsDados!DATA_FIM_PREV_ANTERIOR), _
                                   Format(RsDados!DATA_FIM_PREV_ANTERIOR, "DD/MM/YYYY"), "")
            
           Justificativa = "" & RsDados!Justificativa
            
           RecuperarDados = True
        
       End If
    
    End If
    
    RsDados.Close
    Set RsDados = Nothing
    
Finalizar:

    On Error GoTo 0
    
    Exit Function
    
ErrorHandler:

    If objBanco.objErro Is Nothing Then
       
       objMensagem.TipoMsg = "E"
       objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                           vbCrLf & "Houve um erro de execução no método: " & vbCrLf & vbCrLf & "RecuperarDados da classe " & _
                          "clsgenDadosControleCorrecao"
    Else
       objMensagem.TipoMsg = "E"
       objMensagem.Descricao = "Erro: " & objBanco.objErro(0).Number & " - " & objBanco.objErro(0).Description & vbCrLf & vbCrLf & _
                           vbCrLf & "Houve um erro de execução no método: " & vbCrLf & vbCrLf & "RecuperarDados da classe " & _
                          "clsgenDadosControleCorrecao"
    End If

    GoTo Finalizar
   
End Function

Public Function IncluirDados(psTipoAnotacao As String) As Boolean

    Dim bRetorno As Boolean ' Controla o retorno dos métodos do banco
    Dim sSQL As String
    
    IncluirDados = False
    
    On Error GoTo ErrorHandler
          
    sSQL = "INSERT INTO CONTROLE_CORRECAO ("
    sSQL = sSQL & " ID, "
    sSQL = sSQL & " FASE_ID, "
    sSQL = sSQL & " TIPO_CORRECAO, "
    sSQL = sSQL & " DATA_SOLICITACAO, "
    sSQL = sSQL & " DATA_EXECUCAO, "
    sSQL = sSQL & " DATA_INICIO_PREV_ANTERIOR, "
    sSQL = sSQL & " DATA_FIM_PREV_ANTERIOR, "
    sSQL = sSQL & " JUSTIFICATIVA, "
    sSQL = sSQL & " USUARIO_ID_SOLICITACAO, "
    sSQL = sSQL & " USUARIO_ID_EXECUCAO, "
    sSQL = sSQL & " ULTIMO_REGISTRO "
    sSQL = sSQL & " ) VALUES ("
    sSQL = sSQL & "SQ_CONTROLE_CORRECAO.NEXTVAL, "
    sSQL = sSQL & FaseID & ","
    
    sSQL = sSQL & "'" & TipoCorrecao & "', "
    
    If DataSolicitacao <> "" Then
    
       sSQL = sSQL & "TO_DATE('" & Format(DataSolicitacao, "DD/MM/YYYY") & "', 'DD/MM/YYYY'), "
    
    Else
       
       sSQL = sSQL & "NULL, "
    
    End If
    
    If DataExecucao <> "" Then
    
       sSQL = sSQL & "TO_DATE('" & Format(DataExecucao, "DD/MM/YYYY") & "', 'DD/MM/YYYY'), "
    
    Else
       
       sSQL = sSQL & "NULL, "
    
    End If
    
    If DataInicioPrevAnt <> "" Then
    
       sSQL = sSQL & "TO_DATE('" & Format(DataInicioPrevAnt, "DD/MM/YYYY") & "', 'DD/MM/YYYY'), "
    
    Else
       
       sSQL = sSQL & "NULL, "
    
    End If
    
    If DataFimPrevAnt <> "" Then
    
       sSQL = sSQL & "TO_DATE('" & Format(DataFimPrevAnt, "DD/MM/YYYY") & "', 'DD/MM/YYYY'), "
    
    Else
       
       sSQL = sSQL & "NULL, "
    
    End If
    
    sSQL = sSQL & "'" & Justificativa & "', "
    
    If UsuarioIDSolic <> 0 Then
    
       sSQL = sSQL & UsuarioIDSolic & ", "
    
    Else
    
       sSQL = sSQL & "0, "
    
    End If
    
    If UsuarioIDExec <> 0 Then
    
       sSQL = sSQL & UsuarioIDExec & ", "
    
    Else
    
       sSQL = sSQL & "0, "
    
    End If
        
    If TipoCorrecao = "D" Then
       
       sSQL = sSQL & "'N'"
       
    Else
       
       sSQL = sSQL & "'S'"
    
    End If
    sSQL = sSQL & ")"
    
    bRetorno = objBanco.Executar_SQL(sSQL)
    
    If Not bRetorno Then
    
        GoTo ErrorHandler
    
    End If
        
    IncluirDados = True
    
    GoTo Finalizar
    
Finalizar:
    
   On Error GoTo 0
   
   Exit Function
    
ErrorHandler:

    If objBanco.objErro Is Nothing Then
       
       objMensagem.TipoMsg = "E"
       objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                           vbCrLf & "Houve um erro de execução no método: " & vbCrLf & vbCrLf & "IncluirDados da classe " & _
                          "clsGenDadosControleCorrecao"
    Else
       objMensagem.TipoMsg = "E"
       objMensagem.Descricao = "Erro: " & objBanco.objErro(0).Number & " - " & objBanco.objErro(0).Description & vbCrLf & vbCrLf & _
                   "Houve um erro de execução no Método: " & vbCrLf & vbCrLf & "IncluirDados " & _
                   "da classe clsGenDadosControleCorrecao"
    
    End If
    
    GoTo Finalizar

End Function


Public Function FaseExistente(plFase_ID As Long, psTipoCorrecao As String) As Boolean

    Dim sSQL As String
    Dim RsDados As ADODB.Recordset
    
    On Error GoTo ErrorHandler

    FaseExistente = False
    
    sSQL = "SELECT ID "
    sSQL = sSQL & " FROM CONTROLE_CORRECAO "
    sSQL = sSQL & "   WHERE FASE_ID         = " & plFase_ID
    sSQL = sSQL & "     AND TIPO_CORRECAO   = '" & psTipoCorrecao & "'"
    sSQL = sSQL & "     AND ULTIMO_REGISTRO = 'S'"
    
    Set RsDados = objBanco.Obter_Recordset(sSQL, "", "", 30)
            
    If RsDados.EOF Then
    
       GoTo Finalizar
    
    End If
    
    FaseExistente = True
    
    RsDados.Close
    Set RsDados = Nothing
    
    GoTo Finalizar
    
Finalizar:
    
    On Error GoTo 0
    
    Exit Function
    
ErrorHandler:

    If objBanco.objErro Is Nothing Then
       
       objMensagem.TipoMsg = "E"
       objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                           vbCrLf & "Houve um erro de execução no método: " & vbCrLf & vbCrLf & "FaseExistente da classe " & _
                          "clsGenDadosControleCorrecao"
    Else
       objMensagem.TipoMsg = "E"
       objMensagem.Descricao = "Erro: " & objBanco.objErro(0).Number & " - " & objBanco.objErro(0).Description & vbCrLf & vbCrLf & _
                           vbCrLf & "Houve um erro de execução no método: " & vbCrLf & vbCrLf & "FaseExistente da classe " & _
                          "clsGenDadosControleCorrecao"
    End If
    
    GoTo Finalizar

End Function

Public Function AtualizarDados(plFaseID As Long, psTipoCorrecao As String) As Boolean

    Dim bRetorno As Boolean 'Controla o retorno dos métodos do banco
    Dim sSQL     As String
    
    AtualizarDados = False
    
    On Error GoTo ErrorHandler
        
    sSQL = "UPDATE CONTROLE_CORRECAO SET "
    sSQL = sSQL & " JUSTIFICATIVA = '" & Justificativa & "'"
    sSQL = sSQL & " WHERE FASE_ID         = " & plFaseID
    sSQL = sSQL & "   AND TIPO_CORRECAO   = '" & psTipoCorrecao & "'"
    sSQL = sSQL & "   AND ULTIMO_REGISTRO = 'S'"
    
    bRetorno = objBanco.Executar_SQL(sSQL)
    
    If Not bRetorno Then
    
        GoTo ErrorHandler
    
    End If
        
    AtualizarDados = True
    
    GoTo Finalizar
    
Finalizar:
    
   On Error GoTo 0
   
   Exit Function
    
ErrorHandler:

    If objBanco.objErro Is Nothing Then
       
       objMensagem.TipoMsg = "E"
       objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                           vbCrLf & "Houve um erro de execução no método: " & vbCrLf & vbCrLf & "AtualizarDados da classe " & _
                          "clsGenDadosControleCorrecao"
    Else
       
       objMensagem.TipoMsg = "E"
       objMensagem.Descricao = "Erro: " & objBanco.objErro(0).Number & " - " & objBanco.objErro(0).Description & vbCrLf & vbCrLf & _
                   "Houve um erro de execução no Método: " & vbCrLf & vbCrLf & "AtualizarDados " & _
                   "da classe clsGenDadosControleCorrecao"
    
    End If
    
    GoTo Finalizar

End Function

Public Function AtualizarExecucao(plFaseID As Long, psTipoCorrecao As String) As Boolean

    Dim bRetorno As Boolean 'Controla o retorno dos métodos do banco
    Dim sSQL     As String
    
    AtualizarExecucao = False
    
    On Error GoTo ErrorHandler
        
    sSQL = "UPDATE CONTROLE_CORRECAO SET "
    sSQL = sSQL & " DATA_EXECUCAO        = TO_DATE('" & Format(Date, "DD/MM/YYYY") & "', 'DD/MM/YYYY'), "
    sSQL = sSQL & " USUARIO_ID_EXECUCAO  = " & UsuarioIDExec & ", "
    sSQL = sSQL & " ULTIMO_REGISTRO = 'N'"
    sSQL = sSQL & " WHERE FASE_ID         = " & plFaseID
    sSQL = sSQL & "   AND TIPO_CORRECAO   = '" & psTipoCorrecao & "'"
    sSQL = sSQL & "   AND ULTIMO_REGISTRO = 'S'"
    
    bRetorno = objBanco.Executar_SQL(sSQL)
    
    If Not bRetorno Then
    
        GoTo ErrorHandler
    
    End If
        
    AtualizarExecucao = True
    
    GoTo Finalizar
    
Finalizar:
    
   On Error GoTo 0
   
   Exit Function
    
ErrorHandler:

    If objBanco.objErro Is Nothing Then
       
       objMensagem.TipoMsg = "E"
       objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                           vbCrLf & "Houve um erro de execução no método: " & vbCrLf & vbCrLf & "AtualizarDados da classe " & _
                          "clsGenDadosControleCorrecao"
    Else
       
       objMensagem.TipoMsg = "E"
       objMensagem.Descricao = "Erro: " & objBanco.objErro(0).Number & " - " & objBanco.objErro(0).Description & vbCrLf & vbCrLf & _
                   "Houve um erro de execução no Método: " & vbCrLf & vbCrLf & "AtualizarDados " & _
                   "da classe clsGenDadosControleCorrecao"
    
    End If
    
    GoTo Finalizar

End Function


