VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsgenOC_InterfacePRF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private sTipoRegistro As String * 1
Private sSiglaInstalacao As String * 3
Private sCodigoArea As String * 4
Private sCNL_PontaA As String * 6
Private sCNL_PontaB As String * 6
Private sDesignacaoFacilidade As String * 30
Private sDataAtivacao As String * 8
Private sDataCancelamento As String * 8
Private sSGA_NacionalInternacional As String * 1
Private sInstalacao As String * 1
Private sSiglaServico As String * 3
Private sIndicador As String * 1
Private sContaCorrente As String * 15
Private sVelocidadeCircuito As String * 6
Private sNumeroPedido As String * 16
Private sDesignacaoE1 As String * 40
Private sCentroFuncional As String * 10
Private sCodigoAcessorio As String * 6
Private sQuantidadeAcessorio As String * 3
Private sTipoAtualizacao As String * 2
Private sAlteracaoEndereco As String * 8
Private sAlteracaoQuantidadeAcessorio As String * 8
Private sMatricula As String * 8

Private sDadosEnviado As String
Private lTotalEnviados As Long

Private objInterface As clsgenDadosInterfacePRF
Private objColecao As Collection
Private objServico As clsGenServico
Private objCircuito  As clsgenCircuito
Private objCentroFuncional As clsgenCentroFuncional
Private objFase As clsgenFase
Private objOC_Acesso As clsGenOC_TratarAcessoEmbratel
Private objMensagem As clsgenMensagem

Public Function GerarArquivo(ByVal objForm As frmgenGerarArqPRF) As Boolean
    
    Dim bRetorno        As Boolean
    Dim nArqWrite       As Byte
    Dim lTotalRegistros As Long
    Dim lTotalLidos     As Long
    Dim bRegistroOK     As Boolean

On Error GoTo ErrorHandler

    GerarArquivo = False

    nArqWrite = FreeFile

    Open App.Path & "\SGA_PRF.TXT" For Output As nArqWrite

    ' Instância as classes
    Set objMensagem = New clsgenMensagem
    Set objInterface = New clsgenDadosInterfacePRF
    Set objColecao = New Collection
    Set objCircuito = New clsgenCircuito
    Set objCentroFuncional = New clsgenCentroFuncional
    Set objFase = New clsgenFase
    Set objOC_Acesso = New clsGenOC_TratarAcessoEmbratel

    ' Informa as classes de banco e mensagem para a classe de interface
    objInterface.InformaBanco frmgenMDI_SGAPlus.ObjBanco
    objInterface.InformaMensagem objMensagem
    objCircuito.InformaBanco frmgenMDI_SGAPlus.ObjBanco
    objCircuito.InformaMensagem objMensagem
    objCentroFuncional.InformaBanco frmgenMDI_SGAPlus.ObjBanco
    objCentroFuncional.InformaMensagem objMensagem
    objFase.InformaBanco frmgenMDI_SGAPlus.ObjBanco
    objFase.InformaMensagem objMensagem

    ' Busca os dados
    Set objColecao = objInterface.ObtemServicos()

    If objMensagem.TipoMsg <> "" Then

       GoTo Finalizar

    End If

    lTotalRegistros = objColecao.Count
    lTotalLidos = 0
    lTotalEnviados = 0

    objForm.PrgBar.Min = 0
    objForm.PrgBar.Max = lTotalRegistros
    objForm.LblTotalRegistros.Caption = Format(lTotalRegistros, "0000")

    For Each objServico In objColecao

        lTotalLidos = lTotalLidos + 1
        
        ' Iniciar valores na variáveis para a montagem do arquivo
        bRetorno = IniciarVariaveis()

        If objMensagem.TipoMsg <> "" Then

           GoTo Finalizar

        End If

        ' Atribuir valores na variáveis para a montagem do arquivo
        bRegistroOK = False
        bRetorno = PreencherVariaveis(bRegistroOK)

        If objMensagem.TipoMsg <> "" Then

           GoTo Finalizar

        End If

        If bRegistroOK Then
        
            Print #nArqWrite, sDadosEnviado
            lTotalEnviados = lTotalEnviados + 1
            
        End If
        
        objForm.PrgBar.Value = lTotalLidos
        objForm.LblTotalRegistros.Caption = Format(objForm.PrgBar.Max, "00000")
        objForm.LblRegistrosEnviados.Caption = Format(lTotalEnviados, "00000")
        objForm.LblTotalRegistros.Refresh

    Next

    GerarArquivo = True

Finalizar:

    Close #nArqWrite

    If objMensagem.TipoMsg <> "" Then

        objMensagem.ExibeMensagem

    End If

    Set objMensagem = Nothing
    Set objColecao = Nothing
    Set objInterface = Nothing
    Set objServico = Nothing
    Set objCircuito = Nothing
    Set objCentroFuncional = Nothing
    Set objFase = Nothing
    Set objOC_Acesso = Nothing

    Exit Function

ErrorHandler:

    If Err.Number <> 0 Then
       objMensagem.TipoMsg = "E"
       objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                   "Houve um erro de execução no Método: " & vbCrLf & vbCrLf & "GerarArquivo " & _
                   "da classe clsgenOC_InterfacePRF"
    End If

    GoTo Finalizar

End Function

Private Sub Class_Terminate()

    Set objMensagem = Nothing

End Sub

Private Function IniciarVariaveis() As Boolean

On Error GoTo ErrorHandler:

    IniciarVariaveis = False
    
    sTipoRegistro = ""
    sSiglaInstalacao = ""
    sCodigoArea = ""
    sCNL_PontaA = ""
    sCNL_PontaB = ""
    sDesignacaoFacilidade = ""
    sDataAtivacao = "00000000"
    sDataCancelamento = "00000000"
    sSGA_NacionalInternacional = ""
    sInstalacao = ""
    sSiglaServico = ""
    sIndicador = ""
    sContaCorrente = ""
    sVelocidadeCircuito = ""
    sNumeroPedido = ""
    sDesignacaoE1 = ""
    sCentroFuncional = ""
    sCodigoAcessorio = ""
    sQuantidadeAcessorio = "000"
    sTipoAtualizacao = "00000000"
    sAlteracaoEndereco = "00000000"
    sAlteracaoQuantidadeAcessorio = ""
    sMatricula = ""
    
    IniciarVariaveis = True

Finalizar:
    
    On Error GoTo 0
    
    Exit Function

ErrorHandler:

    If Err.Number <> 0 Then
       objMensagem.TipoMsg = "E"
       objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                   "Houve um erro de execução no Método: " & vbCrLf & vbCrLf & "IniciarVariaveis " & _
                   "da classe clsgenOC_InterfacePRF"
    End If
       
    GoTo Finalizar

End Function

Private Function PreencherVariaveis(ByRef pbRegistroOK As Boolean) As Boolean

On Error GoTo ErrorHandler:

    Dim bRetorno As Boolean
    Dim lID_FASE_ACESSO As Long

    PreencherVariaveis = False

    ' Recuperar o ID da fase do acesso
    lID_FASE_ACESSO = objFase.ObtemIDFase(objServico.ID_SERVICO, _
                                          GbCodigoFaseAcesso)

    If objMensagem.TipoMsg <> "" Then

       GoTo Finalizar

    End If

    ' Recuperar os dados de Acesso
    bRetorno = objOC_Acesso.RecuperaDadosAcesso("P", _
                                                objServico.ID_SERVICO, _
                                                lID_FASE_ACESSO)

    If objMensagem.TipoMsg <> "" Then

       GoTo Finalizar

    End If
    
    If objOC_Acesso.objDadosAcesso.PropridedadeAcesso = "T" Then

        If objOC_Acesso.objDadosAcesso.objPropriedadeTerceiros.NomeProvedorUrbano > "" Then

            sSiglaInstalacao = "FUN"

        ElseIf objOC_Acesso.objDadosAcesso.objPropriedadeTerceiros.NomeProvedorInterurbano > "" Then

            sSiglaInstalacao = "FIN"

        End If

        sTipoRegistro = "5"
        sSiglaServico = objServico.Produto
    
        ' Recuperar os dados do circuito
        bRetorno = objCircuito.RecuperarDados(objServico.ID_SERVICO)
    
        If Not bRetorno Then
    
            GoTo Finalizar
    
        End If
    
        sCNL_PontaA = objCircuito.CNL_Ponta_A
        sDataAtivacao = IIf(objCircuito.DataAtivacao = "", "00000000", Format(objCircuito.DataAtivacao, "yyyymmdd"))
        sDataCancelamento = IIf(objCircuito.DataCancelamento = "", "00000000", Format(objCircuito.DataCancelamento, "yyyymmdd"))
    
        ' Recuperar os dados do centro funcional
        bRetorno = objCentroFuncional.RecuperarDados(objServico.ID_CENTRO_FUNCIONAL)
    
        If Not bRetorno Then
    
            GoTo Finalizar
    
        End If
    
        sCentroFuncional = objCentroFuncional.Codigo_Centro_Funcional
        
        ' Montar a variável de envio
        bRetorno = MontarVariavelEnvio()
    
        If objMensagem.TipoMsg <> "" Then
    
           GoTo Finalizar
    
        End If
        
        pbRegistroOK = True
    
    End If

    PreencherVariaveis = True

Finalizar:

    On Error GoTo 0

    Exit Function

ErrorHandler:

    If Err.Number <> 0 Then
       objMensagem.TipoMsg = "E"
       objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                   "Houve um erro de execução no Método: " & vbCrLf & vbCrLf & "PreencherVariaveis " & _
                   "da classe clsgenOC_InterfacePRF"
    End If

    GoTo Finalizar

End Function

Private Function MontarVariavelEnvio() As Boolean

On Error GoTo ErrorHandler:

    Dim bRetorno As Boolean
    Dim sFiller As String * 30

    MontarVariavelEnvio = False
    
    sDadosEnviado = sTipoRegistro
    sDadosEnviado = sDadosEnviado & sSiglaInstalacao
    sDadosEnviado = sDadosEnviado & sCodigoArea
    sDadosEnviado = sDadosEnviado & sCNL_PontaA
    sDadosEnviado = sDadosEnviado & sCNL_PontaB
    sDadosEnviado = sDadosEnviado & sDesignacaoFacilidade
    sDadosEnviado = sDadosEnviado & sDataAtivacao
    sDadosEnviado = sDadosEnviado & sDataAtivacao
    sDadosEnviado = sDadosEnviado & sDataCancelamento
    sDadosEnviado = sDadosEnviado & sSGA_NacionalInternacional
    sDadosEnviado = sDadosEnviado & sInstalacao
    sDadosEnviado = sDadosEnviado & sSiglaServico
    sDadosEnviado = sDadosEnviado & sIndicador
    sDadosEnviado = sDadosEnviado & sContaCorrente
    sDadosEnviado = sDadosEnviado & sVelocidadeCircuito
    sDadosEnviado = sDadosEnviado & sNumeroPedido
    sDadosEnviado = sDadosEnviado & sDesignacaoE1
    sDadosEnviado = sDadosEnviado & sCentroFuncional
    sDadosEnviado = sDadosEnviado & sCodigoAcessorio
    sDadosEnviado = sDadosEnviado & sQuantidadeAcessorio
    sDadosEnviado = sDadosEnviado & sTipoAtualizacao
    sDadosEnviado = sDadosEnviado & sAlteracaoEndereco
    sDadosEnviado = sDadosEnviado & sAlteracaoQuantidadeAcessorio
    sDadosEnviado = sDadosEnviado & sFiller
    sDadosEnviado = sDadosEnviado & sMatricula
        
    MontarVariavelEnvio = True

Finalizar:
    
    On Error GoTo 0
    
    Exit Function

ErrorHandler:

    If Err.Number <> 0 Then
       objMensagem.TipoMsg = "E"
       objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                   "Houve um erro de execução no Método: " & vbCrLf & vbCrLf & "MontarVariavelEnvio " & _
                   "da classe clsgenOC_InterfacePRF"
    End If
       
    GoTo Finalizar

End Function
