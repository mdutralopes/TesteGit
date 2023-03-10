VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsgenOC_PreAtivacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Declara as classes utilizadas
Private objMensagem        As clsgenMensagem
Private objPreAtivacao     As clsgenPreAtivacao
Private objFase            As clsgenFase
Private objServico         As clsGenServico
Private objSelecionarOTS As clsgenOC_SelecionarOTS

' Declara os formul?rios usados nesta classe
Private objFormPreAtivacao As frmwidPreAtivacao
Private objFormAnotacoes As frmgenAnotacoes

' Declara vari?veis
Private iSequenciaFase As Integer
Private sTipoGravacao As String

Public Sub TelaPreAtivacao(psTitulo As String, _
                           piSequenciaFase As Integer, _
                           psStatus As String)

On Error GoTo ErrorHandler:

    Dim bRetorno As Boolean
    Dim lUltimaFaseRealizada As Long
    Dim sSituacaoOTS As String
    Dim lID_FASE As Long
    
    ' Atualiza a vari?vel que guarda a sequencia
    iSequenciaFase = piSequenciaFase
    
    ' Inst?ncia o formul?rio de pr? ativa??o e chama em seguida
    Set objFormPreAtivacao = New frmwidPreAtivacao

    ' Atribui o conte?do do t?tulo
    objFormPreAtivacao.Caption = psTitulo
    
    ' Int?ncia a classe de mensagem
    Set objMensagem = New clsgenMensagem
    
    ' Informar o objeto controlador
    objFormPreAtivacao.InformaClasseOC Me
    
    ' Inst?ncia a classe de pr? ativa??o
    Set objPreAtivacao = New clsgenPreAtivacao
    
    ' Informar o nome do banco desejado
    objPreAtivacao.InformaBanco frmgenMDI_SGAPlus.objBanco
            
    ' Informar o nome do objeto de mensagem
    objPreAtivacao.InformaMensagem objMensagem
    
    ' Informar o nome do banco desejado
    objFase.InformaBanco frmgenMDI_SGAPlus.objBanco
    
    ' Inst?ncia a classe de servi?o
    Set objServico = New clsGenServico
        
    ' Informar o nome do banco desejado
    objServico.InformaBanco frmgenMDI_SGAPlus.objBanco
    
    ' Informar o nome da mensagem
    objServico.InformaMensagem objMensagem
    
    ' Recupera a situa??o da OTS
    sSituacaoOTS = objServico.RecuperarSituacaoOTS(objFase.ID_SERVICO)
    
    If objMensagem.TipoMsg <> "" Then
    
       Set objServico = Nothing
       
       GoTo Finalizar
       
    End If
    
    ' Verifica se qual foi a ?ltima fase realizada, se a fase for maior do que a fase atual
    ' N?o permitir a realiza??o do processo
    lUltimaFaseRealizada = objFase.ObtemUltimaFaseRealizada(objFase.ID_SERVICO)
    
    If objMensagem.TipoMsg = "E" Then
    
        GoTo Finalizar
        
    End If
    
    ' Atribui default para a vari?vel de identifica??o da fase com a propriedade da mesma
    lID_FASE = objFase.ID_FASE
    
    'Caso exista Servico Anterior e a fase ainda n?o foi iniciada,
    'Obtem os dados da OTS anterior
    If objFase.ID_SERVICO_ANTERIOR <> 0 And lUltimaFaseRealizada < objFase.ID_FASE Then
       
       ' (GbCodigoFaseWidPreAtivacao) - ID da fase na tabela Tipo_Fase
       lID_FASE = objServico.ObtemIDFase(objFase.ID_SERVICO_ANTERIOR, GbCodigoFaseWidPreAtivacao)
    
       If objMensagem.TipoMsg <> "" Then
              
          GoTo Finalizar
    
       End If
    
    End If
    
    Call CarregaCabecalho_Wid(objFormPreAtivacao.UsrWidCabecalho1, _
                              objFase.ID_SERVICO, _
                              objFase.ID_FASE, _
                              objMensagem)
        
    'Executa o m?todo para mostrar os dados, caso este exista
    bRetorno = MostraDados(lID_FASE)
        
    ' Se ? a primeira vez que entra coloca o a vari?vel que controla se inclus?o ou n?o como
    ' (I)
    If lID_FASE <> objFase.ID_FASE Then
    
        sTipoGravacao = "I"
        
    End If
    
    ' Verifica se qual foi a ?ltima fase realizada, se a fase for maior do que a fase atual
    ' N?o permitir a realiza??o do processo
    lUltimaFaseRealizada = objFase.ObtemUltimaFaseRealizada(objFase.ID_SERVICO)
    
    If objMensagem.TipoMsg = "E" Then
    
        GoTo Finalizar
        
    End If
    
    ' Se a situa??o estiver cancelada ou corrigida n?o deixa entrar
    If sSituacaoOTS = "CAN" Or sSituacaoOTS = "CR" Then
       
       objMensagem.TipoMsg = "A"
       objMensagem.Descricao = "Este servi?o foi paralisado pelo SOE!"
       
       lUltimaFaseRealizada = objFase.ID_FASE + 1
       
    End If
    
    ' Controla os objetos da tela se habilita ou n?o
    ControlaObjetosTela (lUltimaFaseRealizada)
    
    ' Atribui valor ao status da OTS
    objFormPreAtivacao.lblStatus.Caption = psStatus
    
    objFormPreAtivacao.Show
    
Finalizar:

    Screen.MousePointer = vbDefault
    
    If objMensagem.TipoMsg <> "" Then
    
        objMensagem.ExibeMensagem
        
    End If
    
    Set objMensagem = Nothing
    
    Exit Sub
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execu??o no m?todo: TelaPreAtivacao " & _
                            "da classe clsgenOC_PreAtivacao"
    GoTo Finalizar

End Sub

Public Sub TelaAnotacoes()

On Error GoTo ErrorHandler:

    Dim bRetorno As Boolean
     
    Screen.MousePointer = vbArrowHourglass
    
    ' Int?ncia a classe de mensagem
    Set objMensagem = New clsgenMensagem

    ' Inst?ncia o formul?rio de que exibe os n?meros de designa??o e OTS
    Set objFormAnotacoes = New frmgenAnotacoes
    
    ' Atribui o valor da propriedade no objeto da tela correspondente ao texto
    objFormAnotacoes.txtTexto = objFase.Anotacao
    
    ' Atribui o t?tulo da tela
    objFormAnotacoes.Caption = "Worldcom Internet Dedicated"
     
    objFormAnotacoes.Show vbModal
     
    Screen.MousePointer = vbDefault
    
    objFase.Anotacao = objFormAnotacoes.sTexto
        
Finalizar:

    Set objFormAnotacoes = Nothing
    
    Exit Sub
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execu??o no m?todo: TelaAnotacoes " & _
                            "da classe clswidOC_PreAtivacao"
    
    GoTo Finalizar

End Sub

Public Function MostraDados(ByVal plID_FASE As Long) As Boolean

On Error GoTo ErrorHandler:

    Dim bRetorno As Boolean
    
    ' Informar o nome do objeto de mensagem
    objPreAtivacao.InformaMensagem objMensagem
    
    sTipoGravacao = "I"
   
    ' Excutar o m?todo para recuperar os dados de pr? ativa??o
    bRetorno = objPreAtivacao.RecuperarDados(plID_FASE)
    
    ' Se ocorreu algum erro na DLL
    If Not bRetorno Then
        
        GoTo Finalizar
    
    End If
    
    ' Se o tipo de mensagem for = "A", ent?o encontrou registros
    If objMensagem.TipoMsg = "A" Then
    
        sTipoGravacao = "A" '
        
        objMensagem.TipoMsg = ""
    
    End If
    
    ' Executa o m?todo para recuperar os de anota??o da fase
    bRetorno = objFase.RecuperarAnotacao(plID_FASE)
    
    ' Se ocorreu algum erro na DLL
    If Not bRetorno Then
        
        GoTo Finalizar
    
    End If
    
    Call AtribuirCamposNaTela
    
    MostraDados = True
    
Finalizar:

    Screen.MousePointer = vbDefault
    
    Exit Function
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execu??o no m?todo: MostraDados " & _
                            "da classe clsgenOC_PreAtivacao"
    
    GoTo Finalizar

End Function

Public Function GravarDados() As Boolean

On Error GoTo ErrorHandler:

    Dim bRetorno As Boolean
    Dim bFinalizaFase As Boolean
    Dim lUltimaFaseRealizada As Long
    
    GravarDados = False
    
    Screen.MousePointer = vbArrowHourglass

    ' Verifica se finaliza a fase
    bFinalizaFase = (objFormPreAtivacao.spdDtExecTeste.Text > "")
    
    ' Int?ncia a classe de mensagem
    Set objMensagem = New clsgenMensagem
    
    ' Informar o nome do objeto de mensagem
    objPreAtivacao.InformaMensagem objMensagem
    
    ' Atribui os campos da tela nas propriedades
    Call AtribuirCamposDaTelaNasPropriedades
        
    ' Verifica se qual foi a ?ltima fase realizada, se a fase for maior do que a fase atual
    ' N?o permitir a realiza??o do processo
    lUltimaFaseRealizada = objFase.ObtemUltimaFaseRealizada(objFase.ID_SERVICO)
    
    If lUltimaFaseRealizada > objFase.ID_FASE Then
    
       objMensagem.TipoMsg = "A"
       objMensagem.Descricao = "A fase seguinte j? foi iniciada, n?o foi poss?vel executar " & _
                                "esta grava??o."
                                
       GravarDados = True
       
       GoTo Finalizar
        
    End If
    
    If objMensagem.TipoMsg = "E" Then
    
        GoTo Finalizar
        
    End If
    
    ' Inicia a transa??o
    bRetorno = frmgenMDI_SGAPlus.objBanco.Iniciar_Transacao()
    
        ' Se ocorreu algum erro na DLL
        If Not bRetorno Then
        
            GoTo Finalizar
            
        End If
    
        ' Excutar o m?todo para gravar os dados de pr? ativa??o
        bRetorno = objPreAtivacao.GravarDados(objFase.ID_FASE, _
                                              sTipoGravacao)
        
        ' Se ocorreu algum erro na DLL
        If Not bRetorno Then
        
            frmgenMDI_SGAPlus.objBanco.Cancelar_Transacao
        
            GoTo Finalizar
            
        End If
    
        ' Informa o banco para a classe de fase
        objFase.InformaBanco frmgenMDI_SGAPlus.objBanco
        
        ' Informa a mensagem para a classe de fase
        objFase.InformaMensagem objMensagem
        
        ' Grava a anota??o da fase
        bRetorno = objFase.GravarAnotacao(objFase.ID_FASE, _
                                          objFase.Anotacao)
        
        ' Ocorreu um erro ao executar o m?todo
        If Not bRetorno Then
    
            frmgenMDI_SGAPlus.objBanco.Cancelar_Transacao
    
            GoTo Finalizar

        End If
        
        ' Verifica se a fase j? foi iniciada, caso contr?rio grava da data de in?cio real
        If objFase.Data_Inicio_Real = "00:00:00" Then
        
            bRetorno = objFase.IniciaFase(objFase.ID_FASE)
        
            ' Ocorreu um erro ao executar o m?todo
            If Not bRetorno Then
        
                frmgenMDI_SGAPlus.objBanco.Cancelar_Transacao
        
                GoTo Finalizar
    
            End If
            
        End If
                
        If bFinalizaFase Then
           
            bRetorno = objFase.FinalizaFase(objFase.ID_FASE)
            
            ' Ocorreu um erro ao executar o m?todo
            If Not bRetorno Then
                
                frmgenMDI_SGAPlus.objBanco.Cancelar_Transacao
        
                GoTo Finalizar
            
            End If
            
            ' Calcula as datas previstas para a pr?xima fase
            bRetorno = objFase.CalcularProximaFase(objFase.ID_FASE)
                                                            
            ' Ocorreu um erro ao executar o m?todo
            If Not bRetorno Then
        
                frmgenMDI_SGAPlus.objBanco.Cancelar_Transacao
        
                GoTo Finalizar
    
            End If
            
            ' Atualiza a mensagem de retorno como grava??o realizada com sucesso
            objMensagem.TipoMsg = "A"
            objMensagem.Descricao = "Fase Finalizada com sucesso"
       
        Else
        
            ' Executa o m?todo
            bRetorno = objFase.RetirarFinalizacaoFase(objFase.ID_SERVICO, _
                                                      objFase.ID_FASE, _
                                                      objFase.SequenciaFase)
        
            ' Ocorreu um erro ao executar o m?todo
            If Not bRetorno Then
        
                frmgenMDI_SGAPlus.objBanco.Cancelar_Transacao
                
                GoTo Finalizar
            
            End If
            
            ' Atualiza a mensagem de retorno como grava??o realizada com sucesso
            objMensagem.TipoMsg = "A"
            objMensagem.Descricao = "Grava??o realizada com sucesso"
    
        End If
    
    ' Finaliza a transa??o
    frmgenMDI_SGAPlus.objBanco.Finalizar_Transacao
    
    GravarDados = True
    
Finalizar:

    Screen.MousePointer = vbDefault
    
    If objMensagem.TipoMsg <> "" Then
    
        objMensagem.ExibeMensagem
        
    End If
    
    Set objMensagem = Nothing
    
    Exit Function
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execu??o no m?todo: GravarDados " & _
                            "da classe clsgenOC_PreAtivacao"
    
    GoTo Finalizar

End Function

Public Function ExcluirDados() As Boolean

On Error GoTo ErrorHandler:

    Dim bRetorno As Boolean
    
    ExcluirDados = False
    
    Screen.MousePointer = vbArrowHourglass
    
    ' Int?ncia a classe de mensagem
    Set objMensagem = New clsgenMensagem
    
    ' Informar o nome do objeto de mensagem
    objPreAtivacao.InformaMensagem objMensagem
    
    ' Verifica se a fase esta pronta para ser grava
    bRetorno = objFase.VerificarFaseProntaIniciar(objFase.ID_FASE)
    
    If Not bRetorno Then
    
        GoTo Finalizar
        
    End If
    
    ' Inicia a transa??o
    frmgenMDI_SGAPlus.objBanco.Iniciar_Transacao
   
       ' Excutar o m?todo de exclus?o dos dados de pr? ativa??o
        bRetorno = objPreAtivacao.ExcluirDados(objFase.ID_FASE)
    
        ' Se ocorreu algum erro na DLL
        If Not bRetorno Then
        
            frmgenMDI_SGAPlus.objBanco.Cancelar_Transacao
    
            GoTo Finalizar
        
        End If
            
        ' Informar o nome do banco desejado
        objFase.InformaBanco frmgenMDI_SGAPlus.objBanco
        
        ' Informar o nome do objeto de mensagem
        objFase.InformaMensagem objMensagem
    
        ' Executa o m?todo
        bRetorno = objFase.RetirarFinalizacaoFase(objFase.ID_SERVICO, _
                                                      objFase.ID_FASE, _
                                                      objFase.SequenciaFase)
    
        ' Ocorreu um erro ao executar o m?todo
        If Not bRetorno Then
        
            frmgenMDI_SGAPlus.objBanco.Cancelar_Transacao
    
            GoTo Finalizar
        
        End If
        
        ' Prepara a mensagem de retorno para a tela
        objMensagem.TipoMsg = "A"
        objMensagem.Descricao = "Exclus?o realizada com sucesso"
    
    ' Finaliza a transa??o
    frmgenMDI_SGAPlus.objBanco.Finalizar_Transacao
    
    ExcluirDados = True

Finalizar:

    Screen.MousePointer = vbDefault
    
    If objMensagem.TipoMsg <> "" Then
    
        objMensagem.ExibeMensagem
        
    End If
    
    Set objMensagem = Nothing
    
    Exit Function
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execu??o no m?todo: ExcluirDados " & _
                            "da classe clsgenOC_PreAtivacao"
    
    Exit Function

End Function

Public Function LimparDados() As Boolean

On Error GoTo ErrorHandler:

    Dim bRetorno As Boolean
    
    LimparDados = False
   
    objFormPreAtivacao.spdDtPrevRealizacao.Text = ""
    objFormPreAtivacao.spdDtPedTeste.Text = ""
    objFormPreAtivacao.spdDtExecTeste.Text = ""
    objFormPreAtivacao.txtTecCampo = ""
    
    LimparDados = True

    Exit Function
        
ErrorHandler:

    MsgBox "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                   "Houve um erro de execu??o no m?todo: LimparDados " & _
                   "da classe clsgenOC_PreAtivacao"
    
    Exit Function

End Function

Public Function ControlaObjetosTela(ByVal piFase As Integer) As Boolean

    objFormPreAtivacao.framePreAtivacao.Enabled = (objFase.ID_FASE >= piFase) And _
                                                  ((objFase.Acao <> "DES") And _
                                                   (objFase.Acao <> "CAN"))
    objFormPreAtivacao.cmdSalvar.Enabled = (objFase.ID_FASE >= piFase) And _
                                           ((objFase.Acao <> "DES") And _
                                            (objFase.Acao <> "CAN"))
    objFormPreAtivacao.cmdLimpar.Enabled = (objFase.ID_FASE >= piFase) And _
                                           ((objFase.Acao <> "DES") And _
                                            (objFase.Acao <> "CAN"))
    objFormPreAtivacao.cmdExcluir.Enabled = (objFase.ID_FASE >= piFase) And _
                                            ((objFase.Acao <> "DES") And _
                                             (objFase.Acao <> "CAN"))
    objFormPreAtivacao.cmdAnotacao.Enabled = (objFase.ID_FASE >= piFase) And _
                                             ((objFase.Acao <> "DES") And _
                                              (objFase.Acao <> "CAN"))
                                              
End Function

Public Function InformaClasses(ByVal objNomeFase As clsgenFase)

On Error GoTo ErrorHandler:

    InformaClasses = False
    
    Set objFase = objNomeFase
    
    InformaClasses = True
    
Finalizar:

    Exit Function
    
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execu??o no m?todo: InformaClasses " & _
                            "da classe clsgenOC_PreAtivacao"
    
    GoTo Finalizar

End Function

Public Sub AtribuirCamposNaTela()

    ' Atribui as propriedades nos objetos da tela
    objFormPreAtivacao.spdDtPrevRealizacao.Text = Format(objPreAtivacao.DataPrevistaRealizacao, "dd/mm/yyyy")
    objFormPreAtivacao.spdDtPedTeste.Text = Format(objPreAtivacao.DataPedidoTeste, "dd/mm/yyyy")
    objFormPreAtivacao.spdDtExecTeste.Text = Format(objPreAtivacao.DataExecucaoTeste, "dd/mm/yyyy")
    objFormPreAtivacao.txtTecCampo = Format(objPreAtivacao.Responsavel, "dd/mm/yyyy")

End Sub

Public Sub AtribuirCamposDaTelaNasPropriedades()

    ' Atribui os objetos da tela nas propriedades
    If Len(objFormPreAtivacao.spdDtPrevRealizacao.Text) > 0 Then
    
        objPreAtivacao.DataPrevistaRealizacao = objFormPreAtivacao.spdDtPrevRealizacao.Text
        
    End If
    If Len(objFormPreAtivacao.spdDtPedTeste.Text) > 0 Then
    
        objPreAtivacao.DataPedidoTeste = objFormPreAtivacao.spdDtPedTeste.Text
        
    End If
    If Len(objFormPreAtivacao.spdDtExecTeste.Text) > 0 Then
    
        objPreAtivacao.DataExecucaoTeste = objFormPreAtivacao.spdDtExecTeste.Text
        
    End If
    objPreAtivacao.Responsavel = TrataPlick(objFormPreAtivacao.txtTecCampo)

End Sub

Private Sub Class_Terminate()

    Set objFase = Nothing
    Set objServico = Nothing
    Set objPreAtivacao = Nothing
    Set objMensagem = Nothing

End Sub

Public Function CriticarCamposTela() As Boolean

On Error GoTo ErrorHandler:

    CriticarCamposTela = False
    
    ' Int?ncia a classe de mensagem
    Set objMensagem = New clsgenMensagem
    
    ' Criticar a data de pedido do teste, caso esteja preenchida
    If (Len(objFormPreAtivacao.spdDtPedTeste.Text) > 0) Then
    
        ' Critica se a data de pedido de teste ? v?lida
        If (Len(objFormPreAtivacao.spdDtPedTeste.Text) > 0) And _
           (Len(objFormPreAtivacao.spdDtPedTeste.Text) < 10) Then
        
            objMensagem.TipoMsg = "A"
            objMensagem.Descricao = "Data do pedido de teste inv?lida."
            objFormPreAtivacao.spdDtPedTeste.SetFocus
        
            GoTo Finalizar
    
        ' Critica a data de pedido de teste ? diferente do dia atual
        ElseIf objFormPreAtivacao.spdDtPedTeste.Text < Format(Date, "dd/mm/yyyy") Then
            
            objMensagem.TipoMsg = "A"
            objMensagem.Descricao = "Data do pedido de teste n?o pode ser diferente da data atual."
            objFormPreAtivacao.spdDtPedTeste.SetFocus
            
            GoTo Finalizar
            
        End If
    
    End If
    
    ' Criticar a data de previs?o de realiza??o, caso esteja preenchida, pois esta
    ' data quando preenchida finaliza a fase, por este motivo necessita de v?rias
    ' cr?ticas
    If (Len(objFormPreAtivacao.spdDtPrevRealizacao.Text) > 0) Then

        ' Critica se a data de previs?o de realiza??o ? v?lida
        If (Len(objFormPreAtivacao.spdDtPrevRealizacao.Text) > 0) And _
           (Len(objFormPreAtivacao.spdDtPrevRealizacao.Text) < 10) Then
        
            objMensagem.TipoMsg = "A"
            objMensagem.Descricao = "Data de prevista para realiza??o inv?lida."
            objFormPreAtivacao.spdDtPrevRealizacao.SetFocus
        
            GoTo Finalizar
            
        ' Critica a data de previs?o de realiza??o ? diferente do dia atual
        ElseIf objFormPreAtivacao.spdDtPrevRealizacao.Text < Format(Date, "dd/mm/yyyy") Then
            
            objMensagem.TipoMsg = "A"
            objMensagem.Descricao = "Data de prevista para realiza??o n?o pode ser diferente da data atual."
            objFormPreAtivacao.spdDtPrevRealizacao.SetFocus
            
            GoTo Finalizar
            
        End If
    
    End If
    
    ' Criticar a data de execu??o do teste, caso esteja preenchida
    If (Len(objFormPreAtivacao.spdDtExecTeste.Text) > 0) Then
    
        ' Critica se a data de execu??o do teste ? v?lida
        If (Len(objFormPreAtivacao.spdDtExecTeste.Text) > 0) And _
           (Len(objFormPreAtivacao.spdDtExecTeste.Text) < 10) Then
        
            objMensagem.TipoMsg = "A"
            objMensagem.Descricao = "Data de execu??o do teste inv?lida."
            objFormPreAtivacao.spdDtPedTeste.SetFocus
        
            GoTo Finalizar
            
        ' Neste caso, exige que a data de pedido do teste, prevista para realiza??o e
        ' respons?vel sejam obirgat?rios
        '
        ' Data de pedido de teste
        ElseIf (Len(objFormPreAtivacao.spdDtPedTeste.Text) = 0) Then
        
            objMensagem.TipoMsg = "A"
            objMensagem.Descricao = "Data do pedido de teste obrigat?ria."
            objFormPreAtivacao.spdDtPedTeste.SetFocus
        
            GoTo Finalizar
    
        ' Data de prevista para realiza??o
        ElseIf (Len(objFormPreAtivacao.spdDtPrevRealizacao.Text) = 0) Then
        
            objMensagem.TipoMsg = "A"
            objMensagem.Descricao = "Data da prevista para realiza??o obrigat?ria."
            objFormPreAtivacao.spdDtPrevRealizacao.SetFocus
        
            GoTo Finalizar
            
        ' Respons?vel t?cnico
        ElseIf (Len(objFormPreAtivacao.txtTecCampo.Text) = 0) Then
        
            objMensagem.TipoMsg = "A"
            objMensagem.Descricao = "T?cnico respons?vel obrigat?rio."
            objFormPreAtivacao.txtTecCampo.SetFocus
        
            GoTo Finalizar
        
        ' Critica a data de execu??o do teste ? diferente do dia atual
        ElseIf objFormPreAtivacao.spdDtExecTeste.Text < Format(Date, "dd/mm/yyyy") Then
            
            objMensagem.TipoMsg = "A"
            objMensagem.Descricao = "Data de execu??o do teste n?o pode ser diferente da data atual."
            objFormPreAtivacao.spdDtExecTeste.SetFocus
            
            GoTo Finalizar
            
        ' Faz as cr?ticas do controle das datas, verificando se uma ? maior do que a outra
        '
        '
        ' Data de previs?o de teste n?o pode ser menor do que a data do pedido de teste
        ElseIf objFormPreAtivacao.spdDtPrevRealizacao.Text < objFormPreAtivacao.spdDtPedTeste.Text Then
        
            objMensagem.TipoMsg = "A"
            objMensagem.Descricao = "Data de prevista para realiza??o n?o pode ser menor do que a do pedido de teste."
            objFormPreAtivacao.spdDtPrevRealizacao.SetFocus
        
            GoTo Finalizar
            
        ' Data de de execu??o do teste n?o pode ser menor do que a data de previs?o de teste
        ElseIf objFormPreAtivacao.spdDtExecTeste.Text < objFormPreAtivacao.spdDtPrevRealizacao.Text Then
        
            objMensagem.TipoMsg = "A"
            objMensagem.Descricao = "Data de execu??o do teste n?o pode ser menor do que a de previs?o do teste."
            objFormPreAtivacao.spdDtExecTeste.SetFocus
        
            GoTo Finalizar
        
        End If
    
    End If
    
    CriticarCamposTela = True

Finalizar:

    If objMensagem.TipoMsg <> "" Then
    
        Screen.MousePointer = vbDefault
        
        objMensagem.ExibeMensagem
        
    End If
    
    Set objMensagem = Nothing
    
    Exit Function
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execu??o no m?todo: CriticarCamposTela " & _
                            "da classe clsgenOC_PreAtivacao"
    
    GoTo Finalizar

End Function
Public Function SaiTela() As Boolean

On Error GoTo ErrorHandler:

    SaiTela = False

    ' Chama o m?todo para verificar se foram alterados alguns dados
    If AlterouDados Then
       
       objMensagem.TipoMsg = "Q"
       objMensagem.Descricao = "Foram feitas altera??es nos dados da tela que " & _
                               "ainda n?o foram Salvas! Deseja sair assim mesmo ?"
       
       
    End If
    
    SaiTela = True
    
Finalizar:

    If objMensagem.TipoMsg <> "" Then
    
        Screen.MousePointer = vbDefault
    
        SaiTela = objMensagem.ExibeMensagem = "S"
    
    End If
    
    Exit Function
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execu??o no m?todo: SaiTela " & _
                            "da classe clswidOC_PreAtivacao"
    
    GoTo Finalizar

End Function

Public Function AlterouDados() As Boolean

On Error GoTo ErrorHandler:

    AlterouDados = True
    
    Set objMensagem = New clsgenMensagem
    
    If objPreAtivacao.DataPrevistaRealizacao <> objFormPreAtivacao.spdDtPrevRealizacao.Text Then
            
        GoTo Finalizar
    
    End If
        
    If objPreAtivacao.DataPedidoTeste <> objFormPreAtivacao.spdDtPedTeste.Text Then
        
        GoTo Finalizar
    
    End If
    
    If objPreAtivacao.DataExecucaoTeste <> objFormPreAtivacao.spdDtExecTeste.Text Then
        
        GoTo Finalizar
    
    End If
    
    If objPreAtivacao.Responsavel <> TrataPlick(objFormPreAtivacao.txtTecCampo) Then
    
        GoTo Finalizar
    
    End If

    AlterouDados = False
    
Finalizar:

    Exit Function
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execu??o no m?todo: AlterouDados " & _
                            "da classe clswidOC_PreAtivacao"
    
    GoTo Finalizar

End Function

Public Sub InformaClasseOC(ByVal objNomeClasseOC As clsgenOC_SelecionarOTS)

    Set objSelecionarOTS = objNomeClasseOC
    
End Sub

Public Sub Saida()
    
    objSelecionarOTS.ExibirListaOTS
    
    Unload objFormPreAtivacao

End Sub

Public Sub MontarImpressaoOTS()

On Error GoTo ErrorHandler:

    Dim bRetorno As Boolean

    ' Int?ncia a classe de mensagem
    Set objMensagem = New clsgenMensagem
    
    bRetorno = ImprimirOTS_WID(objMensagem, _
                               objFase.ID_SERVICO, _
                               objFase.ID_SERVICO_ANTERIOR, _
                               objFase.Acao)

    ' Se ocorreu algum erro na DLL exibi-lo
    If Not bRetorno Then
        
        GoTo Finalizar
    
    End If

Finalizar:

    Screen.MousePointer = vbDefault
    
    If objMensagem.TipoMsg <> "" Then
    
        objMensagem.ExibeMensagem
        
    End If
    
    On Error GoTo 0
    
    Set objMensagem = Nothing
    
    Exit Sub
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execu??o no m?todo: MontarImpressaoOTS " & _
                            "da classe clsgenOC_Ativa??o"
    
    GoTo Finalizar

End Sub

