VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsgenOC_CadastrarTipoInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Declara as classes utilizadas
Private objMensagem                     As clsgenMensagem
Private objCadastrarTipoInterface       As clsgenTipoInterface
Private objListaTipoInterface           As clsgenListaTipoInterface

' Declara os formul?rios usados nesta classe
Private objFormCadastrarTipoInterface As frmgenCadastroTipoInterface

'Private clsgenOC_CadastrarTipoInterface   As frmgenCadastroTipoInterface

' Declara vari?veis comuns
Private lID              As Long
Private sTipoOperacao                   As String
Private lQuantidadeRegistros            As Long

Public Property Get TipoOperacao() As String

    TipoOperacao = sTipoOperacao
    
End Property

Public Property Let TipoOperacao(ByVal sNewValue As String)

    sTipoOperacao = sNewValue
    
    ' Chama  m?todo que controla a habilita??o e desabilita??o dos objetos,
    ' conforme o tipo de opera??o
    Call HabilitaDesabilitaObjetos
    
End Property

Public Sub TelaCadastrarTipoInterfaces()

On Error GoTo ErrorHandler:

    ' Inst?ncia o formul?rio de pr? ativa??o e chama em seguida
    Set objFormCadastrarTipoInterface = New frmgenCadastroTipoInterface
    
    ' Int?ncia a classe de mensagem
    Set objMensagem = New clsgenMensagem
    
    ' Informar o objeto controlador
    objFormCadastrarTipoInterface.InformaClasseOC Me
    
    ' Inst?ncia a classe de dados para cadastrar usu?rios
    Set objCadastrarTipoInterface = New clsgenTipoInterface
    
    ' Informar o nome do banco desejado
    objCadastrarTipoInterface.InformaBanco frmgenMDI_SGAPlus.objBanco
            
    ' Informar o nome do objeto de mensagem
    objCadastrarTipoInterface.InformaMensagem objMensagem
    
    ' Chama o m?todo de exibi??o do GRID
    Call ExibirLista
    
Finalizar:

    Screen.MousePointer = vbDefault
    
    If objMensagem.TipoMsg <> "" Then
    
        objMensagem.ExibeMensagem
        
    End If
    
    Set objMensagem = Nothing
    
    On Error GoTo 0
    
    Exit Sub
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execu??o no m?todo: TelaCadastrarTipoInterfaces " & _
                            "da classe clsgenOC_CadastrarTipoInterface"
    GoTo Finalizar

End Sub

Public Sub ExibirLista()

On Error GoTo ErrorHandler:

    Dim bRetorno As Boolean
    
    Screen.MousePointer = vbArrowHourglass
    
    lQuantidadeRegistros = 0
    
    ' Int?ncia a classe de mensagem
    Set objMensagem = New clsgenMensagem
    
    ' Informar o nome do objeto de mensagem
    objCadastrarTipoInterface.InformaMensagem objMensagem
    
    ' Int?ncia a classe da lista de usu?rios
    Set objListaTipoInterface = Nothing
    Set objListaTipoInterface = New clsgenListaTipoInterface
    
    ' Limpa a lista
    objFormCadastrarTipoInterface.lstTipoInterface.Clear
    
    ' Executar o m?todo de montagem da lista na classe de analisar order entry
    bRetorno = objCadastrarTipoInterface.MontarLista(objFormCadastrarTipoInterface.optConsultaEmbratel.Value, _
                                                     objFormCadastrarTipoInterface.optConsultaCliente.Value, _
                                                     objFormCadastrarTipoInterface.txtBusca.Text)
    
    
    ' Se ocorreu algum erro na DLL exibi-lo
    If Not bRetorno Then

        GoTo Finalizar

    End If
    
    ' Preenche a lista de usu?rios
    bRetorno = PreencheLista()
    
    If Not bRetorno Then
        
        GoTo Finalizar
    
    End If
    
    ' Posiciona o na primeira lista caso encontre algum registro
    objFormCadastrarTipoInterface.lstTipoInterface.ListIndex = 0
    
Finalizar:

    ' Se n?o encontrou registro desabilitar os bot?es de altera??o e exclus?o
    If lQuantidadeRegistros = 0 Then
    
        objFormCadastrarTipoInterface.cmdAlterar.Enabled = False
        objFormCadastrarTipoInterface.cmdExcluir.Enabled = False
        
        Call LimparDadosTelaTipoInterface
        
    Else
    
        objFormCadastrarTipoInterface.cmdAlterar.Enabled = True
        objFormCadastrarTipoInterface.cmdExcluir.Enabled = True
        
    End If
    
    Screen.MousePointer = vbDefault
    
    If objMensagem.TipoMsg <> "" Then
    
        objMensagem.ExibeMensagem
        
    End If
    
    objMensagem.TipoMsg = ""
    
    On Error GoTo 0
    
    Exit Sub
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execu??o no m?todo: ExibirLista " & _
                            "da classe clswidOC_CadastrarTipoInterface"

    GoTo Finalizar

End Sub

Public Function GravarDados(ByVal pyTela As Byte) As Boolean

On Error GoTo ErrorHandler:

    Dim bRetorno As Boolean

    GravarDados = False

    ' Int?ncia a classe de mensagem
    Set objMensagem = New clsgenMensagem

    ' Inicia a transa??o
    bRetorno = frmgenMDI_SGAPlus.objBanco.Iniciar_Transacao()

        ' Se ocorreu algum erro na DLL
        If Not bRetorno Then

            GoTo Finalizar

        End If

        ' Criticar os campos da tela
        bRetorno = CriticarCamposTelaTipoInterface
            
        If Not bRetorno Then
        
            GoTo Finalizar
            
        End If
    
        ' Informar o nome do objeto de mensagem
        objCadastrarTipoInterface.InformaMensagem objMensagem
        
        ' Atribui os campos da tela nas propriedades
        Call AtribuirCamposNasPropriedades
            
        ' Excutar o m?todo para gravar os dados do usu?rio
        bRetorno = objCadastrarTipoInterface.GravarDados(sTipoOperacao)
            
        ' Se ocorreu algum erro na DLL
        If Not bRetorno Then

            GoTo Finalizar

        End If

    ' Finaliza a transa??o
    frmgenMDI_SGAPlus.objBanco.Finalizar_Transacao
            
    ' Atualiza a mensagem de retorno como grava??o realizada com sucesso
    objMensagem.TipoMsg = "A"
    objMensagem.Descricao = "Grava??o realizada com sucesso"

    GravarDados = True

Finalizar:

    Screen.MousePointer = vbDefault

    If objMensagem.TipoMsg <> "" Then

        objMensagem.ExibeMensagem

    End If
    
    If Not GravarDados Then
    
        frmgenMDI_SGAPlus.objBanco.Cancelar_Transacao

    End If

    Set objMensagem = Nothing
    
    On Error GoTo 0

    Exit Function

ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execu??o no m?todo: GravarDados " & _
                            "da classe clsgenOC_CadastrarTipoInterface"

    GoTo Finalizar

End Function

Public Sub ExcluirDados()

On Error GoTo ErrorHandler:

    Dim bRetorno As Boolean

    ' Int?ncia a classe de mensagem
    Set objMensagem = New clsgenMensagem
    
    objMensagem.TipoMsg = "Q"
    objMensagem.Descricao = "Confirma a exlus?o do Tipo de Interface selecionado?"
    If objMensagem.ExibeMensagem = "S" Then

        ' Informar o nome do objeto de mensagem
        objCadastrarTipoInterface.InformaMensagem objMensagem
    
        ' Inicia a transa??o
        frmgenMDI_SGAPlus.objBanco.Iniciar_Transacao
    
           ' Excutar o m?todo de exclus?o dos dados do usu?rio
            bRetorno = objCadastrarTipoInterface.ExcluirDados(objCadastrarTipoInterface.ID)
    
            ' Se ocorreu algum erro na DLL
            If Not bRetorno Then
    
                frmgenMDI_SGAPlus.objBanco.Cancelar_Transacao
    
                GoTo Finalizar
    
            End If
    
        ' Finaliza a transa??o
        frmgenMDI_SGAPlus.objBanco.Finalizar_Transacao
        
        ' Se existia somente um registro limpar o objeto de entrada para buscar
        If lQuantidadeRegistros = 1 Then
        
            ' Limpa o campo de busca
            objFormCadastrarTipoInterface.txtBusca.Text = ""
            
        End If
        
        ' Chama o m?todo de exibi??o do GRID
        Call ExibirLista
        
    Else
    
        objMensagem.TipoMsg = ""
        
    End If
    
Finalizar:

    Screen.MousePointer = vbDefault

    If objMensagem.TipoMsg <> "" Then

        objMensagem.ExibeMensagem

    End If

    Set objMensagem = Nothing
    
    On Error GoTo 0

    Exit Sub

ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execu??o no m?todo: ExcluirDados " & _
                            "da classe clsgenOC_CadastrarTipoInterface"

    Exit Sub

End Sub

Private Sub AtribuirCamposNasPropriedades()

    objCadastrarTipoInterface.Descricao = TrataPlick(objFormCadastrarTipoInterface.TxtDescricao)
'    objCadastrarTipoInterface.Tipo_Interface =
    If objFormCadastrarTipoInterface.OptEmbratel.Value = True Then
    
        objCadastrarTipoInterface.Tipo_Interface = "E"
    
    ElseIf objFormCadastrarTipoInterface.optCliente.Value = True Then
    
        objCadastrarTipoInterface.Tipo_Interface = "C"
    
    End If

End Sub

Private Sub Class_Terminate()

    Set objCadastrarTipoInterface = Nothing
    Set objMensagem = Nothing
    Set objListaTipoInterface = Nothing
    
End Sub

Private Function CriticarCamposTelaTipoInterface() As Boolean

On Error GoTo ErrorHandler:

    CriticarCamposTelaTipoInterface = False
    
'    If Len(clsgenOC_CadastrarTipoInterface.txtCodigo_Controle.Text) = 0 Then
''        objMensagem.TipoMsg = "A"
''        objMensagem.Descricao = "C?digo de Controle ? obrigat?ria."
''        clsgenOC_CadastrarTipoInterface.txtCodigo_Controle.SetFocus
''
'        GoTo Finalizar
'    Else
    
    If Len(objFormCadastrarTipoInterface.TxtDescricao.Text) = 0 Then
        objMensagem.TipoMsg = "A"
        objMensagem.Descricao = "Descri??o ? obrigat?ria."
        objFormCadastrarTipoInterface.TxtDescricao.SetFocus
        
        GoTo Finalizar
    End If
    
    If ((objFormCadastrarTipoInterface.OptEmbratel.Value = False) And (objFormCadastrarTipoInterface.optCliente.Value = False)) Then
        objMensagem.TipoMsg = "A"
        objMensagem.Descricao = "Tipo de Interface ? obrigat?ria."
        objFormCadastrarTipoInterface.TxtDescricao.SetFocus
        
        GoTo Finalizar
    End If
    
    CriticarCamposTelaTipoInterface = True

Finalizar:

    If objMensagem.TipoMsg <> "" Then
    
        Screen.MousePointer = vbDefault
        
        objMensagem.ExibeMensagem
        
    End If
    
    On Error GoTo 0
    
    Exit Function
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execu??o no m?todo: CriticarCamposTelaTipoInterface " & _
                            "da classe clsgenOC_CadastrarTipoInterface"
    
    GoTo Finalizar

End Function

Public Function SaiTela(ByVal pyTela As Byte) As Boolean

On Error GoTo ErrorHandler:

    Dim bAlterou As Boolean

    SaiTela = False

    bAlterou = AlterouDadosTelaTipoInterface
    
    ' Chama o m?todo para verificar se foram alterados alguns dados
    If bAlterou Then
       
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
    
    On Error GoTo 0
    
    Exit Function
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execu??o no m?todo: SaiTela " & _
                            "da classe clsgenOC_CadastrarTipoInterfaces"
    
    GoTo Finalizar

End Function

Public Function AlterouDadosTelaTipoInterface() As Boolean

On Error GoTo ErrorHandler:

    AlterouDadosTelaTipoInterface = True
    
    Set objMensagem = New clsgenMensagem
    
    If objCadastrarTipoInterface.Descricao <> objFormCadastrarTipoInterface.TxtDescricao Then
    
        GoTo Finalizar
        
    End If

    AlterouDadosTelaTipoInterface = False
    
Finalizar:

    On Error GoTo 0

    Exit Function
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execu??o no m?todo: AlterouDadosTelaTipoInterface " & _
                            "da classe clsgenOC_CadastrarTipoInterfaces"
    
    GoTo Finalizar

End Function

Private Function PreencheLista() As Boolean

    lQuantidadeRegistros = 0
    
    PreencheLista = False
    
    ' Ler o objeto de cole??o da lista de tipos de pend?ncias
    For Each objListaTipoInterface In objCadastrarTipoInterface.ListaTipoInterface

        With objFormCadastrarTipoInterface
            
            ' Descricao Motivo Recusa
             .lstTipoInterface.AddItem (objListaTipoInterface.Descricao)
        
        End With

        lQuantidadeRegistros = lQuantidadeRegistros + 1
    
    Next
    
    PreencheLista = True
    
End Function

Public Sub RefreshDados()

    If lQuantidadeRegistros > 0 Then
    ' Atualiza os dados de usu?rio na tela
    With objFormCadastrarTipoInterface
        
        .TxtDescricao = objCadastrarTipoInterface.ListaTipoInterface(objFormCadastrarTipoInterface.lstTipoInterface.ListIndex + 1).Descricao
        
        If (objCadastrarTipoInterface.ListaTipoInterface(objFormCadastrarTipoInterface.lstTipoInterface.ListIndex + 1).Tipo_Interface) = "E" Then
        
        .OptEmbratel.Value = True
        
        ElseIf (objCadastrarTipoInterface.ListaTipoInterface(objFormCadastrarTipoInterface.lstTipoInterface.ListIndex + 1).Tipo_Interface) = "C" Then
        
            .optCliente.Value = True
        
        End If
        
    End With
    
        objCadastrarTipoInterface.Descricao = objCadastrarTipoInterface.ListaTipoInterface(objFormCadastrarTipoInterface.lstTipoInterface.ListIndex + 1).Descricao
        objCadastrarTipoInterface.Tipo_Interface = objCadastrarTipoInterface.ListaTipoInterface(objFormCadastrarTipoInterface.lstTipoInterface.ListIndex + 1).Tipo_Interface
        objCadastrarTipoInterface.ID = objCadastrarTipoInterface.ListaTipoInterface(objFormCadastrarTipoInterface.lstTipoInterface.ListIndex + 1).ID
    End If


End Sub

Private Sub HabilitaDesabilitaObjetos()

    ' Coloca cor de fundo como default para habilita??o
    If (sTipoOperacao = "I") Or (sTipoOperacao = "A") Then ' Inclus?o ou altera??o
        
        ' Muda cor de fundo dos objetos de entrada de dados para branco
        objFormCadastrarTipoInterface.TxtDescricao.BackColor = &H80000005

        
        ' Muda cor de escrita dos objetos de entrada de dados para preto
        objFormCadastrarTipoInterface.TxtDescricao.ForeColor = &H80000008

        
        ' Habilita o frame dos objetos de dados
        objFormCadastrarTipoInterface.frameDados.Enabled = True
        
        ' Desabilita bot?es de inclus?o, altera??o, exlus?o e frames de busca e lista
        objFormCadastrarTipoInterface.cmdIncluir.Enabled = False
        objFormCadastrarTipoInterface.cmdAlterar.Enabled = False
        objFormCadastrarTipoInterface.cmdExcluir.Enabled = False
        objFormCadastrarTipoInterface.frameBusca.Enabled = False
        objFormCadastrarTipoInterface.frameLista.Enabled = False
        
        ' Habilita os bot?es gravar, limpar e cancelar
        objFormCadastrarTipoInterface.cmdSalvar.Enabled = True
        objFormCadastrarTipoInterface.cmdLimpar.Enabled = True
        objFormCadastrarTipoInterface.cmdCancelar.Enabled = True
        
    Else
    
        ' Muda cor de fundo dos objetos de entrada de dados para cinza
        objFormCadastrarTipoInterface.TxtDescricao.BackColor = &H80000004

        
        ' Muda cor de escrita dos objetos de entrada de dados azul
        objFormCadastrarTipoInterface.TxtDescricao.ForeColor = &HC00000
        
        ' Desabilita o frame dos objetos de dados
        objFormCadastrarTipoInterface.frameDados.Enabled = False
    
        ' Habilita bot?es de inclus?o, altera??o, exlus?o e frames de busca e lista
        objFormCadastrarTipoInterface.cmdIncluir.Enabled = True
        objFormCadastrarTipoInterface.cmdAlterar.Enabled = True
        objFormCadastrarTipoInterface.cmdExcluir.Enabled = True
        objFormCadastrarTipoInterface.frameBusca.Enabled = True
        objFormCadastrarTipoInterface.frameLista.Enabled = True
        
        ' Desabilita os bot?es gravar, limpar e cancelar
        objFormCadastrarTipoInterface.cmdSalvar.Enabled = False
        objFormCadastrarTipoInterface.cmdLimpar.Enabled = False
        objFormCadastrarTipoInterface.cmdCancelar.Enabled = False
    
    End If

End Sub

Public Sub LimparDadosTelaTipoInterface()

    objFormCadastrarTipoInterface.TxtDescricao.Text = ""
    objFormCadastrarTipoInterface.OptEmbratel.Value = False
    objFormCadastrarTipoInterface.optCliente.Value = False
    
    objCadastrarTipoInterface.Descricao = ""
    objCadastrarTipoInterface.Tipo_Interface = ""

    
    Exit Sub
        
End Sub

Public Sub Saida(ByVal pyTela As Byte)
    
    Unload objFormCadastrarTipoInterface
    
End Sub


