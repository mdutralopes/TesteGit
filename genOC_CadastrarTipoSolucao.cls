VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsgenOC_CadastrarTipoSolucao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Declara as classes utilizadas
'Private objCadPend_Prod_Fase            As clsgenPend_Prod_Fases
'Private lID_TIPO_SOLUCAO
'Private objFormCadPend_Prod_Fase        As frmgenCadPend_Prod_Fase


Private objMensagem                     As clsgenMensagem
Private objCadastrarTipoSolucao         As clsgenTipoSolucao
Private objListaTipoSolucao             As clsgenListaTipoSolucao

' Declara o formulário usado nesta classe
Private objFormCadastrarTipoSolucao     As frmgenCadastroTipoSolucao

' Declara variáveis comuns
Private lID_TIPO_SOLUCAO                As Long
Private sTipoOperacao                   As String
Private lQuantidadeRegistros            As Long
Private sTipoSolucao                    As String
Private aListaTipSolucao()              As Variant

Public Property Get TipoSolucao() As String

    TipoSolucao = sTipoSolucao
    
End Property
Public Property Let TipoSolucao(ByVal sNewValue As String)

    sTipoSolucao = sNewValue
    
End Property

Public Property Get TipoOperacao() As String

    TipoOperacao = sTipoOperacao
    
End Property

Public Property Let TipoOperacao(ByVal sNewValue As String)

    sTipoOperacao = sNewValue
    
    ' Chama  método que controla a habilitação e desabilitação dos objetos,
    ' conforme o tipo de operação
    Call HabilitaDesabilitaObjetos
    
End Property

Public Sub TelaCadastrarTipoSolucao()

On Error GoTo ErrorHandler:

    ' Instância o formulário de pré ativação e chama em seguida
    Set objFormCadastrarTipoSolucao = New frmgenCadastroTipoSolucao
    
    ' Intância a classe de mensagem
    Set objMensagem = New clsgenMensagem
    
    ' Informar o objeto controlador
    objFormCadastrarTipoSolucao.InformaClasseOC Me
    
    ' Instância a classe de dados para cadastrar usuários
    Set objCadastrarTipoSolucao = New clsgenTipoSolucao
    
    ' Informar o nome do banco desejado
    objCadastrarTipoSolucao.InformaBanco frmgenMDI_SGAPlus.objBanco
            
    ' Informar o nome do objeto de mensagem
    objCadastrarTipoSolucao.InformaMensagem objMensagem
    
    ' Chama o método de exibição do GRID
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
                            "Houve um erro de execução no método: TelaCadastrarTipoSolucao " & _
                            "da classe clsgenOC_CadastrarTipoSolucao"
    GoTo Finalizar

End Sub

Public Sub ExibirLista()

On Error GoTo ErrorHandler:

    Dim bRetorno As Boolean
    
    Screen.MousePointer = vbArrowHourglass
    
    lQuantidadeRegistros = 0
    
    ' Intância a classe de mensagem
    Set objMensagem = New clsgenMensagem
    
    ' Informar o nome do objeto de mensagem
    objCadastrarTipoSolucao.InformaMensagem objMensagem
    
    ' Intância a classe da lista de usuários
    Set objListaTipoSolucao = Nothing
    Set objListaTipoSolucao = New clsgenListaTipoSolucao
    
    ' Limpa a lista
    objFormCadastrarTipoSolucao.lstTipoSolucao.Clear
    
      
    ' Executar o método de montagem da lista na classe de analisar order entry
    bRetorno = objCadastrarTipoSolucao.MontarLista(objFormCadastrarTipoSolucao.txtBusca.Text)
    
    
    ' Se ocorreu algum erro na DLL exibi-lo
    If Not bRetorno Then

        GoTo Finalizar

    End If
    
    ' Preenche a lista de usuários
    bRetorno = PreencheLista()
    
    If Not bRetorno Then
        
        GoTo Finalizar
    
    End If
    
    ' Posiciona o na primeira lista caso encontre algum registro
    objFormCadastrarTipoSolucao.lstTipoSolucao.ListIndex = 0
    
Finalizar:

    ' Se não encontrou registro desabilitar os botôes de alteração e exclusão
    If lQuantidadeRegistros = 0 Then
    
        objFormCadastrarTipoSolucao.cmdAlterar.Enabled = False
        objFormCadastrarTipoSolucao.cmdExcluir.Enabled = False
        
        Call LimparDadosTelaTipoSolucao
        
    Else
    
        objFormCadastrarTipoSolucao.cmdAlterar.Enabled = True
        objFormCadastrarTipoSolucao.cmdExcluir.Enabled = True
        
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
                            "Houve um erro de execução no método: ExibirLista " & _
                            "da classe clsgenOC_CadastrarTipoSolucao"

    GoTo Finalizar

End Sub

Public Function GravarDados(ByVal pyTela As Byte) As Boolean

On Error GoTo ErrorHandler:

    Dim bRetorno As Boolean

    GravarDados = False

    ' Intância a classe de mensagem
    Set objMensagem = New clsgenMensagem

    ' Inicia a transação
    bRetorno = frmgenMDI_SGAPlus.objBanco.Iniciar_Transacao()

        ' Se ocorreu algum erro na DLL
        If Not bRetorno Then

            GoTo Finalizar

        End If

        ' Criticar os campos da tela
        bRetorno = CriticarCamposTelaTipoSolucao
            
        If Not bRetorno Then
        
            GoTo Finalizar
            
        End If
    
        ' Informar o nome do objeto de mensagem
        objCadastrarTipoSolucao.InformaMensagem objMensagem
        
        ' Atribui os campos da tela nas propriedades
        Call AtribuirCamposNasPropriedades
            
        ' Excutar o método para gravar os dados do usuário
        bRetorno = objCadastrarTipoSolucao.GravarDados(sTipoOperacao)
            
        ' Se ocorreu algum erro na DLL
        If Not bRetorno Then

            GoTo Finalizar

        End If

    ' Finaliza a transação
    frmgenMDI_SGAPlus.objBanco.Finalizar_Transacao
            
    ' Atualiza a mensagem de retorno como gravação realizada com sucesso
    objMensagem.TipoMsg = "A"
    objMensagem.Descricao = "Gravação realizada com sucesso"

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
                            "Houve um erro de execução no método: GravarDados " & _
                            "da classe clsgenOC_CadastrarTipoSolucao"

    GoTo Finalizar

End Function

Public Sub ExcluirDados()

On Error GoTo ErrorHandler:

    Dim bRetorno As Boolean

    ' Intância a classe de mensagem
    Set objMensagem = New clsgenMensagem
    
    objMensagem.TipoMsg = "Q"
    objMensagem.Descricao = "Confirma a exlusão do Tipo de Solução selecionado?"
    If objMensagem.ExibeMensagem = "S" Then

        ' Informar o nome do objeto de mensagem
        objCadastrarTipoSolucao.InformaMensagem objMensagem
    
        ' Inicia a transação
        frmgenMDI_SGAPlus.objBanco.Iniciar_Transacao
    
           ' Excutar o método de exclusão dos dados do usuário
            bRetorno = objCadastrarTipoSolucao.ExcluirDados(objCadastrarTipoSolucao.ID_TIPO_SOLUCAO)
    
            ' Se ocorreu algum erro na DLL
            If Not bRetorno Then
    
                frmgenMDI_SGAPlus.objBanco.Cancelar_Transacao
    
                GoTo Finalizar
    
            End If
    
        ' Finaliza a transação
        frmgenMDI_SGAPlus.objBanco.Finalizar_Transacao
        
        ' Se existia somente um registro limpar o objeto de entrada para buscar
        If lQuantidadeRegistros = 1 Then
        
            ' Limpa o campo de busca
            objFormCadastrarTipoSolucao.txtBusca.Text = ""
            
        End If
        
        ' Chama o método de exibição do GRID
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
                            "Houve um erro de execução no método: ExcluirDados " & _
                            "da classe clsgenOC_CadastrarTipoSolucao"

    Exit Sub

End Sub

Private Sub AtribuirCamposNasPropriedades()

    objCadastrarTipoSolucao.Descricao = TrataPlick(objFormCadastrarTipoSolucao.TxtDescricao)
    
End Sub

Private Sub Class_Terminate()

    Set objCadastrarTipoSolucao = Nothing
    Set objMensagem = Nothing
    Set objListaTipoSolucao = Nothing
    
End Sub

Private Function CriticarCamposTelaTipoSolucao() As Boolean

On Error GoTo ErrorHandler:

    CriticarCamposTelaTipoSolucao = False
    
'    If Len(objFormCadastrarTipoSolucao.txtCodigo_Controle.Text) = 0 Then
''        objMensagem.TipoMsg = "A"
''        objMensagem.Descricao = "Código de Controle é obrigatória."
''        objFormCadastrarTipoSolucao.txtCodigo_Controle.SetFocus
''
'        GoTo Finalizar
'    Else
    
    If Len(objFormCadastrarTipoSolucao.TxtDescricao.Text) = 0 Then
        objMensagem.TipoMsg = "A"
        objMensagem.Descricao = "Descrição é obrigatória."
        objFormCadastrarTipoSolucao.TxtDescricao.SetFocus
        
        GoTo Finalizar
    End If
    
    CriticarCamposTelaTipoSolucao = True

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
                            "Houve um erro de execução no método: CriticarCamposTelaTipoSolucao " & _
                            "da classe clsgenOC_CadastrarTipoSolucao"
    
    GoTo Finalizar

End Function

Public Function SaiTela(ByVal pyTela As Byte) As Boolean

On Error GoTo ErrorHandler:

    Dim bAlterou As Boolean

    SaiTela = False

    bAlterou = AlterouDadosTelaTipoSolucao
    
    ' Chama o método para verificar se foram alterados alguns dados
    If bAlterou Then
       
       objMensagem.TipoMsg = "Q"
       objMensagem.Descricao = "Foram feitas alterações nos dados da tela que " & _
                               "ainda não foram Salvas! Deseja sair assim mesmo ?"
       
       
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
                            "Houve um erro de execução no método: SaiTela " & _
                            "da classe clsgenOC_CadastrarTipoSolucaos"
    
    GoTo Finalizar

End Function

Public Function AlterouDadosTelaTipoSolucao() As Boolean

On Error GoTo ErrorHandler:

    AlterouDadosTelaTipoSolucao = True
    
    Set objMensagem = New clsgenMensagem
    
    If objCadastrarTipoSolucao.Descricao <> objFormCadastrarTipoSolucao.TxtDescricao Then
    
        GoTo Finalizar
        
    End If

'    If objCadastrarTipoSolucao.Codigo_Controle <> objFormCadastrarTipoSolucao.txtCodigo_Controle Then
'
'        GoTo Finalizar
'
'    End If
    AlterouDadosTelaTipoSolucao = False
    
Finalizar:

    On Error GoTo 0

    Exit Function
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: AlterouDadosTelaTipoSolucao " & _
                            "da classe clsgenOC_CadastrarTipoSolucaos"
    
    GoTo Finalizar

End Function

Private Function PreencheLista() As Boolean

    lQuantidadeRegistros = 0
    
    PreencheLista = False
    
    ' Ler o objeto de coleção da lista de tipos de pendências
    For Each objListaTipoSolucao In objCadastrarTipoSolucao.ListaTipoSolucao

        With objFormCadastrarTipoSolucao

            ' Descricao Motivo Recusa
             .lstTipoSolucao.AddItem (objListaTipoSolucao.Descricao)
        
        End With

        lQuantidadeRegistros = lQuantidadeRegistros + 1
    
    Next
    
    PreencheLista = True
    
End Function

Public Sub RefreshDados()

    If lQuantidadeRegistros > 0 Then
        ' Atualiza os dados de usuário na tela
        With objFormCadastrarTipoSolucao
            
            .TxtDescricao = objCadastrarTipoSolucao.ListaTipoSolucao(objFormCadastrarTipoSolucao.lstTipoSolucao.ListIndex + 1).Descricao
            '.txtCodigo_Controle = objCadastrarTipoSolucao.ListaTipoSolucao(objFormCadastrarTipoSolucao.lstTipoSolucao.ListIndex + 1).Codigo_Controle
            
            objCadastrarTipoSolucao.ID_TIPO_SOLUCAO = objCadastrarTipoSolucao.ListaTipoSolucao(objFormCadastrarTipoSolucao.lstTipoSolucao.ListIndex + 1).ID_TIPO_SOLUCAO
            objCadastrarTipoSolucao.Descricao = objCadastrarTipoSolucao.ListaTipoSolucao(objFormCadastrarTipoSolucao.lstTipoSolucao.ListIndex + 1).Descricao
            
            'objCadastrarTipoSolucao.Codigo_Controle = objCadastrarTipoSolucao.ListaTipoSolucao(objFormCadastrarTipoSolucao.lstTipoSolucao.ListIndex + 1).Codigo_Controle
            
        End With
        End If


End Sub

Private Sub HabilitaDesabilitaObjetos()

    ' Coloca cor de fundo como default para habilitação
    If (sTipoOperacao = "I") Or (sTipoOperacao = "A") Then ' Inclusão ou alteração
        
        ' Muda cor de fundo dos objetos de entrada de dados para branco
        objFormCadastrarTipoSolucao.TxtDescricao.BackColor = &H80000005
        'objFormCadastrarTipoSolucao.txtCodigo_Controle.BackColor = &H80000005
        
        ' Muda cor de escrita dos objetos de entrada de dados para preto
        objFormCadastrarTipoSolucao.TxtDescricao.ForeColor = &H80000008
        'objFormCadastrarTipoSolucao.txtCodigo_Controle.ForeColor = &H80000008
        
        ' Habilita o frame dos objetos de dados
        objFormCadastrarTipoSolucao.frameDados.Enabled = True
        
        ' Desabilita botões de inclusão, alteração, exlusão e frames de busca e lista
        objFormCadastrarTipoSolucao.cmdIncluir.Enabled = False
        objFormCadastrarTipoSolucao.cmdAlterar.Enabled = False
        objFormCadastrarTipoSolucao.cmdExcluir.Enabled = False
        objFormCadastrarTipoSolucao.frameBusca.Enabled = False
        objFormCadastrarTipoSolucao.frameLista.Enabled = False
        
        ' Habilita os botões gravar, limpar e cancelar
        objFormCadastrarTipoSolucao.cmdSalvar.Enabled = True
        objFormCadastrarTipoSolucao.cmdLimpar.Enabled = True
        objFormCadastrarTipoSolucao.cmdCancelar.Enabled = True
        
    Else
    
        ' Muda cor de fundo dos objetos de entrada de dados para cinza
        objFormCadastrarTipoSolucao.TxtDescricao.BackColor = &H80000004
        'objFormCadastrarTipoSolucao.txtCodigo_Controle.BackColor = &H80000004
        
        ' Muda cor de escrita dos objetos de entrada de dados azul
        objFormCadastrarTipoSolucao.TxtDescricao.ForeColor = &HC00000
        'objFormCadastrarTipoSolucao.txtCodigo_Controle.ForeColor = &HC00000
        
        ' Desabilita o frame dos objetos de dados
        objFormCadastrarTipoSolucao.frameDados.Enabled = False
    
        ' Habilita botões de inclusão, alteração, exlusão e frames de busca e lista
        objFormCadastrarTipoSolucao.cmdIncluir.Enabled = True
        objFormCadastrarTipoSolucao.cmdAlterar.Enabled = True
        objFormCadastrarTipoSolucao.cmdExcluir.Enabled = True
        objFormCadastrarTipoSolucao.frameBusca.Enabled = True
        objFormCadastrarTipoSolucao.frameLista.Enabled = True
        
        ' Desabilita os botões gravar, limpar e cancelar
        objFormCadastrarTipoSolucao.cmdSalvar.Enabled = False
        objFormCadastrarTipoSolucao.cmdLimpar.Enabled = False
        objFormCadastrarTipoSolucao.cmdCancelar.Enabled = False
    
    End If

End Sub

Public Sub LimparDadosTelaTipoSolucao()

    objFormCadastrarTipoSolucao.TxtDescricao.Text = ""
    'objFormCadastrarTipoSolucao.txtCodigo_Controle.Text = ""
    
    objCadastrarTipoSolucao.Descricao = ""
    'objCadastrarTipoSolucao.Codigo_Controle = ""
    
    Exit Sub
        
End Sub

Public Sub Saida(ByVal pyTela As Byte)
    
    Unload objFormCadastrarTipoSolucao
    
End Sub

'Public Sub Saida_PendProdFase()
'
'    Unload objFormCadPend_Prod_Fase
'
'End Sub
'

'Public Sub CadastrarPendenciaProdFase()
'
'    On Error GoTo ErrorHandler:
'
'    ' Instância o formulário de pré ativação e chama em seguida
'    'Dim objCadastrarPendProdFase As clsgenOC_CadPendenciaProdFase
'    'Set objCadastrarPendProdFase = New clsgenOC_CadPendenciaProdFase
'
'    ' Intância a classe de mensagem
'    Set objMensagem = New clsgenMensagem
'
'    TelaCadastrarPend_Prod_Fase objCadastrarTipoSolucao.Descricao, objCadastrarTipoSolucao.ID_TIPO_SOLUCAO
'
'
'
'Finalizar:
'
'    Screen.MousePointer = vbDefault
'
'    If objMensagem.TipoMsg <> "" Then
'
'        objMensagem.ExibeMensagem
'
'    End If
'
'    Set objMensagem = Nothing
'
'    On Error GoTo 0
'
'    Exit Sub
'
'ErrorHandler:
'
'    objMensagem.TipoMsg = "E"
'    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
'                            "Houve um erro de execução no método:  CadastrarPendenciaPRodFase" & _
'                            "da classe clsgenOC_CadastrarTipoSolucao"
'    GoTo Finalizar
'
'End Sub
'
'
'Public Function GravarDados_PendProdFase() As Boolean
'
'On Error GoTo ErrorHandler:
'
'    Dim bRetorno As Boolean
'    Dim iRow As Integer
'    Dim iCol As Integer
'
'    GravarDados_PendProdFase = False
'
'    ' Instância a classe de dados para cadastrar usuários
'    Set objCadPend_Prod_Fase = New clsgenPend_Prod_Fases
'
'    ' Informar o nome do banco desejado
'    objCadPend_Prod_Fase.InformaBanco frmgenMDI_SGAPlus.objBanco
'
'    ' Intância a classe de mensagem
'    Set objMensagem = New clsgenMensagem
'
'    ' Informar o nome do objeto de mensagem
'    objCadPend_Prod_Fase.InformaMensagem objMensagem
'
'    ' Inicia a transação
'    bRetorno = frmgenMDI_SGAPlus.objBanco.Iniciar_Transacao()
'
'    ' Se ocorreu algum erro na DLL
'    If Not bRetorno Then
'
'        GoTo Finalizar
'
'    End If
'
'    ' Excutar o método para gravar os dados
'    'Verifica array para saber se info será excluida
'    iCol = 1
'    sTipoOperacao = ""
'
'    For iRow = 1 To UBound(aListaTipSolucao, 2)
'            'Verifica se ainda tem info no array
'            If IsEmpty(aListaTipSolucao(iCol, iRow)) Then Exit For
'
'            If IsNull(aListaTipSolucao(iCol, iRow)) Then
'                'Marcado para exclusão
'                sTipoOperacao = "E"
'
'            ElseIf aListaTipSolucao(iCol, iRow) = "NOVO" Then
'                'Novo registro
'                sTipoOperacao = "I"
'            End If
'
'            If sTipoOperacao <> "" Then
'                'Pega Codigos de Pendencia, Produto e Fase
'                objCadPend_Prod_Fase.TIPO_PENDENCIA_ID = aListaTipSolucao(0, iRow)
'                objCadPend_Prod_Fase.Produto_ID = aListaTipSolucao(2, iRow)
'                objCadPend_Prod_Fase.Codigo_Tipo_Fase = aListaTipSolucao(4, iRow)
'
'                bRetorno = objCadPend_Prod_Fase.GravarDados(sTipoOperacao)
'
'                ' Se ocorreu algum erro na DLL
'                If Not bRetorno Then
'
'                    GoTo Finalizar
'
'                End If
'                sTipoOperacao = ""
'            End If
'    Next
'
'    ' Finaliza a transação
'    frmgenMDI_SGAPlus.objBanco.Finalizar_Transacao
'
'    ' Atualiza a mensagem de retorno como gravação realizada com sucesso
'    objMensagem.TipoMsg = "A"
'    objMensagem.Descricao = "Gravação realizada com sucesso"
'
'    GravarDados_PendProdFase = True
'
'Finalizar:
'
'    Screen.MousePointer = vbDefault
'
'    If objMensagem.TipoMsg <> "" Then
'
'        objMensagem.ExibeMensagem
'
'    End If
'
'    If Not GravarDados_PendProdFase Then
'
'        frmgenMDI_SGAPlus.objBanco.Cancelar_Transacao
'
'    End If
'
'    Set objMensagem = Nothing
'
'    On Error GoTo 0
'
'    Exit Function
'
'ErrorHandler:
'
'    objMensagem.TipoMsg = "E"
'    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
'                            "Houve um erro de execução no método: GravarDados " & _
'                            "da classe clsgenOC_CadastrarTipoSolucao"
'
'    GoTo Finalizar
'
'End Function
'
'
''Public Sub TelaCadastrarPend_Prod_Fase(ByVal sTexto As String, ByVal lID As Long)
'''Parametros sTexto e lID serão descricao e ID passado pelas telas do Usuario ou da CF
''
''    Dim bRetorno As Boolean
''
''On Error GoTo ErrorHandler:
''
''    ' Instância o formulário de pré ativação e chama em seguida
''    Set objFormCadPend_Prod_Fase = New frmgenCadPend_Prod_Fase
''
''    ' Intância a classe de mensagem
''    Set objMensagem = New clsgenMensagem
''
''    ' Informar o objeto controlador
''    objFormCadPend_Prod_Fase.InformaClasseOC Me
''
''    ' Instância a classe de dados para cadastrar usuários
''    Set objCadPend_Prod_Fase = New clsgenPend_Prod_Fases
''
''    ' Informar o nome do banco desejado
''    objCadPend_Prod_Fase.InformaBanco frmgenMDI_SGAPlus.objBanco
''
''    ' Informar o nome do objeto de mensagem
''    objCadPend_Prod_Fase.InformaMensagem objMensagem
''
''    ' Chama o método de exibição do GRID
''    objFormCadPend_Prod_Fase.lblPendencia.Caption = sTexto
''
''
''    'Monta Combo Produto e Fase
''    Call ExibirListaProduto(objFormCadPend_Prod_Fase.cboProduto)
''
''    'Monta inicialmente lista de Fases baseada no Produto BLF e Acao=Ativação
''    Call ExibirListaFase(GbIDProdutoBLF, GbIDAcaoAtivacao, objFormCadPend_Prod_Fase.cboFase)
''
''    objFormCadPend_Prod_Fase.grdAlocacoes.Col = 1
''    objFormCadPend_Prod_Fase.grdAlocacoes.ColHidden = False
''    objFormCadPend_Prod_Fase.grdAlocacoes.Col = 2
''    objFormCadPend_Prod_Fase.grdAlocacoes.ColHidden = False
''
''    'Atribui valores parametros nas propriedades para tela Usu
''    objFormCadPend_Prod_Fase.lblPendencia = sTexto
''    objFormCadPend_Prod_Fase.Id_Pendencia = lID
''
''
''    'Seleciona CFs do Usuario corrente
''    bRetorno = ExibirListaGrid(lID)
''
''    objFormCadPend_Prod_Fase.grdAlocacoes.Col = 1
''
''    objFormCadPend_Prod_Fase.grdAlocacoes.Row = objFormCadPend_Prod_Fase.grdAlocacoes.ActiveRow
''
''    If objFormCadPend_Prod_Fase.grdAlocacoes.Text <> "" Then
''        objFormCadPend_Prod_Fase.cboProduto = objFormCadPend_Prod_Fase.grdAlocacoes.Text
''    End If
''
''    objFormCadPend_Prod_Fase.grdAlocacoes.Col = 2
''
''    If objFormCadPend_Prod_Fase.grdAlocacoes.Text <> "" Then
''        'objFormCadPend_Prod_Fase.cboFase = objFormCadPend_Prod_Fase.grdAlocacoes.Text
''    End If
''
''
''    ' Se ocorreu algum erro na DLL
''    If Not bRetorno Then
''
''        GoTo Finalizar
''
''    End If
''
''
''Finalizar:
''
''    Screen.MousePointer = vbDefault
''
''    If objMensagem.TipoMsg <> "" Then
''
''        objMensagem.ExibeMensagem
''
''    End If
''
''    On Error GoTo 0
''
''    Exit Sub
''
''ErrorHandler:
''
''    objMensagem.TipoMsg = "E"
''    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
''                            "Houve um erro de execução no método: TelaCadastrarPend_Prod_Fase " & _
''                            "da classe clsgenOC_CadastrarTipoSolucao"
''    GoTo Finalizar
''
''End Sub
''
''
''Private Sub PreencheGrid_PendProdFase()
''
''    Dim iRow As Integer
''
''    'Limpa o Grid
''    objFormCadPend_Prod_Fase.grdAlocacoes.MaxRows = 1
''
''    objFormCadPend_Prod_Fase.grdAlocacoes.Col = 1
''    objFormCadPend_Prod_Fase.grdAlocacoes.Row = 1
''    objFormCadPend_Prod_Fase.grdAlocacoes.Text = ""
''
''    objFormCadPend_Prod_Fase.grdAlocacoes.Col = 2
''    objFormCadPend_Prod_Fase.grdAlocacoes.Row = 1
''    objFormCadPend_Prod_Fase.grdAlocacoes.Text = ""
''
''    objFormCadPend_Prod_Fase.grdAlocacoes.Col = 3
''    objFormCadPend_Prod_Fase.grdAlocacoes.Row = 1
''    objFormCadPend_Prod_Fase.grdAlocacoes.Text = ""
''
''
''    ' Ler o array para preencher o GRID
''
''
''   For iRow = objFormCadPend_Prod_Fase.grdAlocacoes.DataRowCnt + 1 To UBound(aListaTipSolucao, 2)
''
''
''        If Not IsEmpty(aListaTipSolucao(0, iRow)) Then
''            'Produto
''            objFormCadPend_Prod_Fase.grdAlocacoes.Col = 1
''            objFormCadPend_Prod_Fase.grdAlocacoes.Row = iRow
''            objFormCadPend_Prod_Fase.grdAlocacoes.Text = aListaTipSolucao(3, iRow)
''
''            'Fases
''            objFormCadPend_Prod_Fase.grdAlocacoes.Col = 2
''            objFormCadPend_Prod_Fase.grdAlocacoes.Text = aListaTipSolucao(5, iRow)
''
''            'Flag para Identificar se é um Registro Novo
''            objFormCadPend_Prod_Fase.grdAlocacoes.Col = 3
''            objFormCadPend_Prod_Fase.grdAlocacoes.Text = "N"
''
''            objFormCadPend_Prod_Fase.grdAlocacoes.MaxRows = iRow + 1
''        Else
''            objFormCadPend_Prod_Fase.grdAlocacoes.MaxRows = objFormCadPend_Prod_Fase.grdAlocacoes.DataRowCnt
''            Exit For
''
''        End If
''    Next
''
''End Sub
''
''
''Public Sub ExcluirDoGrid_PendProdFase()
''
''    Dim iContador As Integer
''    Dim sCampo1, sCampo2 As String
''    Dim iLinha As Integer
''
''    'Verifica se grid esta vazio
''    If objFormCadPend_Prod_Fase.grdAlocacoes.MaxRows = 0 Then
''
''        Exit Sub
''
''    End If
''
''    'Seta variavel numero de linhas do grid
''    'Pode haver linhas escondidas, que marcam os regs para exclusão
''    iLinha = objFormCadPend_Prod_Fase.grdAlocacoes.MaxRows
''
''    'no array marca com Excluido os registros a serem excluidos
''    'ID Pendencia
''    'aListaTipSolucao(0, objFormCadPend_Prod_Fase.grdAlocacoes.ActiveRow) = Null
''    'Pendencia - Descricao
''    aListaTipSolucao(1, objFormCadPend_Prod_Fase.grdAlocacoes.ActiveRow) = Null
''
''    objFormCadPend_Prod_Fase.grdAlocacoes.Row = objFormCadPend_Prod_Fase.grdAlocacoes.ActiveRow
''    objFormCadPend_Prod_Fase.grdAlocacoes.RowHidden = True
''
''    'Atualiza texto combo
''    objFormCadPend_Prod_Fase.grdAlocacoes.Row = objFormCadPend_Prod_Fase.grdAlocacoes.ActiveRow - 1
''
''End Sub
''
''
''
''Public Function CriticarCamposTela_PendProdFase() As Boolean
''
''On Error GoTo ErrorHandler:
''
''    CriticarCamposTela_PendProdFase = False
''
''    If Len(objFormCadPend_Prod_Fase.cboFase.Text) = "" Then
''        objMensagem.TipoMsg = "A"
''        objMensagem.Descricao = "Selecione a Fase informação obrigatória."
''        objFormCadPend_Prod_Fase.cboFase.SetFocus
''
''        GoTo Finalizar
''    End If
''
''    If Len(objFormCadPend_Prod_Fase.cboProduto.Text) = "" Then
''        objMensagem.TipoMsg = "A"
''        objMensagem.Descricao = "Selecione o Produto, informação obrigatória."
''        objFormCadPend_Prod_Fase.cboProduto.SetFocus
''
''        GoTo Finalizar
''    End If
''
''    CriticarCamposTela_PendProdFase = True
''
''Finalizar:
''
''    If objMensagem.TipoMsg <> "" Then
''
''        Screen.MousePointer = vbDefault
''
''        objMensagem.ExibeMensagem
''
''    End If
''
''    On Error GoTo 0
''
''    Exit Function
''
''ErrorHandler:
''
''    objMensagem.TipoMsg = "E"
''    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
''                            "Houve um erro de execução no método: CriticarCamposTela " & _
''                            "da classe clsgenOC_CadastrarPendencia_Produto_Fase"
''
''    GoTo Finalizar
''
''End Function
''
''
''Public Sub IncluirNoGrid_PendProdFase(ByVal sTextoCbo As String, ByVal sTextoCbo2 As String, _
''                         ByVal sTipoSolucao As String, _
''                         ByVal iLinha As Integer, _
''                         ByVal piOrigem As Integer)
'
'
'    Dim LcRow  As Integer
'    Dim cColecao As New Collection
'    Dim lRetorno As Long
'    Dim lmax As Long
'    Dim iColuna As Integer
'
'    ' Intância a classe de mensagem
'    Set objMensagem = New clsgenMensagem
'
'    ' Intância a classe de mensagem
'    Set objCadPend_Prod_Fase = New clsgenPend_Prod_Fases
'
'    ' Informar o nome do banco desejado
'    objCadPend_Prod_Fase.InformaBanco frmgenMDI_SGAPlus.objBanco
'    objCadPend_Prod_Fase.InformaMensagem objMensagem
'
'    'Verifca se grid esta vazio
'    If objFormCadPend_Prod_Fase.grdAlocacoes.DataRowCnt <> 0 Then
'
'        'Verifica se item do combo já existe no grid
'        For LcRow = 1 To iLinha
'
'            objFormCadPend_Prod_Fase.grdAlocacoes.Col = 1
'            objFormCadPend_Prod_Fase.grdAlocacoes.Row = LcRow
'
'            If objFormCadPend_Prod_Fase.grdAlocacoes.Text = sTextoCbo Then
'
'                objFormCadPend_Prod_Fase.grdAlocacoes.Col = 2
'                objFormCadPend_Prod_Fase.grdAlocacoes.Row = LcRow
'
'                If objFormCadPend_Prod_Fase.grdAlocacoes.Text = sTextoCbo2 Then
'
'                    If objFormCadPend_Prod_Fase.grdAlocacoes.RowHidden = False Then
'                        ' Atualiza a mensagem de retorno
'                        objMensagem.TipoMsg = "A"
'                        objMensagem.Descricao = "Informação já cadastrada."
'
'                        GoTo Finalizar
'                    Else
'                         ''Verifica se o registro referente a colona Oculta ja se encontra na Base
'                         '"S" igual registro novo ( não esta na Base )
'                         objFormCadPend_Prod_Fase.grdAlocacoes.Col = 3
'                         If objFormCadPend_Prod_Fase.grdAlocacoes.Text = "N" Then
'                            aListaTipSolucao(1, objFormCadPend_Prod_Fase.grdAlocacoes.Row) = sTipoSolucao
'                         Else
'                            aListaTipSolucao(1, objFormCadPend_Prod_Fase.grdAlocacoes.Row) = "NOVO"
'                         End If
'                        objFormCadPend_Prod_Fase.grdAlocacoes.RowHidden = False
'                        GoTo Finalizar
'                    End If
'                End If
'            End If
'        Next
'        objFormCadPend_Prod_Fase.grdAlocacoes.MaxRows = iLinha + 1
'    Else
'        'Não ha dados no grid
'        lmax = 1000
'        ReDim aListaTipSolucao(7, lmax)
'        objFormCadPend_Prod_Fase.grdAlocacoes.MaxRows = iLinha
'    End If
'
'    'Inclui info no grid e no array
'    objFormCadPend_Prod_Fase.grdAlocacoes.Row = objFormCadPend_Prod_Fase.grdAlocacoes.MaxRows
'
'
'
'    ''Obter ID's do produto e Fase
'    lRetorno = objCadPend_Prod_Fase.ObtemID(sTipoSolucao)
'    lID_TIPO_SOLUCAO = lRetorno
'
'    'Coluna 1 Produto
'     objFormCadPend_Prod_Fase.grdAlocacoes.Col = 1
'     objFormCadPend_Prod_Fase.grdAlocacoes.Text = objFormCadPend_Prod_Fase.cboProduto
'
'    'Coluna 2 Fase
'     objFormCadPend_Prod_Fase.grdAlocacoes.Col = 2
'     objFormCadPend_Prod_Fase.grdAlocacoes.Text = objFormCadPend_Prod_Fase.cboFase
'
'    'Coluna 3 Flag q indica se é um novo Registro
'     objFormCadPend_Prod_Fase.grdAlocacoes.Col = 3
'     objFormCadPend_Prod_Fase.grdAlocacoes.Text = "S"
'    '
'
'
'    ' Atualiza o array
'    'Pendencia_ID
'     aListaTipSolucao(0, objFormCadPend_Prod_Fase.grdAlocacoes.Row) = lID_TIPO_SOLUCAO
'
'    'Inclui Descricao referente a Pendencia -
'    'Esta será utilizada como Flag para indicar que foi excluido (para Excluido valor = Null)
'     aListaTipSolucao(1, objFormCadPend_Prod_Fase.grdAlocacoes.Row) = "NOVO"
'
'    'ID Produto
'     aListaTipSolucao(2, objFormCadPend_Prod_Fase.grdAlocacoes.Row) = objFormCadPend_Prod_Fase.cboProduto.ItemData(objFormCadPend_Prod_Fase.cboProduto.ListIndex)
'
'    'Inclui Descricao do Produto referente ao CF selecionado no combo
'     aListaTipSolucao(3, objFormCadPend_Prod_Fase.grdAlocacoes.Row) = objFormCadPend_Prod_Fase.cboProduto
'
'    'ID Fase
'     aListaTipSolucao(4, objFormCadPend_Prod_Fase.grdAlocacoes.Row) = objFormCadPend_Prod_Fase.cboFase.ItemData(objFormCadPend_Prod_Fase.cboFase.ListIndex)
'
'    'Inclui Descricao da Fase  selecionada no combo
'     aListaTipSolucao(5, objFormCadPend_Prod_Fase.grdAlocacoes.Row) = objFormCadPend_Prod_Fase.cboFase
'
'
'Finalizar:
'
'    Screen.MousePointer = vbDefault
'
'    If objMensagem.TipoMsg <> "" Then
'
'        objMensagem.ExibeMensagem
'
'    End If
'
'End Sub
'
'
Public Function ExibirListaProduto(ByRef cbo As ComboBox) As Boolean

On Error GoTo ErrorHandler:

    Dim objProduto As clsGenProduto
    Dim objMensagem As clsgenMensagem
    Dim bRetorno As Boolean
    Dim aListaProduto() As Variant
    
    ExibirListaProduto = False
    
    
    ' Limpa a combo
    cbo.Clear
    
    ' Intância a classe de mensagem
    Set objMensagem = New clsgenMensagem
    
    ' Instância a classe de centro funcional
    Set objProduto = New clsGenProduto

    ' Informar o nome do banco desejado
    objProduto.InformaBanco frmgenMDI_SGAPlus.objBanco
                
    ' Informar o nome do objeto de mensagem
    objProduto.InformaMensagem objMensagem
    
    ' Executar o método de montagem da lista de centro funcional na classe de analisar order entry
    bRetorno = objProduto.MontarListaProduto(aListaProduto, "I")
    
    ' Se ocorreu algum erro na DLL
    If Not bRetorno Then
        
        Exit Function
    
    End If
    
    If objMensagem.TipoMsg = "" Then
    
        PreencheCombo cbo, aListaProduto
        
    End If
    
    ExibirListaProduto = True
    
Finalizar:

    Screen.MousePointer = vbDefault
    
    If objMensagem.TipoMsg <> "" Then
    
        objMensagem.ExibeMensagem
        
    End If
    
    Set objProduto = Nothing
    Set objMensagem = Nothing
    
    Exit Function
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: ExibirListaCentroFuncional " & _
                            "da classe GenOC_CadastrarPendencia_Produto_Fase"
    
    GoTo Finalizar

End Function


Public Function ExibirListaFase(ByVal vProduto, ByVal vAcao, ByRef cbo As ComboBox) As Boolean

On Error GoTo ErrorHandler:

    Dim objMensagem As clsgenMensagem
    Dim bRetorno As Boolean
    Dim aListaFase() As Variant
    Dim objFase As New clsgenFase
    
    
    ExibirListaFase = False
    
    ' Limpa a combo
    cbo.Clear
    
    ' Intância a classe de mensagem
    Set objMensagem = New clsgenMensagem
    
    ' Informar o nome do banco desejado
    objFase.InformaBanco frmgenMDI_SGAPlus.objBanco
                
    ' Informar o nome do objeto de mensagem
    objFase.InformaMensagem objMensagem
    
    ' Executar o método de montagem da lista de centro funcional na classe de analisar order entry
    bRetorno = objFase.MontarListaFase(vProduto, vAcao, aListaFase, "S")
    
    
    ' Se ocorreu algum erro na DLL
    If Not bRetorno Then
        
        Exit Function
    
    End If
    
    If objMensagem.TipoMsg = "" Then
    
        PreencheCombo cbo, aListaFase
        
    End If
    
    ExibirListaFase = True
    
Finalizar:

    Screen.MousePointer = vbDefault
    
    If objMensagem.TipoMsg <> "" Then
    
        objMensagem.ExibeMensagem
        
    End If
    
    Set objMensagem = Nothing
    Set objFase = Nothing
    
    Exit Function
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: ExibirListaCentroFuncional " & _
                            "da classe clsgenOC_CadastrarPendencia_Produto_Fase"
    
    GoTo Finalizar

End Function


'Public Function ExibirListaGrid(ByVal lID_TIPO_SOLUCAO As Long) As Boolean
'
'On Error GoTo ErrorHandler:
'
'    Dim objMensagem As clsgenMensagem
'    Dim bRetorno As Boolean
'    Dim lID_Produto As Long
'
'    ExibirListaGrid = False
'
'    ' Intância a classe de mensagem
'    Set objMensagem = New clsgenMensagem
'    Set objCadPend_Prod_Fase = New clsgenPend_Prod_Fases
'
'    ' Informar o nome do banco desejado
'    objCadPend_Prod_Fase.InformaBanco frmgenMDI_SGAPlus.objBanco
'
'    ' Informar o nome do objeto de mensagem
'    objCadPend_Prod_Fase.InformaMensagem objMensagem
'
'    'Verifica se o PRoduto foi selecionado
'    'Caso não seja, atribui o Produto BLFlex como padrão
'    If objFormCadPend_Prod_Fase.cboProduto.ListIndex >= 0 Then
'        lID_Produto = objFormCadPend_Prod_Fase.cboProduto.ItemData(objFormCadPend_Prod_Fase.cboProduto.ListIndex)
'    Else
'        lID_Produto = GbIDProdutoBLF
'    End If
'    ' Executar o método de montagem da lista de centro funcional relacionados ao usuário
'    bRetorno = objCadPend_Prod_Fase.MontaListaGrid(lID_TIPO_SOLUCAO, aListaTipSolucao, lID_Produto)
'
'   ' Se ocorreu algum erro na DLL exibi-lo
'    If Not bRetorno Then
'
'        GoTo Finalizar
'
'    End If
'
'    ' Preenche a lista de usuários
'    Call PreencheGrid_PendProdFase
'
'    ExibirListaGrid = True
'
'Finalizar:
'
'    Screen.MousePointer = vbDefault
'
'    If objMensagem.TipoMsg <> "" Then
'
'        objMensagem.ExibeMensagem
'
'    End If
'
'    Set objCadPend_Prod_Fase = Nothing
'    Set objMensagem = Nothing
'
'    Exit Function
'
'ErrorHandler:
'
'    objMensagem.TipoMsg = "E"
'    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
'                            "Houve um erro de execução no método: ExibirListaGrid " & _
'                            "da classe clsGenOC_CadastrarPendenciaProdutoFase"
'
'    GoTo Finalizar
'
'End Function
'
'
