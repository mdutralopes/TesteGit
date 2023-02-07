VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsgenOC_CadastrarTipoPendencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Declara as classes utilizadas
Private objMensagem                     As clsgenMensagem
Private objCadastrarTipoPendencia       As clsgenTipoPendencia
Private objListaTipoPendencia           As clsgenListaTipoPendencia
Private objCadPend_Prod_Fase            As clsgenPend_Prod_Fases
'Private lID_TIPO_PENDENCIA

' Declara os formulários usados nesta classe
Private objFormCadastrarTipoPendencia   As frmgenCadastroTipoPendencias
Private objFormCadPend_Prod_Fase        As frmgenCadPend_Prod_Fase
' Declara variáveis comuns
Private lID_TIPO_PENDENCIA              As Long
Private sTipoOperacao                   As String
Private lQuantidadeRegistros            As Long
Private sPendencia                      As String
Private aListaPend()                      As Variant

Public Property Get Pendencia() As String

    Pendencia = sPendencia
    
End Property
Public Property Let Pendencia(ByVal sNewValue As String)

    sPendencia = sNewValue
    
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

Public Sub TelaCadastrarTipoPendencias()

On Error GoTo ErrorHandler:

    ' Instância o formulário de pré ativação e chama em seguida
    Set objFormCadastrarTipoPendencia = New frmgenCadastroTipoPendencias
    
    ' Intância a classe de mensagem
    Set objMensagem = New clsgenMensagem
    
    ' Informar o objeto controlador
    objFormCadastrarTipoPendencia.InformaClasseOC Me
    
    ' Instância a classe de dados para cadastrar usuários
    Set objCadastrarTipoPendencia = New clsgenTipoPendencia
    
    ' Informar o nome do banco desejado
    objCadastrarTipoPendencia.InformaBanco frmgenMDI_SGAPlus.objBanco
            
    ' Informar o nome do objeto de mensagem
    objCadastrarTipoPendencia.InformaMensagem objMensagem
    
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
                            "Houve um erro de execução no método: TelaCadastrarTipoPendencias " & _
                            "da classe clsgenOC_CadastrarTipoPendencia"
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
    objCadastrarTipoPendencia.InformaMensagem objMensagem
    
    ' Intância a classe da lista de usuários
    Set objListaTipoPendencia = Nothing
    Set objListaTipoPendencia = New clsgenListaTipoPendencia
    
    ' Limpa a lista
    objFormCadastrarTipoPendencia.lstTipoPendencias.Clear
    
      
    ' Executar o método de montagem da lista na classe de analisar order entry
    bRetorno = objCadastrarTipoPendencia.MontarLista(objFormCadastrarTipoPendencia.optCodigoControle.Value, _
                                                       objFormCadastrarTipoPendencia.txtBusca.Text)
    
    
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
    objFormCadastrarTipoPendencia.lstTipoPendencias.ListIndex = 0
    
Finalizar:

    ' Se não encontrou registro desabilitar os botôes de alteração e exclusão
    If lQuantidadeRegistros = 0 Then
    
        objFormCadastrarTipoPendencia.cmdAlterar.Enabled = False
        objFormCadastrarTipoPendencia.cmdExcluir.Enabled = False
        
        Call LimparDadosTelaTipoPendencia
        
    Else
    
        objFormCadastrarTipoPendencia.cmdAlterar.Enabled = True
        objFormCadastrarTipoPendencia.cmdExcluir.Enabled = True
        
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
                            "da classe clswidOC_CadastrarTipoPendencia"

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
        bRetorno = CriticarCamposTelaTipoPendencia
            
        If Not bRetorno Then
        
            GoTo Finalizar
            
        End If
    
        ' Informar o nome do objeto de mensagem
        objCadastrarTipoPendencia.InformaMensagem objMensagem
        
        ' Atribui os campos da tela nas propriedades
        Call AtribuirCamposNasPropriedades
            
        ' Excutar o método para gravar os dados do usuário
        bRetorno = objCadastrarTipoPendencia.GravarDados(sTipoOperacao)
            
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
                            "da classe clsgenOC_CadastrarTipoPendencia"

    GoTo Finalizar

End Function

Public Sub ExcluirDados()

On Error GoTo ErrorHandler:

    Dim bRetorno As Boolean

    ' Intância a classe de mensagem
    Set objMensagem = New clsgenMensagem
    
    objMensagem.TipoMsg = "Q"
    objMensagem.Descricao = "Confirma a exlusão do Tipo de Pendência selecionado?"
    If objMensagem.ExibeMensagem = "S" Then

        ' Informar o nome do objeto de mensagem
        objCadastrarTipoPendencia.InformaMensagem objMensagem
    
        ' Inicia a transação
        frmgenMDI_SGAPlus.objBanco.Iniciar_Transacao
    
           ' Excutar o método de exclusão dos dados do usuário
            bRetorno = objCadastrarTipoPendencia.ExcluirDados(objCadastrarTipoPendencia.ID_TIPO_PENDENCIA)
    
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
            objFormCadastrarTipoPendencia.txtBusca.Text = ""
            
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
                            "da classe clsgenOC_CadastrarTipoPendencia"

    Exit Sub

End Sub

Private Sub AtribuirCamposNasPropriedades()

    objCadastrarTipoPendencia.Descricao = TrataPlick(objFormCadastrarTipoPendencia.TxtDescricao)
    objCadastrarTipoPendencia.Codigo_Controle = TrataPlick(objFormCadastrarTipoPendencia.txtCodigo_Controle)
    
End Sub

Private Sub Class_Terminate()

    Set objCadastrarTipoPendencia = Nothing
    Set objMensagem = Nothing
    Set objListaTipoPendencia = Nothing
    
End Sub

Private Function CriticarCamposTelaTipoPendencia() As Boolean

On Error GoTo ErrorHandler:

    CriticarCamposTelaTipoPendencia = False
    
'    If Len(objFormCadastrarTipoPendencia.txtCodigo_Controle.Text) = 0 Then
''        objMensagem.TipoMsg = "A"
''        objMensagem.Descricao = "Código de Controle é obrigatória."
''        objFormCadastrarTipoPendencia.txtCodigo_Controle.SetFocus
''
'        GoTo Finalizar
'    Else
    
    If Len(objFormCadastrarTipoPendencia.TxtDescricao.Text) = 0 Then
        objMensagem.TipoMsg = "A"
        objMensagem.Descricao = "Descrição é obrigatória."
        objFormCadastrarTipoPendencia.TxtDescricao.SetFocus
        
        GoTo Finalizar
    End If
    
    CriticarCamposTelaTipoPendencia = True

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
                            "Houve um erro de execução no método: CriticarCamposTelaTipoPendencia " & _
                            "da classe clsgenOC_CadastrarTipoPendencia"
    
    GoTo Finalizar

End Function

Public Function SaiTela(ByVal pyTela As Byte) As Boolean

On Error GoTo ErrorHandler:

    Dim bAlterou As Boolean

    SaiTela = False

    bAlterou = AlterouDadosTelaTipoPendencia
    
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
                            "da classe clsgenOC_CadastrarTipoPendencias"
    
    GoTo Finalizar

End Function

Public Function AlterouDadosTelaTipoPendencia() As Boolean

On Error GoTo ErrorHandler:

    AlterouDadosTelaTipoPendencia = True
    
    Set objMensagem = New clsgenMensagem
    
    If objCadastrarTipoPendencia.Descricao <> objFormCadastrarTipoPendencia.TxtDescricao Then
    
        GoTo Finalizar
        
    End If

    If objCadastrarTipoPendencia.Codigo_Controle <> objFormCadastrarTipoPendencia.txtCodigo_Controle Then
    
        GoTo Finalizar
        
    End If
    AlterouDadosTelaTipoPendencia = False
    
Finalizar:

    On Error GoTo 0

    Exit Function
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: AlterouDadosTelaTipoPendencia " & _
                            "da classe clsgenOC_CadastrarTipoPendencias"
    
    GoTo Finalizar

End Function

Private Function PreencheLista() As Boolean

    lQuantidadeRegistros = 0
    
    PreencheLista = False
    
    ' Ler o objeto de coleção da lista de tipos de pendências
    For Each objListaTipoPendencia In objCadastrarTipoPendencia.ListaTipoPendencia

        With objFormCadastrarTipoPendencia

            ' Descricao Motivo Recusa
             .lstTipoPendencias.AddItem (objListaTipoPendencia.Descricao)
        
        End With

        lQuantidadeRegistros = lQuantidadeRegistros + 1
    
    Next
    
    PreencheLista = True
    
End Function

Public Sub RefreshDados()

    If lQuantidadeRegistros > 0 Then
    ' Atualiza os dados de usuário na tela
    With objFormCadastrarTipoPendencia
        
        .TxtDescricao = objCadastrarTipoPendencia.ListaTipoPendencia(objFormCadastrarTipoPendencia.lstTipoPendencias.ListIndex + 1).Descricao
        .txtCodigo_Controle = objCadastrarTipoPendencia.ListaTipoPendencia(objFormCadastrarTipoPendencia.lstTipoPendencias.ListIndex + 1).Codigo_Controle
        
        objCadastrarTipoPendencia.ID_TIPO_PENDENCIA = objCadastrarTipoPendencia.ListaTipoPendencia(objFormCadastrarTipoPendencia.lstTipoPendencias.ListIndex + 1).ID_TIPO_PENDENCIA
        objCadastrarTipoPendencia.Descricao = objCadastrarTipoPendencia.ListaTipoPendencia(objFormCadastrarTipoPendencia.lstTipoPendencias.ListIndex + 1).Descricao
        
        objCadastrarTipoPendencia.Codigo_Controle = objCadastrarTipoPendencia.ListaTipoPendencia(objFormCadastrarTipoPendencia.lstTipoPendencias.ListIndex + 1).Codigo_Controle
        
    End With
    End If


End Sub

Private Sub HabilitaDesabilitaObjetos()

    ' Coloca cor de fundo como default para habilitação
    If (sTipoOperacao = "I") Or (sTipoOperacao = "A") Then ' Inclusão ou alteração
        
        ' Muda cor de fundo dos objetos de entrada de dados para branco
        objFormCadastrarTipoPendencia.TxtDescricao.BackColor = &H80000005
        objFormCadastrarTipoPendencia.txtCodigo_Controle.BackColor = &H80000005
        
        ' Muda cor de escrita dos objetos de entrada de dados para preto
        objFormCadastrarTipoPendencia.TxtDescricao.ForeColor = &H80000008
        objFormCadastrarTipoPendencia.txtCodigo_Controle.ForeColor = &H80000008
        
        ' Habilita o frame dos objetos de dados
        objFormCadastrarTipoPendencia.frameDados.Enabled = True
        
        ' Desabilita botões de inclusão, alteração, exlusão e frames de busca e lista
        objFormCadastrarTipoPendencia.cmdIncluir.Enabled = False
        objFormCadastrarTipoPendencia.cmdAlterar.Enabled = False
        objFormCadastrarTipoPendencia.cmdExcluir.Enabled = False
        objFormCadastrarTipoPendencia.frameBusca.Enabled = False
        objFormCadastrarTipoPendencia.frameLista.Enabled = False
        
        ' Habilita os botões gravar, limpar e cancelar
        objFormCadastrarTipoPendencia.cmdSalvar.Enabled = True
        objFormCadastrarTipoPendencia.cmdLimpar.Enabled = True
        objFormCadastrarTipoPendencia.cmdCancelar.Enabled = True
        
    Else
    
        ' Muda cor de fundo dos objetos de entrada de dados para cinza
        objFormCadastrarTipoPendencia.TxtDescricao.BackColor = &H80000004
        objFormCadastrarTipoPendencia.txtCodigo_Controle.BackColor = &H80000004
        
        ' Muda cor de escrita dos objetos de entrada de dados azul
        objFormCadastrarTipoPendencia.TxtDescricao.ForeColor = &HC00000
        objFormCadastrarTipoPendencia.txtCodigo_Controle.ForeColor = &HC00000
        
        ' Desabilita o frame dos objetos de dados
        objFormCadastrarTipoPendencia.frameDados.Enabled = False
    
        ' Habilita botões de inclusão, alteração, exlusão e frames de busca e lista
        objFormCadastrarTipoPendencia.cmdIncluir.Enabled = True
        objFormCadastrarTipoPendencia.cmdAlterar.Enabled = True
        objFormCadastrarTipoPendencia.cmdExcluir.Enabled = True
        objFormCadastrarTipoPendencia.frameBusca.Enabled = True
        objFormCadastrarTipoPendencia.frameLista.Enabled = True
        
        ' Desabilita os botões gravar, limpar e cancelar
        objFormCadastrarTipoPendencia.cmdSalvar.Enabled = False
        objFormCadastrarTipoPendencia.cmdLimpar.Enabled = False
        objFormCadastrarTipoPendencia.cmdCancelar.Enabled = False
    
    End If

End Sub

Public Sub LimparDadosTelaTipoPendencia()

    objFormCadastrarTipoPendencia.TxtDescricao.Text = ""
    objFormCadastrarTipoPendencia.txtCodigo_Controle.Text = ""
    
    objCadastrarTipoPendencia.Descricao = ""
    objCadastrarTipoPendencia.Codigo_Controle = ""
    
    Exit Sub
        
End Sub

Public Sub Saida(ByVal pyTela As Byte)
    
    Unload objFormCadastrarTipoPendencia
    
End Sub

Public Sub Saida_PendProdFase()
    
    Unload objFormCadPend_Prod_Fase
    
End Sub


Public Sub CadastrarPendenciaProdFase()

    On Error GoTo ErrorHandler:

    ' Instância o formulário de pré ativação e chama em seguida
    'Dim objCadastrarPendProdFase As clsgenOC_CadPendenciaProdFase
    'Set objCadastrarPendProdFase = New clsgenOC_CadPendenciaProdFase
    
    ' Intância a classe de mensagem
    Set objMensagem = New clsgenMensagem
    
    TelaCadastrarPend_Prod_Fase objCadastrarTipoPendencia.Descricao, objCadastrarTipoPendencia.ID_TIPO_PENDENCIA
    


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
                            "Houve um erro de execução no método:  CadastrarPendenciaPRodFase" & _
                            "da classe clsgenOC_CadastrarTipoPendencia"
    GoTo Finalizar

End Sub


Public Function GravarDados_PendProdFase() As Boolean

On Error GoTo ErrorHandler:

    Dim bRetorno As Boolean
    Dim iRow As Integer
    Dim iCol As Integer
    
    GravarDados_PendProdFase = False
    
    ' Instância a classe de dados para cadastrar usuários
    Set objCadPend_Prod_Fase = New clsgenPend_Prod_Fases

    ' Informar o nome do banco desejado
    objCadPend_Prod_Fase.InformaBanco frmgenMDI_SGAPlus.objBanco

    ' Intância a classe de mensagem
    Set objMensagem = New clsgenMensagem
    
    ' Informar o nome do objeto de mensagem
    objCadPend_Prod_Fase.InformaMensagem objMensagem
    
    ' Inicia a transação
    bRetorno = frmgenMDI_SGAPlus.objBanco.Iniciar_Transacao()

    ' Se ocorreu algum erro na DLL
    If Not bRetorno Then

        GoTo Finalizar

    End If
    
    ' Excutar o método para gravar os dados
    'Verifica array para saber se info será excluida
    iCol = 1
    sTipoOperacao = ""
    
    For iRow = 1 To UBound(aListaPend, 2)
            'Verifica se ainda tem info no array
            If IsEmpty(aListaPend(iCol, iRow)) Then Exit For
            
            If IsNull(aListaPend(iCol, iRow)) Then
                'Marcado para exclusão
                sTipoOperacao = "E"
                
            ElseIf aListaPend(iCol, iRow) = "NOVO" Then
                'Novo registro
                sTipoOperacao = "I"
            End If
            
            If sTipoOperacao <> "" Then
                'Pega Codigos de Pendencia, Produto e Fase
                objCadPend_Prod_Fase.Tipo_Pendencia_ID = aListaPend(0, iRow)
                objCadPend_Prod_Fase.Produto_ID = aListaPend(2, iRow)
                objCadPend_Prod_Fase.Codigo_Tipo_Fase = aListaPend(4, iRow)
                
                bRetorno = objCadPend_Prod_Fase.GravarDados(sTipoOperacao)
                
                ' Se ocorreu algum erro na DLL
                If Not bRetorno Then
                
                    GoTo Finalizar
                
                End If
                sTipoOperacao = ""
            End If
    Next
        
    ' Finaliza a transação
    frmgenMDI_SGAPlus.objBanco.Finalizar_Transacao
    
    ' Atualiza a mensagem de retorno como gravação realizada com sucesso
    objMensagem.TipoMsg = "A"
    objMensagem.Descricao = "Gravação realizada com sucesso"

    GravarDados_PendProdFase = True

Finalizar:

    Screen.MousePointer = vbDefault

    If objMensagem.TipoMsg <> "" Then

        objMensagem.ExibeMensagem

    End If
    
    If Not GravarDados_PendProdFase Then
    
        frmgenMDI_SGAPlus.objBanco.Cancelar_Transacao

    End If

    Set objMensagem = Nothing
    
    On Error GoTo 0

    Exit Function

ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: GravarDados " & _
                            "da classe clsgenOC_CadastrarTipoPendencia"

    GoTo Finalizar

End Function


Public Sub TelaCadastrarPend_Prod_Fase(ByVal sTexto As String, ByVal lID As Long)
'Parametros sTexto e lID serão descricao e ID passado pelas telas do Usuario ou da CF

    Dim bRetorno As Boolean
    
On Error GoTo ErrorHandler:
    
    ' Instância o formulário de pré ativação e chama em seguida
    Set objFormCadPend_Prod_Fase = New frmgenCadPend_Prod_Fase

    ' Intância a classe de mensagem
    Set objMensagem = New clsgenMensagem

    ' Informar o objeto controlador
    objFormCadPend_Prod_Fase.InformaClasseOC Me

    ' Instância a classe de dados para cadastrar usuários
    Set objCadPend_Prod_Fase = New clsgenPend_Prod_Fases

    ' Informar o nome do banco desejado
    objCadPend_Prod_Fase.InformaBanco frmgenMDI_SGAPlus.objBanco

    ' Informar o nome do objeto de mensagem
    objCadPend_Prod_Fase.InformaMensagem objMensagem

    ' Chama o método de exibição do GRID
    objFormCadPend_Prod_Fase.lblPendencia.Caption = sTexto


    'Monta Combo Produto e Fase
    Call ExibirListaProduto(objFormCadPend_Prod_Fase.cboProduto)

    'Monta inicialmente lista de Fases baseada no Produto BLF e Acao=Ativação
    Call ExibirListaFase(GbIDProdutoBLF, GbIDAcaoAtivacao, objFormCadPend_Prod_Fase.cboFase)

    objFormCadPend_Prod_Fase.grdAlocacoes.Col = 1
    objFormCadPend_Prod_Fase.grdAlocacoes.ColHidden = False
    objFormCadPend_Prod_Fase.grdAlocacoes.Col = 2
    objFormCadPend_Prod_Fase.grdAlocacoes.ColHidden = False

    'Atribui valores parametros nas propriedades para tela Usu
    objFormCadPend_Prod_Fase.lblPendencia = sTexto
    objFormCadPend_Prod_Fase.Id_Pendencia = lID


    'Seleciona CFs do Usuario corrente
    bRetorno = ExibirListaGrid(lID)

    objFormCadPend_Prod_Fase.grdAlocacoes.Col = 1

    objFormCadPend_Prod_Fase.grdAlocacoes.Row = objFormCadPend_Prod_Fase.grdAlocacoes.ActiveRow

    If objFormCadPend_Prod_Fase.grdAlocacoes.Text <> "" Then
        objFormCadPend_Prod_Fase.cboProduto = objFormCadPend_Prod_Fase.grdAlocacoes.Text
    End If
    
    objFormCadPend_Prod_Fase.grdAlocacoes.Col = 2
        
    If objFormCadPend_Prod_Fase.grdAlocacoes.Text <> "" Then
        'objFormCadPend_Prod_Fase.cboFase = objFormCadPend_Prod_Fase.grdAlocacoes.Text
    End If
    
        
    ' Se ocorreu algum erro na DLL
    If Not bRetorno Then
        
        GoTo Finalizar
    
    End If
    
    
Finalizar:

    Screen.MousePointer = vbDefault
    
    If objMensagem.TipoMsg <> "" Then
    
        objMensagem.ExibeMensagem
        
    End If
    
    On Error GoTo 0
    
    Exit Sub
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: TelaCadastrarPend_Prod_Fase " & _
                            "da classe clsgenOC_CadastrarTipoPendencia"
    GoTo Finalizar

End Sub


Private Sub PreencheGrid_PendProdFase()

    Dim iRow As Integer
    
    'Limpa o Grid
    objFormCadPend_Prod_Fase.grdAlocacoes.MaxRows = 1
    
    objFormCadPend_Prod_Fase.grdAlocacoes.Col = 1
    objFormCadPend_Prod_Fase.grdAlocacoes.Row = 1
    objFormCadPend_Prod_Fase.grdAlocacoes.Text = ""

    objFormCadPend_Prod_Fase.grdAlocacoes.Col = 2
    objFormCadPend_Prod_Fase.grdAlocacoes.Row = 1
    objFormCadPend_Prod_Fase.grdAlocacoes.Text = ""

    objFormCadPend_Prod_Fase.grdAlocacoes.Col = 3
    objFormCadPend_Prod_Fase.grdAlocacoes.Row = 1
    objFormCadPend_Prod_Fase.grdAlocacoes.Text = ""

    
    ' Ler o array para preencher o GRID

    
   For iRow = objFormCadPend_Prod_Fase.grdAlocacoes.DataRowCnt + 1 To UBound(aListaPend, 2)
        
        
        If Not IsEmpty(aListaPend(0, iRow)) Then
            'Produto
            objFormCadPend_Prod_Fase.grdAlocacoes.Col = 1
            objFormCadPend_Prod_Fase.grdAlocacoes.Row = iRow
            objFormCadPend_Prod_Fase.grdAlocacoes.Text = aListaPend(3, iRow)
            
            'Fases
            objFormCadPend_Prod_Fase.grdAlocacoes.Col = 2
            objFormCadPend_Prod_Fase.grdAlocacoes.Text = aListaPend(5, iRow)
            
            'Flag para Identificar se é um Registro Novo
            objFormCadPend_Prod_Fase.grdAlocacoes.Col = 3
            objFormCadPend_Prod_Fase.grdAlocacoes.Text = "N"
            
            objFormCadPend_Prod_Fase.grdAlocacoes.MaxRows = iRow + 1
        Else
            objFormCadPend_Prod_Fase.grdAlocacoes.MaxRows = objFormCadPend_Prod_Fase.grdAlocacoes.DataRowCnt
            Exit For
            
        End If
    Next
    
End Sub


Public Sub ExcluirDoGrid_PendProdFase()
    
    Dim iContador As Integer
    Dim sCampo1, sCampo2 As String
    Dim iLinha As Integer
    
    'Verifica se grid esta vazio
    If objFormCadPend_Prod_Fase.grdAlocacoes.MaxRows = 0 Then
    
        Exit Sub
        
    End If

    'Seta variavel numero de linhas do grid
    'Pode haver linhas escondidas, que marcam os regs para exclusão
    iLinha = objFormCadPend_Prod_Fase.grdAlocacoes.MaxRows
    
    'no array marca com Excluido os registros a serem excluidos
    'ID Pendencia
    'aListaPend(0, objFormCadPend_Prod_Fase.grdAlocacoes.ActiveRow) = Null
    'Pendencia - Descricao
    aListaPend(1, objFormCadPend_Prod_Fase.grdAlocacoes.ActiveRow) = Null
    
    objFormCadPend_Prod_Fase.grdAlocacoes.Row = objFormCadPend_Prod_Fase.grdAlocacoes.ActiveRow
    objFormCadPend_Prod_Fase.grdAlocacoes.RowHidden = True
    
    'Atualiza texto combo
    objFormCadPend_Prod_Fase.grdAlocacoes.Row = objFormCadPend_Prod_Fase.grdAlocacoes.ActiveRow - 1
    
End Sub



Public Function CriticarCamposTela_PendProdFase() As Boolean

On Error GoTo ErrorHandler:

    CriticarCamposTela_PendProdFase = False
    
    If Len(objFormCadPend_Prod_Fase.cboFase.Text) = "" Then
        objMensagem.TipoMsg = "A"
        objMensagem.Descricao = "Selecione a Fase informação obrigatória."
        objFormCadPend_Prod_Fase.cboFase.SetFocus
        
        GoTo Finalizar
    End If
    
    If Len(objFormCadPend_Prod_Fase.cboProduto.Text) = "" Then
        objMensagem.TipoMsg = "A"
        objMensagem.Descricao = "Selecione o Produto, informação obrigatória."
        objFormCadPend_Prod_Fase.cboProduto.SetFocus
        
        GoTo Finalizar
    End If
    
    CriticarCamposTela_PendProdFase = True

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
                            "Houve um erro de execução no método: CriticarCamposTela " & _
                            "da classe clsgenOC_CadastrarPendencia_Produto_Fase"
    
    GoTo Finalizar

End Function


Public Sub IncluirNoGrid_PendProdFase(ByVal sTextoCbo As String, ByVal sTextoCbo2 As String, _
                         ByVal sPendencia As String, _
                         ByVal iLinha As Integer, _
                         ByVal piOrigem As Integer)


    Dim LcRow  As Integer
    Dim cColecao As New Collection
    Dim lRetorno As Long
    Dim lmax As Long
    Dim iColuna As Integer
    
    ' Intância a classe de mensagem
    Set objMensagem = New clsgenMensagem
       
    ' Intância a classe de mensagem
    Set objCadPend_Prod_Fase = New clsgenPend_Prod_Fases
    
    ' Informar o nome do banco desejado
    objCadPend_Prod_Fase.InformaBanco frmgenMDI_SGAPlus.objBanco
    objCadPend_Prod_Fase.InformaMensagem objMensagem
    
    'Verifca se grid esta vazio
    If objFormCadPend_Prod_Fase.grdAlocacoes.DataRowCnt <> 0 Then
              
        'Verifica se item do combo já existe no grid
        For LcRow = 1 To iLinha
            
            objFormCadPend_Prod_Fase.grdAlocacoes.Col = 1
            objFormCadPend_Prod_Fase.grdAlocacoes.Row = LcRow
            
            If objFormCadPend_Prod_Fase.grdAlocacoes.Text = sTextoCbo Then
                
                objFormCadPend_Prod_Fase.grdAlocacoes.Col = 2
                objFormCadPend_Prod_Fase.grdAlocacoes.Row = LcRow
                
                If objFormCadPend_Prod_Fase.grdAlocacoes.Text = sTextoCbo2 Then
                
                    If objFormCadPend_Prod_Fase.grdAlocacoes.RowHidden = False Then
                        ' Atualiza a mensagem de retorno
                        objMensagem.TipoMsg = "A"
                        objMensagem.Descricao = "Informação já cadastrada."
                    
                        GoTo Finalizar
                    Else
                         ''Verifica se o registro referente a colona Oculta ja se encontra na Base
                         '"S" igual registro novo ( não esta na Base )
                         objFormCadPend_Prod_Fase.grdAlocacoes.Col = 3
                         If objFormCadPend_Prod_Fase.grdAlocacoes.Text = "N" Then
                            aListaPend(1, objFormCadPend_Prod_Fase.grdAlocacoes.Row) = sPendencia
                         Else
                            aListaPend(1, objFormCadPend_Prod_Fase.grdAlocacoes.Row) = "NOVO"
                         End If
                        objFormCadPend_Prod_Fase.grdAlocacoes.RowHidden = False
                        GoTo Finalizar
                    End If
                End If
            End If
        Next
        objFormCadPend_Prod_Fase.grdAlocacoes.MaxRows = iLinha + 1
    Else
        'Não ha dados no grid
        lmax = 1000
        ReDim aListaPend(7, lmax)
        objFormCadPend_Prod_Fase.grdAlocacoes.MaxRows = iLinha
    End If
    
    'Inclui info no grid e no array
    objFormCadPend_Prod_Fase.grdAlocacoes.Row = objFormCadPend_Prod_Fase.grdAlocacoes.MaxRows
    
    
    
    ''Obter ID's do produto e Fase
    lRetorno = objCadPend_Prod_Fase.ObtemID(sPendencia)
    lID_TIPO_PENDENCIA = lRetorno
    
    'Coluna 1 Produto
     objFormCadPend_Prod_Fase.grdAlocacoes.Col = 1
     objFormCadPend_Prod_Fase.grdAlocacoes.Text = objFormCadPend_Prod_Fase.cboProduto
        
    'Coluna 2 Fase
     objFormCadPend_Prod_Fase.grdAlocacoes.Col = 2
     objFormCadPend_Prod_Fase.grdAlocacoes.Text = objFormCadPend_Prod_Fase.cboFase
        
    'Coluna 3 Flag q indica se é um novo Registro
     objFormCadPend_Prod_Fase.grdAlocacoes.Col = 3
     objFormCadPend_Prod_Fase.grdAlocacoes.Text = "S"
    '
        
        
    ' Atualiza o array
    'Pendencia_ID
     aListaPend(0, objFormCadPend_Prod_Fase.grdAlocacoes.Row) = lID_TIPO_PENDENCIA
         
    'Inclui Descricao referente a Pendencia -
    'Esta será utilizada como Flag para indicar que foi excluido (para Excluido valor = Null)
     aListaPend(1, objFormCadPend_Prod_Fase.grdAlocacoes.Row) = "NOVO"
        
    'ID Produto
     aListaPend(2, objFormCadPend_Prod_Fase.grdAlocacoes.Row) = objFormCadPend_Prod_Fase.cboProduto.ItemData(objFormCadPend_Prod_Fase.cboProduto.ListIndex)
        
    'Inclui Descricao do Produto referente ao CF selecionado no combo
     aListaPend(3, objFormCadPend_Prod_Fase.grdAlocacoes.Row) = objFormCadPend_Prod_Fase.cboProduto
        
    'ID Fase
     aListaPend(4, objFormCadPend_Prod_Fase.grdAlocacoes.Row) = objFormCadPend_Prod_Fase.cboFase.ItemData(objFormCadPend_Prod_Fase.cboFase.ListIndex)
        
    'Inclui Descricao da Fase  selecionada no combo
     aListaPend(5, objFormCadPend_Prod_Fase.grdAlocacoes.Row) = objFormCadPend_Prod_Fase.cboFase
        
    
Finalizar:

    Screen.MousePointer = vbDefault

    If objMensagem.TipoMsg <> "" Then

        objMensagem.ExibeMensagem

    End If
     
End Sub


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


Public Function ExibirListaGrid(ByVal lID_TIPO_PENDENCIA As Long) As Boolean

On Error GoTo ErrorHandler:

    Dim objMensagem As clsgenMensagem
    Dim bRetorno As Boolean
    Dim lID_Produto As Long
    
    ExibirListaGrid = False
    
    ' Intância a classe de mensagem
    Set objMensagem = New clsgenMensagem
    Set objCadPend_Prod_Fase = New clsgenPend_Prod_Fases
    
    ' Informar o nome do banco desejado
    objCadPend_Prod_Fase.InformaBanco frmgenMDI_SGAPlus.objBanco
                
    ' Informar o nome do objeto de mensagem
    objCadPend_Prod_Fase.InformaMensagem objMensagem
    
    'Verifica se o PRoduto foi selecionado
    'Caso não seja, atribui o Produto BLFlex como padrão
    If objFormCadPend_Prod_Fase.cboProduto.ListIndex >= 0 Then
        lID_Produto = objFormCadPend_Prod_Fase.cboProduto.ItemData(objFormCadPend_Prod_Fase.cboProduto.ListIndex)
    Else
        lID_Produto = GbIDProdutoBLF
    End If
    ' Executar o método de montagem da lista de centro funcional relacionados ao usuário
    bRetorno = objCadPend_Prod_Fase.MontaListaGrid(lID_TIPO_PENDENCIA, aListaPend, lID_Produto)
    
   ' Se ocorreu algum erro na DLL exibi-lo
    If Not bRetorno Then
        
        GoTo Finalizar
    
    End If
    
    ' Preenche a lista de usuários
    Call PreencheGrid_PendProdFase
    
    ExibirListaGrid = True
    
Finalizar:

    Screen.MousePointer = vbDefault
    
    If objMensagem.TipoMsg <> "" Then
    
        objMensagem.ExibeMensagem
        
    End If
    
    Set objCadPend_Prod_Fase = Nothing
    Set objMensagem = Nothing
    
    Exit Function

ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: ExibirListaGrid " & _
                            "da classe clsGenOC_CadastrarPendenciaProdutoFase"

    GoTo Finalizar

End Function


