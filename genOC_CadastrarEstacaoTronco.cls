VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGenOC_CadastrarEstacaoTronco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private objFormCadastrarEstacao As frmGenCadastroEstacaoTronco
Private objMensagem As clsgenMensagem
Private objCadastrarEstacao As clsGenEstacaoTronco
Private objListaEstacaoTronco As clsGenListaEstacaoTronco


Dim bRetorno As Boolean

Private sTipoOperacao As String

Public Property Get TipoOperacao() As String

    TipoOperacao = sTipoOperacao
    
End Property

Public Property Let TipoOperacao(ByVal sNewValue As String)

    sTipoOperacao = sNewValue
    
    ' Chama  método que controla a habilitação e desabilitação dos objetos,
    ' conforme o tipo de operação
    
End Property



Public Sub TelaCadastrarEstacao()

    ' Instância o formulário de pré ativação e chama em seguida
    Set objFormCadastrarEstacao = New frmGenCadastroEstacaoTronco
    
    ' Intância a classe de mensagem
    Set objMensagem = New clsgenMensagem
    
    ' Informar o objeto controlador
    objFormCadastrarEstacao.InformaClasseOC Me


'Instancia a classe Estação de dados.
    Set objCadastrarEstacao = New clsGenEstacaoTronco
    

    objCadastrarEstacao.InformaBanco frmgenMDI_SGAPlus.objBanco
    
    objCadastrarEstacao.InformaMensagem objMensagem

'Carregar os dados das estações na tela.
    bRetorno = objCadastrarEstacao.ListarDados
    
    bRetorno = PreencheLista
    
End Sub

Public Sub ExibirLista()
On Error GoTo ErrorHandler:

    Dim bRetorno As Boolean
    
    Screen.MousePointer = vbArrowHourglass
    
    ' Intância a classe de mensagem
    Set objMensagem = New clsgenMensagem
    
    ' Informar o nome do objeto de mensagem
    objCadastrarEstacao.InformaMensagem objMensagem
    
    ' Intância a classe da lista de usuários
    Set objListaEstacaoTronco = Nothing
    Set objListaEstacaoTronco = New clsGenListaEstacaoTronco
    
    ' Limpa a lista
    objFormCadastrarEstacao.lstEstacao.Clear

    ' Executar o método de montagem da lista na classe de analisar order entry
    bRetorno = objCadastrarEstacao.MontarLista(objFormCadastrarEstacao.Option1.Value, _
                                                                objFormCadastrarEstacao.txtBusca.Text)
                                               
                                               
    
    ' Se ocorreu algum erro na DLL exibi-lo, ou não retornou registro
    If Not bRetorno Then
    
        LimparDadosTelaCadastroEstacao
        
        GoTo Finalizar
    
    End If
    
    ' Preenche a lista de usuários
    bRetorno = PreencheLista()
    
    If Not bRetorno Then
        
        GoTo Finalizar
    
    End If
    
    
Finalizar:

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
                            "da classe clswidOC_CadastrarUsuario"

    GoTo Finalizar

End Sub

Public Function AlterouDadosTela() As Boolean

On Error GoTo ErrorHandler:

    AlterouDadosTela = True
    
    Set objMensagem = New clsgenMensagem
    
    If Trim(objCadastrarEstacao.CodEstacao) <> UCase(Trim(objFormCadastrarEstacao.txtCodEstacao.Text)) Then
    
        GoTo Finalizar
        
    End If

    If Trim(objCadastrarEstacao.NomeEstacao) <> UCase(Trim(objFormCadastrarEstacao.txtNmEstacao.Text)) Then
    
        GoTo Finalizar
        
    End If
    
    
    AlterouDadosTela = False
    
Finalizar:

    On Error GoTo 0

    Exit Function
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: AlterouDadosTela " & _
                            "da classe clsgenOC_CadastrarTroncos"
    
    GoTo Finalizar

End Function

Public Function SaiTela() As Boolean

On Error GoTo ErrorHandler:

    SaiTela = False

    ' Chama o método para verificar se foram alterados alguns dados
    If AlterouDadosTela Then
       
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
                            "da classe clsgenOC_CadastrarTroncos"
    
    GoTo Finalizar

End Function


Public Sub Saida()
    
    Unload objFormCadastrarEstacao
        
End Sub

Public Function GravarDados() As Boolean

    GravarDados = True
    
    Set objMensagem = New clsgenMensagem
    
    If Not carregaPropriedades Then
    
        On Error GoTo ErrorHandler:
    
    End If
    
    If criticaCampos Then
    
        frmgenMDI_SGAPlus.objBanco.Iniciar_Transacao
                
        bRetorno = objCadastrarEstacao.GravarDados(TipoOperacao)
        
        If Not bRetorno Then
        
            GoTo ErrorHandler:
        
        End If
    
' Finaliza a transação
        frmgenMDI_SGAPlus.objBanco.Finalizar_Transacao
              
'Limpar as propriedades do objeto
        objCadastrarEstacao.LimparPropriedades

        LimparDadosTelaCadastroEstacao

        sTipoOperacao = ""
        
        HabilitaDesabilitaObjetos
        
        objCadastrarEstacao.ListarDados
        
        PreencheLista
                    
    ' Atualiza a mensagem de retorno como gravação realizada com sucesso
        objMensagem.TipoMsg = "A"
        objMensagem.Descricao = "Gravação realizada com sucesso"

        GravarDados = True

    End If

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
                            "da classe clsgenOC_CadastrarEstacao"

    GoTo Finalizar
    

End Function

Public Function criticaCampos() As Boolean
    
On Error GoTo ErrorHandler:
    
    criticaCampos = True
    
    If objCadastrarEstacao.CodEstacao = "" Then
        
        GoTo ErrorHandler
        
    End If
    
    If objCadastrarEstacao.NomeEstacao = "" Then
    
        GoTo ErrorHandler
    
    End If
    
    Exit Function
    
Finalizar:

    If objMensagem.TipoMsg <> "" Then
    
        Screen.MousePointer = vbDefault
    
        criticaCampos = objMensagem.ExibeMensagem = "S"
    
    End If
    
    On Error GoTo 0
    
    Exit Function
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Por favor, preencher todos os campos !!!"
                            
    
    GoTo Finalizar
    
    
End Function

Public Function carregaPropriedades() As Boolean
    
On Error GoTo ErrorHandler:
    
    carregaPropriedades = True
    
    With objCadastrarEstacao
    
    
        .CodEstacao = TrataPlick(objFormCadastrarEstacao.txtCodEstacao.Text)
        .NomeEstacao = TrataPlick(objFormCadastrarEstacao.txtNmEstacao.Text)
        .Bairro = TrataPlick(objFormCadastrarEstacao.txtBairro.Text)
        .CEP = TrataPlick(objFormCadastrarEstacao.txtCEP.Text)
        .Cidade = TrataPlick(objFormCadastrarEstacao.txtCidade.Text)
        .Logradouro = TrataPlick(objFormCadastrarEstacao.txtLograd.Text)
        .Complemento = TrataPlick(objFormCadastrarEstacao.txtCompl.Text)
        .Numero = TrataPlick(objFormCadastrarEstacao.txtNumero.Text)
        .pais = TrataPlick(objFormCadastrarEstacao.txtPais)
        .UF = TrataPlick(objFormCadastrarEstacao.txtUF.Text)

    End With
    
Finalizar:

    If objMensagem.TipoMsg <> "" Then

        Screen.MousePointer = vbDefault

        carregaPropriedades = objMensagem.ExibeMensagem = "S"

    End If
    
    On Error GoTo 0
    
    Exit Function
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: SaiTela " & _
                            "da classe clsgenOC_CadastrarTroncos"
    
    carregaPropriedades = False
    
    GoTo Finalizar


End Function

Public Function RecuperarDados() As Boolean
        
    bRetorno = objCadastrarEstacao.RecuperarDados(objFormCadastrarEstacao.lstEstacao.ItemData(objFormCadastrarEstacao.lstEstacao.ListIndex))
    
    If Not bRetorno Then
    
        GoTo ErrorHandler
    
    End If
    
    With objFormCadastrarEstacao
        
        .txtCodEstacao.Text = objCadastrarEstacao.CodEstacao
        .txtNmEstacao.Text = objCadastrarEstacao.NomeEstacao
        .txtBairro.Text = objCadastrarEstacao.Bairro
        .txtCEP.Text = objCadastrarEstacao.CEP
        .txtCidade.Text = objCadastrarEstacao.Cidade
        .txtCompl.Text = objCadastrarEstacao.Complemento
        .txtLograd.Text = objCadastrarEstacao.Logradouro
        .txtNumero.Text = objCadastrarEstacao.Numero
        .txtPais.Text = objCadastrarEstacao.pais
        .txtUF.Text = objCadastrarEstacao.UF
    
    End With

Finalizar:
    
    Exit Function
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: Recuperar Dados " & _
                            "da classe clsgenOC_CadastrarEstacaoTroncos"
    
    
    GoTo Finalizar


End Function


Public Sub ExcluirDados()

On Error GoTo ErrorHandler:

    Dim bRetorno As Boolean

    ' Intância a classe de mensagem
    Set objMensagem = New clsgenMensagem
    
    objMensagem.TipoMsg = "Q"
    objMensagem.Descricao = "Confirma a exclusão da estação selecionada?"
    If objMensagem.ExibeMensagem = "S" Then

        ' Informar o nome do objeto de mensagem
        objCadastrarEstacao.InformaMensagem objMensagem
    
        If criticaCampos Then
    
            ' Inicia a transação
            frmgenMDI_SGAPlus.objBanco.Iniciar_Transacao
        
               ' Excutar o método de exclusão dos dados do usuário
                bRetorno = objCadastrarEstacao.ExcluirDados(objCadastrarEstacao.IdEstacao)
        
                ' Se ocorreu algum erro na DLL
                If Not bRetorno Then
        
                    frmgenMDI_SGAPlus.objBanco.Cancelar_Transacao
        
                    GoTo Finalizar
        
                End If
        
            ' Finaliza a transação
            frmgenMDI_SGAPlus.objBanco.Finalizar_Transacao
            
            ' Se existia somente um registro limpar o objeto de entrada para buscar
    '        If lQuantidadeRegistros = 1 Then
    '
    '            ' Limpa o campo de busca
    '            objFormCadastrarTipoPendencia.txtBusca.Text = ""
    '
    '        End If
    '
    '        ' Chama o método de exibição do GRID
    
            objCadastrarEstacao.LimparPropriedades
            
            LimparDadosTelaCadastroEstacao
            
            objCadastrarEstacao.ListarDados
            
            Call PreencheLista
            
            
    ' Atualiza a mensagem de retorno como gravação realizada com sucesso
            objMensagem.TipoMsg = "A"
            objMensagem.Descricao = "Exclusão realizada com sucesso"
            objMensagem.ExibeMensagem
        

        
        End If
            
            

        
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

Private Function PreencheLista() As Boolean

    lQuantidadeRegistros = 0
    
    PreencheLista = False
    
    ' Ler o objeto de coleção da lista de tipos de pendências
    
    objFormCadastrarEstacao.lstEstacao.Clear
    
    For Each objListaEstacaoTronco In objCadastrarEstacao.ListarEstacao
       
        With objFormCadastrarEstacao
            
            .lstEstacao.AddItem (objListaEstacaoTronco.CodEstacao & " " & objListaEstacaoTronco.NomeEstacao)
    
            .lstEstacao.ItemData(objFormCadastrarEstacao.lstEstacao.NewIndex) = CInt(objListaEstacaoTronco.IdEstacao)
    
        End With
        
    Next
        
    If objFormCadastrarEstacao.lstEstacao.ListCount > 0 Then ' Posiciona na primeira estacao
    
        objFormCadastrarEstacao.lstEstacao.ListIndex = 0
    
    End If
        
        
    PreencheLista = True
    
End Function

Public Sub LimparDadosTelaCadastroEstacao()

    With objFormCadastrarEstacao
    
        .txtCodEstacao.Text = ""
        .txtNmEstacao.Text = ""
'        .lstEstacao.ListIndex = -1
        .txtBairro.Text = ""
        .txtCEP.Text = ""
        .txtCidade.Text = ""
        .txtCompl.Text = ""
        .txtNumero.Text = ""
        .txtLograd.Text = ""
        .txtPais.Text = ""
        .txtUF.Text = ""
        
    End With
    
    objCadastrarEstacao.LimparPropriedades
    
    Exit Sub
        
End Sub

Public Sub HabilitaDesabilitaObjetos()

    ' Coloca cor de fundo como default para habilitação
    If (sTipoOperacao = "I") Or (sTipoOperacao = "A") Then ' Inclusão ou alteração
        
        ' Muda cor de fundo dos objetos de entrada de dados para branco
        With objFormCadastrarEstacao
        
            .txtCodEstacao.BackColor = &H80000005
            .txtNmEstacao.BackColor = &H80000005
            .txtBairro.BackColor = &H80000005
            .txtCEP.BackColor = &H80000005
            .txtCidade.BackColor = &H80000005
            .txtCompl.BackColor = &H80000005
            .txtNumero.BackColor = &H80000005
            .txtLograd.BackColor = &H80000005
            .txtPais.BackColor = &H80000005
            .txtUF.BackColor = &H80000005
        
        ' Muda cor de escrita dos objetos de entrada de dados para preto
        
            .txtCodEstacao.ForeColor = &H80000008
            .txtNmEstacao.ForeColor = &H80000008
            .txtBairro.ForeColor = &H80000008
            .txtCEP.ForeColor = &H80000008
            .txtCidade.ForeColor = &H80000008
            .txtCompl.ForeColor = &H80000008
            .txtLograd.ForeColor = &H80000008
            .txtNumero.ForeColor = &H80000008
            .txtPais.ForeColor = &H80000008
            .txtUF.ForeColor = &H80000008
        
        ' Habilita os objetos de dados
            
            .txtCodEstacao.Enabled = True
            .txtNmEstacao.Enabled = True
            .txtBairro.Enabled = True
            .txtCEP.Enabled = True
            .txtCidade.Enabled = True
            .txtCompl.Enabled = True
            .txtLograd.Enabled = True
            .txtNumero.Enabled = True
            .txtPais.Enabled = True
            .txtUF.Enabled = True
        
        ' Desabilita botões de inclusão, alteração, exclusão e frames de busca e lista
        
            .cmdIncluir.Enabled = False
            .cmdAlterar.Enabled = False
            .Cmdexcluir.Enabled = False
            
            .frameBusca.Enabled = False
            .frListaEstacao.Enabled = False
        
'        objFormCadastrarTronco.frameLista.Enabled = False
        
        ' Habilita os botões gravar, limpar e cancelar
            
            .CmdSalvar.Enabled = True
            .CmdLimpar.Enabled = True
            .cmdCancelar.Enabled = True
        
        End With
        
    Else
    
        ' Muda cor de fundo dos objetos de entrada de dados para cinza
        With objFormCadastrarEstacao
            
            .txtCodEstacao.BackColor = &H80000004
            .txtNmEstacao.BackColor = &H80000004
            .txtBairro.BackColor = &H80000004
            .txtCEP.BackColor = &H80000004
            .txtCidade.BackColor = &H80000004
            .txtCompl.BackColor = &H80000004
            .txtLograd.BackColor = &H80000004
            .txtNumero.BackColor = &H80000004
            .txtPais.BackColor = &H80000004
            .txtUF.BackColor = &H80000004

        ' Muda cor de escrita dos objetos de entrada de dados azul
            
            .txtCodEstacao.ForeColor = &HC00000
            .txtNmEstacao.ForeColor = &HC00000
            .txtBairro.ForeColor = &HC00000
            .txtCEP.ForeColor = &HC00000
            .txtCidade.ForeColor = &HC00000
            .txtCompl.ForeColor = &HC00000
            .txtLograd.ForeColor = &HC00000
            .txtNumero.ForeColor = &HC00000
            .txtPais.ForeColor = &HC00000
            .txtUF.ForeColor = &HC00000
        
        ' Desabilita o frame dos objetos de dados
'        objFormCadastrarTronco.frameDados.Enabled = False
    
        ' Habilita botões de inclusão, alteração, exlusão e frames de busca e lista

        
            .cmdIncluir.Enabled = True
            .cmdAlterar.Enabled = True
            .Cmdexcluir.Enabled = True
            
            .txtBusca.Enabled = True
            .frListaEstacao.Enabled = True
        
        ' Desabilita os botões gravar, limpar e cancelar
            .CmdSalvar.Enabled = False
            .CmdLimpar.Enabled = False
            .cmdCancelar.Enabled = False
        
        ' Desabilita os objetos de dados
            
            .txtCodEstacao.Enabled = False
            .txtNmEstacao.Enabled = False
            .txtBairro.Enabled = False
            .txtCEP.Enabled = False
            .txtCidade.Enabled = False
            .txtCompl.Enabled = False
            .txtLograd.Enabled = False
            .txtNumero.Enabled = False
            .txtPais.Enabled = False
            .txtUF.Enabled = False
            
        End With
    
    End If
    
    Exit Sub

End Sub


Private Sub Class_Terminate()
    
    Set objCadastrarEstacao = Nothing
    Set objMensagem = Nothing
    Set objListaEstacaoTronco = Nothing

End Sub
