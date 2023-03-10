VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsgenOC_CadastrarDuracaoFase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Declara as classes utilizadas
Private objMensagem                 As clsgenMensagem
Private objCadastrarDuracaoFase     As clsgenDuracaoFase
Private objListaDuracaoFase         As clsgenListaDuracaoFase

' Declara os formul?rios usados nesta classe
Private objFormCadastrarDuracaoFase As frmgenCadastroDuracaoFase

' Declara vari?veis comuns
Private lID_DuracaoFase             As Long
Private sTipoOperacao               As String
Private lQuantidadeRegistros        As Long
Private bAtualizaLista              As Boolean

Public Property Get TipoOperacao() As String

    TipoOperacao = sTipoOperacao
    
End Property

Public Property Let TipoOperacao(ByVal sNewValue As String)

    sTipoOperacao = sNewValue
    
    ' Int?ncia a classe de mensagem
    Set objMensagem = New clsgenMensagem

    ' Chama  m?todo que controla a habilita??o e desabilita??o dos objetos,
    ' conforme o tipo de opera??o
    If lQuantidadeRegistros > 0 Then
        Call HabilitaDesabilitaObjetos
    Else
        objMensagem.TipoMsg = "A"
        objMensagem.Descricao = " ? obrigat?ria."
    End If
End Property

Public Sub TelaCadastrarDuracaoFase()

    Dim bRetorno As Boolean
    
On Error GoTo ErrorHandler:

    ' Inst?ncia o formul?rio
    Set objFormCadastrarDuracaoFase = New frmgenCadastroDuracaoFase
    
    ' Int?ncia a classe de mensagem
    Set objMensagem = New clsgenMensagem
    
    ' Informar o objeto controlador
    objFormCadastrarDuracaoFase.InformaClasseOC Me
    
    ' Inst?ncia a classe de dados para cadastrar usu?rios
    Set objCadastrarDuracaoFase = New clsgenDuracaoFase
    
    ' Informar o nome do banco desejado
    objCadastrarDuracaoFase.InformaBanco frmgenMDI_SGAPlus.ObjBanco
            
    ' Informar o nome do objeto de mensagem
    objCadastrarDuracaoFase.InformaMensagem objMensagem
    
    ' Chama o m?todo de exibi??o do GRID
    Call ExibirListaDuracaoFase(objFormCadastrarDuracaoFase.cboProduto, objFormCadastrarDuracaoFase.cboAcao)
    
    ' Prepara a vari?vel de controle de pesquisa como falsa
    bAtualizaLista = False
    
    objFormCadastrarDuracaoFase.cboProduto.ListIndex = 0
    objFormCadastrarDuracaoFase.cboAcao.ListIndex = 0
    
    bAtualizaLista = True
    
    bRetorno = ExibirListaPesquisa()

    If Not bRetorno Then
        
        Call LimparDadosTela
             
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
                            "Houve um erro de execu??o no m?todo: TelaCadastrarDuracaoFase " & _
                            "da classe clsgenOC_CadastrarDuracaoFase"
    GoTo Finalizar

End Sub

Public Function ExibirListaDuracaoFase(ByRef cboProd As ComboBox, ByRef cboAcao As ComboBox) As Boolean

On Error GoTo ErrorHandler:

    Dim objDuracaoFase As clsgenDuracaoFase
    Dim objMensagem As clsgenMensagem
    Dim bRetorno As Boolean
    Dim aListaProd() As Variant
    Dim aListaAcao() As Variant
    
    ExibirListaDuracaoFase = False
    
    ' Limpa a combo
    cboProd.Clear
    cboAcao.Clear
    
    ' Int?ncia a classe de mensagem
    Set objMensagem = New clsgenMensagem
    
    ' Inst?ncia a classe de centro funcional
    Set objDuracaoFase = New clsgenDuracaoFase

    ' Informar o nome do banco desejado
    objDuracaoFase.InformaBanco frmgenMDI_SGAPlus.ObjBanco
                
    ' Informar o nome do objeto de mensagem
    objDuracaoFase.InformaMensagem objMensagem
    
    ' Executar o m?todo de montagem da lista de centro funcional na classe de analisar order entry
    bRetorno = objDuracaoFase.MontarCombo(aListaProd, aListaAcao)
    
    ' Se ocorreu algum erro na DLL
    If Not bRetorno Then
        
        Exit Function
    
    End If
       
    If objMensagem.TipoMsg = "" Then
        
        'Combo Produto
        PreencheCombo cboProd, aListaProd
        
        'Combo A??o
        PreencheCombo cboAcao, aListaAcao
        
    End If
    
    ExibirListaDuracaoFase = True
    
Finalizar:

    Screen.MousePointer = vbDefault
    
    If objMensagem.TipoMsg <> "" Then
    
        objMensagem.ExibeMensagem
        
    End If
    
    Set objDuracaoFase = Nothing
    Set objMensagem = Nothing
    
    Exit Function
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execu??o no m?todo: ExibirListaDuracaoFase " & _
                            "da classe clswidOC_AtividadesAbertas"
    
    GoTo Finalizar

End Function

Public Function GravarDados() As Boolean

On Error GoTo ErrorHandler:

    Dim bRetorno As Boolean

    GravarDados = False

    ' Int?ncia a classe de mensagem
    Set objMensagem = New clsgenMensagem

    ' Inicia a transa??o
    bRetorno = frmgenMDI_SGAPlus.ObjBanco.Iniciar_Transacao()

        ' Se ocorreu algum erro na DLL
        If Not bRetorno Then

            GoTo Finalizar

        End If

        ' Criticar os campos da tela
        bRetorno = CriticarCamposTela
        
        If Not bRetorno Then
        
            GoTo Finalizar
            
        End If
    
        ' Informar o nome do objeto de mensagem
        objCadastrarDuracaoFase.InformaMensagem objMensagem
    
        ' Atribui os campos da tela nas propriedades
        Call AtribuirCamposDaTelaNasPropriedades
        
        ' Excutar o m?todo para gravar os dados da recusa
        bRetorno = objCadastrarDuracaoFase.GravarDados(sTipoOperacao)
            
        ' Se ocorreu algum erro na DLL
        If Not bRetorno Then

            GoTo Finalizar

        End If

    ' Finaliza a transa??o
    frmgenMDI_SGAPlus.ObjBanco.Finalizar_Transacao
            
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
    
        frmgenMDI_SGAPlus.ObjBanco.Cancelar_Transacao

    End If

    Set objMensagem = Nothing
    
    On Error GoTo 0

    Exit Function

ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execu??o no m?todo: GravarDados " & _
                            "da classe clsgenOC_CadastrarDuracaoFase"

    GoTo Finalizar

End Function

Public Function ExibirListaPesquisa() As Boolean

On Error GoTo ErrorHandler:

    Dim aLista() As Variant
    Dim objMensagem As clsgenMensagem
    Dim bRetorno As Boolean
    Dim iProduto    As Integer
    Dim iAcao       As Integer
    
    ExibirListaPesquisa = False
    
    lQuantidadeRegistros = 0
    
    ' Int?ncia a classe de mensagem
    Set objMensagem = New clsgenMensagem
        
    ' Esta linha foi criada, porque este m?todo tamb?m ? chamado no evento click das combos,
    ' portanto quando a tela ? ativada esta rotina e chamada 2 vezes, porque iniciamos o
    ' listindex das combos como zero.
    If bAtualizaLista Then
    
        ' Define o tipo da fase
        iAcao = objFormCadastrarDuracaoFase.cboAcao.ItemData(objFormCadastrarDuracaoFase.cboAcao.ListIndex)
            
        ' Define Centro Funciional
        iProduto = objFormCadastrarDuracaoFase.cboProduto.ItemData(objFormCadastrarDuracaoFase.cboProduto.ListIndex)
            
        ' Informar o nome do banco desejado
        objCadastrarDuracaoFase.InformaBanco frmgenMDI_SGAPlus.ObjBanco
                
        ' Informar o nome do objeto de mensagem
        objCadastrarDuracaoFase.InformaMensagem objMensagem
        
            ' Limpa a lista
        objFormCadastrarDuracaoFase.lstDuracaoFase.Clear
    
        ' Executar o m?todo de montagem da lista na classe de dura??o da fase
        bRetorno = objCadastrarDuracaoFase.MontarListaProduto_Acao_Fase(iProduto, iAcao)
            
        If Not bRetorno Then
            
            GoTo Finalizar
        
        End If
            
       ' Preenche a lista
        bRetorno = PreencheLista()
        
        If Not bRetorno Then
            
            GoTo Finalizar
        
        End If
        
        Call HabilitaDesabilitaObjetos
        
        ' Posiciona o na primeira lista caso encontre algum registro
        objFormCadastrarDuracaoFase.lstDuracaoFase.ListIndex = 0
    
    End If
    
    ExibirListaPesquisa = True
    
Finalizar:

    ' Se n?o encontrou registro desabilitar os bot?es de altera??o e exclus?o
    If lQuantidadeRegistros = 0 Then
    
        objFormCadastrarDuracaoFase.cmdAlterar.Enabled = False
        
        Call LimparDadosTela
        
    Else
    
        objFormCadastrarDuracaoFase.cmdAlterar.Enabled = True
        
    End If
    
    Screen.MousePointer = vbDefault
    
    If objMensagem.TipoMsg <> "" Then
    
        objMensagem.ExibeMensagem
        
    End If
    
    Set objMensagem = Nothing
    
    Exit Function
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execu??o no m?todo: ExibirListaDuracaoFase " & _
                            "da classe clsOC_CadastrarDuracaoFase"
    
    GoTo Finalizar

End Function


Private Sub AtribuirCamposDaTelaNasPropriedades()

    objCadastrarDuracaoFase.Duracao = TrataPlick(objFormCadastrarDuracaoFase.txtDuracao)
    'objCadastrarDuracaoFase.ID_DURACAO_FASE = objListaDuracaoFase.ID_DURACAO_FASE
End Sub

Private Sub Class_Terminate()

    Set objCadastrarDuracaoFase = Nothing
    Set objMensagem = Nothing
    Set objListaDuracaoFase = Nothing
    
End Sub

Private Function CriticarCamposTela() As Boolean

On Error GoTo ErrorHandler:

    Dim objMensagem As clsgenMensagem

    ' Int?ncia a classe de mensagem
    Set objMensagem = New clsgenMensagem

    CriticarCamposTela = False
    
    If Len(objFormCadastrarDuracaoFase.txtDuracao.Text) = 0 Then
    
        objMensagem.TipoMsg = "A"
        objMensagem.Descricao = "Dura??o ? obrigat?ria."
        objFormCadastrarDuracaoFase.txtDuracao.SetFocus
        
        GoTo Finalizar
    
    End If
    
    CriticarCamposTela = True

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
                            "Houve um erro de execu??o no m?todo: CriticarCamposTela " & _
                            "da classe clsgenOC_CadastrarDuracaoFase"
    
    GoTo Finalizar

End Function

Public Function SaiTela() As Boolean

On Error GoTo ErrorHandler:

    Dim bAlterou As Boolean

    SaiTela = False

    ' Chama o m?todo para verificar se foram alterados alguns dados
    If AlterouDadosTela Then
       
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
                            "da classe clsgenOC_CadastrarDuracaoFase"
    
    GoTo Finalizar

End Function

Public Function AlterouDadosTela() As Boolean

On Error GoTo ErrorHandler:

    AlterouDadosTela = True
    
    Set objMensagem = New clsgenMensagem
    
    If objCadastrarDuracaoFase.Duracao <> Val(objFormCadastrarDuracaoFase.txtDuracao) Then
    
        GoTo Finalizar
        
    End If

    AlterouDadosTela = False
    
Finalizar:

    On Error GoTo 0

    Exit Function
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execu??o no m?todo: AlterouDadosTela " & _
                            "da classe clsgenOC_CadastrarDuracaoFases"
    
    GoTo Finalizar

End Function

Private Function PreencheLista() As Boolean

    lQuantidadeRegistros = 0
    
    PreencheLista = False
    
    ' Ler o objeto de cole??o da lista de motivos recusa
    For Each objListaDuracaoFase In objCadastrarDuracaoFase.ListaDuracaoFase

        With objFormCadastrarDuracaoFase

            ' Descricao Motivo Recusa
            .lstDuracaoFase.AddItem (objListaDuracaoFase.Fase)
       
        End With

        lQuantidadeRegistros = lQuantidadeRegistros + 1
    
    Next
    
    PreencheLista = True
    
End Function

Public Sub RefreshDados()

    
    If lQuantidadeRegistros > 0 Then
    ' Atualiza os dados de usu?rio na tela
    'Posiciona no primeiro item da lista
    With objFormCadastrarDuracaoFase
    
        .txtDuracao = objCadastrarDuracaoFase.ListaDuracaoFase(objFormCadastrarDuracaoFase.lstDuracaoFase.ListIndex + 1).Duracao
        .lblFase = objCadastrarDuracaoFase.ListaDuracaoFase(objFormCadastrarDuracaoFase.lstDuracaoFase.ListIndex + 1).Fase
                
        objCadastrarDuracaoFase.ID_DURACAO_FASE = objCadastrarDuracaoFase.ListaDuracaoFase(objFormCadastrarDuracaoFase.lstDuracaoFase.ListIndex + 1).ID_DURACAO_FASE
        objCadastrarDuracaoFase.Codigo_Tipo_Fase = objCadastrarDuracaoFase.ListaDuracaoFase(objFormCadastrarDuracaoFase.lstDuracaoFase.ListIndex + 1).Codigo_Tipo_Fase
        objCadastrarDuracaoFase.Duracao = objCadastrarDuracaoFase.ListaDuracaoFase(objFormCadastrarDuracaoFase.lstDuracaoFase.ListIndex + 1).Duracao
        
        
    End With

    End If

End Sub

Private Sub HabilitaDesabilitaObjetos()

    ' Coloca cor de fundo como default para habilita??o
    If (sTipoOperacao = "I") Or (sTipoOperacao = "A") Then ' Inclus?o ou altera??o
        
        ' Muda cor de fundo dos objetos de entrada de dados para branco
        objFormCadastrarDuracaoFase.txtDuracao.BackColor = &H80000005
        
        ' Muda cor de escrita dos objetos de entrada de dados para preto
        objFormCadastrarDuracaoFase.txtDuracao.ForeColor = &H80000008
        
        ' Habilita o frame dos objetos de dados
        objFormCadastrarDuracaoFase.frameDados.Enabled = True
        
        ' Desabilita bot?es de inclus?o, altera??o, exlus?o e frames de busca e lista
        objFormCadastrarDuracaoFase.cmdAlterar.Enabled = False
        objFormCadastrarDuracaoFase.frameBusca.Enabled = False
        objFormCadastrarDuracaoFase.frameLista.Enabled = False
        
        ' Habilita os bot?es gravar, limpar e cancelar
        objFormCadastrarDuracaoFase.cmdSalvar.Enabled = True
        objFormCadastrarDuracaoFase.cmdLimpar.Enabled = True
        objFormCadastrarDuracaoFase.cmdCancelar.Enabled = True
        
    Else
    
        ' Muda cor de fundo dos objetos de entrada de dados para cinza
        objFormCadastrarDuracaoFase.txtDuracao.BackColor = &H80000004
        
        ' Muda cor de escrita dos objetos de entrada de dados azul
        objFormCadastrarDuracaoFase.txtDuracao.ForeColor = &HC00000
        
        ' Desabilita o frame dos objetos de dados
        objFormCadastrarDuracaoFase.frameDados.Enabled = False
    
        ' Habilita bot?es de inclus?o, altera??o, exlus?o e frames de busca e lista
        objFormCadastrarDuracaoFase.cmdAlterar.Enabled = True
        objFormCadastrarDuracaoFase.frameBusca.Enabled = True
        objFormCadastrarDuracaoFase.frameLista.Enabled = True
        
        ' Desabilita os bot?es gravar, limpar e cancelar
        objFormCadastrarDuracaoFase.cmdSalvar.Enabled = False
        objFormCadastrarDuracaoFase.cmdLimpar.Enabled = False
        objFormCadastrarDuracaoFase.cmdCancelar.Enabled = False
    
    End If

End Sub

Public Sub LimparDadosTela()

    objFormCadastrarDuracaoFase.txtDuracao.Text = ""
    objFormCadastrarDuracaoFase.lblFase.Caption = ""
    
    objCadastrarDuracaoFase.Duracao = 0
    
    Exit Sub
        
End Sub

Public Sub Saida()
    
    Unload objFormCadastrarDuracaoFase
        
End Sub

