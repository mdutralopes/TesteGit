VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsgenOC_CadastrarTronco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Declara as classes utilizadas
Private objMensagem         As clsgenMensagem
Private objCadastrarTronco As clsgenTronco
Private objListaTronco     As clsgenListaTroncos
Private objCadastrarEstacao As clsGenEstacaoTronco

' Declara os formulários usados nesta classe
Private objFormCadastrarTronco As frmgenCadastroTroncos

' Declara variáveis comuns
Private lID_USUARIO As Long
Private sTipoOperacao As String
Private lQuantidadeRegistros As Long



Public Property Get TipoOperacao() As String

    TipoOperacao = sTipoOperacao
    
End Property

Public Property Let TipoOperacao(ByVal sNewValue As String)

    sTipoOperacao = sNewValue
    
    ' Chama  método que controla a habilitação e desabilitação dos objetos,
    ' conforme o tipo de operação
    Call HabilitaDesabilitaObjetos
    
End Property

Public Sub TelaCadastrarTronco()

On Error GoTo ErrorHandler:

    'Cria um objeto para receber uma lista das estações.
    Dim cListaEstacao As clsGenListaEstacaoTronco
    
    'Cria um objeto para recuperar dados da velocidade.
    Dim objVelocidade As clsGenVelocidade
    
    'Array para receber uma lista das posições de equipamentos.
    Dim paListaPosicaoEqpto() As Variant
    
    'Array para receber uma lista das velocidades.
    Dim paListaVelocidade() As Variant
    
    Dim bRetorno As Boolean

' variaveis para receber a velocidade
    Dim intDigitoVeloc As Integer 'recebe a parte numerica da velocidade
    Dim strLetraVeloc As String 'recebe a parte alfanumerica da velocidade
    Dim i As Integer
    
    ' Instância o formulário de pré ativação e chama em seguida
    Set objFormCadastrarTronco = New frmgenCadastroTroncos
    
    ' Intância a classe de mensagem
    Set objMensagem = New clsgenMensagem
    
    ' Informar o objeto controlador
    objFormCadastrarTronco.InformaClasseOC Me
    
    ' Instância a classe de dados para cadastrar troncos
    Set objCadastrarTronco = New clsgenTronco
    
    ' Informar o nome do banco desejado
    objCadastrarTronco.InformaBanco frmgenMDI_SGAPlus.objBanco
            
    ' Informar o nome do objeto de mensagem
    objCadastrarTronco.InformaMensagem objMensagem
    
    ' Instância a classe de dados de estação
    Set objCadastrarEstacao = New clsGenEstacaoTronco
    
    'Instância a classe de dados de listar estações de tronco.
    Set cListaEstacao = New clsGenListaEstacaoTronco
    
    'Instância a classe de dados de velocidade.
    Set objVelocidade = New clsGenVelocidade
    
    
'adeil 01/08/2003
'popular os combos de estação
'inicio
    
    objCadastrarEstacao.InformaBanco frmgenMDI_SGAPlus.objBanco
    
    objCadastrarEstacao.InformaMensagem objMensagem
    
    objCadastrarEstacao.ListarDados
        
    For Each cListaEstacao In objCadastrarEstacao.ListarEstacao
    
        With objFormCadastrarTronco
            
            .cmbPonta1.AddItem cListaEstacao.CodEstacao
            .cmbPonta1.ItemData(.cmbPonta1.NewIndex) = CInt(cListaEstacao.IdEstacao)
            
            .cmbPonta2.AddItem cListaEstacao.CodEstacao
            .cmbPonta2.ItemData(.cmbPonta2.NewIndex) = CInt(cListaEstacao.IdEstacao)
            
        End With
    
    Next
'fim

'adeil 01/08/2003
'popular o combs de posição de equipamento.
'inicio
    
    bRetorno = objCadastrarTronco.MontaListaPosicaoEqpto(paListaPosicaoEqpto())

    If bRetorno Then
        PreencheCombo objFormCadastrarTronco.cmbPosEquip, paListaPosicaoEqpto()
    Else
        GoTo ErrorHandler:
    End If

'adeil 01/08/2003
'popular o combo de velocidade
'inicio
        
    objVelocidade.InformaBanco frmgenMDI_SGAPlus.objBanco
    
    objVelocidade.InformaMensagem objMensagem
    
' adeil,  a principio apenas o DIGINET é quem vai usar o cadastramento de tronco.
    bRetorno = objVelocidade.ListaVelocidade(paListaVelocidade(), "DNI")
        
    For i = 1 To UBound(paListaVelocidade, 2)
        
        If Not IsEmpty(paListaVelocidade(1, i)) Then
            
            bRetorno = obtemDigitoVelocidade(CStr(paListaVelocidade(1, i)), intDigitoVeloc, strLetraVeloc)
            objFormCadastrarTronco.cmbVeloc.AddItem intDigitoVeloc & strLetraVeloc
            objFormCadastrarTronco.cmbVeloc.ItemData(objFormCadastrarTronco.cmbVeloc.NewIndex) = intDigitoVeloc
        
        Else
        
            Exit For
        
        End If
    Next
    
    

'fim

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
                            "Houve um erro de execução no método: TelaCadastrar " & _
                            "da classe clsgenOC_CadastrarTronco"
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
    objCadastrarTronco.InformaMensagem objMensagem
    
    ' Intância a classe da lista de tronco
    Set objListaTronco = Nothing
    Set objListaTronco = New clsgenListaTroncos
    
    ' Limpa a lista
    objFormCadastrarTronco.lstTronco.Clear

    ' Executar o método de montagem da lista na classe de analisar order entry
    bRetorno = objCadastrarTronco.MontarLista()
    
    ' Se ocorreu algum erro na DLL exibi-lo
    If Not bRetorno Then
        
        GoTo Finalizar
    
    End If
    
    ' Preenche a lista de troncos
    bRetorno = PreencheLista()
    
    If Not bRetorno Then
        
        GoTo Finalizar
    
    End If
    
Finalizar:

    ' Se não encontrou registro desabilitar os botôes de alteração e exclusão
    If lQuantidadeRegistros = 0 Then
    
        objFormCadastrarTronco.cmdAlterar.Enabled = False
        objFormCadastrarTronco.cmdExcluir.Enabled = False
        
        Call LimparDadosTela
        
    Else
    
        objFormCadastrarTronco.cmdAlterar.Enabled = True
        objFormCadastrarTronco.cmdExcluir.Enabled = True
        
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
                            "da classe clsGenOC_CadastrarTronco"

    GoTo Finalizar

End Sub

Public Function GravarDados() As Boolean

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
        bRetorno = CriticarCamposTela
        
        If Not bRetorno Then
        
            GoTo Finalizar
            
        End If
    
        ' Informar o nome do objeto de mensagem
        objCadastrarTronco.InformaMensagem objMensagem
    
        ' Atribui os campos da tela nas propriedades
        Call AtribuirCamposDaTelaNasPropriedades
                
        ' Excutar o método para gravar os dados do tronco
        bRetorno = objCadastrarTronco.GravarDados(sTipoOperacao)
            
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
                            "da classe clsgenOC_CadastrarTronco"

    GoTo Finalizar

End Function

Public Sub ExcluirDados()

On Error GoTo ErrorHandler:

    Dim bRetorno As Boolean

    ' Intância a classe de mensagem
    Set objMensagem = New clsgenMensagem
    
    objMensagem.TipoMsg = "Q"
    objMensagem.Descricao = "Confirma a exlusão do tronco selecionado?"
    If objMensagem.ExibeMensagem = "S" Then

        ' Informar o nome do objeto de mensagem
        objCadastrarTronco.InformaMensagem objMensagem
    
        ' Inicia a transação
        frmgenMDI_SGAPlus.objBanco.Iniciar_Transacao
    
           ' Excutar o método de exclusão dos dados do usuário
            bRetorno = objCadastrarTronco.ExcluirDados(objCadastrarTronco.ID_TRONCO)
    
            ' Se ocorreu algum erro na DLL
            If Not bRetorno Then
    
                frmgenMDI_SGAPlus.objBanco.Cancelar_Transacao
    
                GoTo Finalizar
    
            End If
    
        ' Finaliza a transação
        frmgenMDI_SGAPlus.objBanco.Finalizar_Transacao
        
    ' Atualiza a mensagem de retorno como gravação realizada com sucesso
        objMensagem.TipoMsg = "A"
        objMensagem.Descricao = "Exclusão realizada com sucesso"
        objMensagem.ExibeMensagem
        
        ' Se existia somente um registro limpar o objeto de entrada para buscar
        If lQuantidadeRegistros = 1 Then
        
            ' Limpa o campo de busca
            objFormCadastrarTronco.txtBusca.Text = ""
            
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
                            "da classe clsgenOC_CadastrarTronco"

    Exit Sub

End Sub

Private Sub AtribuirCamposDaTelaNasPropriedades()

    objCadastrarTronco.Id_Ponta1 = objFormCadastrarTronco.cmbPonta1.ItemData(objFormCadastrarTronco.cmbPonta1.ListIndex)
    objCadastrarTronco.Ponta1 = objFormCadastrarTronco.cmbPonta1.Text
    
    objCadastrarTronco.Id_Ponta2 = objFormCadastrarTronco.cmbPonta2.ItemData(objFormCadastrarTronco.cmbPonta2.ListIndex)
    objCadastrarTronco.Ponta2 = objFormCadastrarTronco.cmbPonta2.Text
    
    objCadastrarTronco.Id_Velocidade = objFormCadastrarTronco.cmbVeloc.ItemData(objFormCadastrarTronco.cmbVeloc.ListIndex)
    objCadastrarTronco.Velocidade = objFormCadastrarTronco.cmbVeloc.Text
    
    objCadastrarTronco.Sequencial = CInt(objFormCadastrarTronco.txtSequencial.Text)
    
    objCadastrarTronco.Id_Pos_Eqpto = objFormCadastrarTronco.cmbPosEquip.ItemData(objFormCadastrarTronco.cmbPosEquip.ListIndex)
    objCadastrarTronco.Pos_Eqpto = objFormCadastrarTronco.cmbPosEquip.Text
    
    If objFormCadastrarTronco.chkInternacional.Value = 0 Then
        
        objCadastrarTronco.Internacional = "N"
    
    Else
        
        objCadastrarTronco.Internacional = "S"
    
    End If

    objCadastrarTronco.Eqpto_Pos_Campo1 = TrataPlick(objFormCadastrarTronco.txtCampo1.Text)
    objCadastrarTronco.Eqpto_Pos_Campo2 = TrataPlick(objFormCadastrarTronco.txtCampo2.Text)
    objCadastrarTronco.Eqpto_Pos_Campo3 = TrataPlick(objFormCadastrarTronco.txtCampo3.Text)
    objCadastrarTronco.Eqpto_Pos_Campo4 = TrataPlick(objFormCadastrarTronco.txtCampo4.Text)
    
    
    
    Select Case objCadastrarTronco.Velocidade
        
        Case "155M"
        
            objCadastrarTronco.quantidade_slot_disponivel = 63
        
        Case "34M"
            
            objCadastrarTronco.quantidade_slot_disponivel = 16
        
        Case "45M"
            
            objCadastrarTronco.quantidade_slot_disponivel = 21
        
        Case Else
            
            objCadastrarTronco.quantidade_slot_disponivel = 31
    
    End Select
    
End Sub

Private Sub Class_Terminate()

    Set objCadastrarTronco = Nothing
    Set objMensagem = Nothing
    Set objListaTronco = Nothing
    
End Sub

Private Function CriticarCamposTela() As Boolean

On Error GoTo ErrorHandler:

    CriticarCamposTela = False
    
    If objFormCadastrarTronco.cmbPonta1.ListIndex = -1 Then
    
        objMensagem.TipoMsg = "A"
        objMensagem.Descricao = "Ponta 1 obrigatória."
        objFormCadastrarTronco.cmbPonta1.SetFocus

        
        GoTo Finalizar
    
    ElseIf objFormCadastrarTronco.cmbPonta2.ListIndex = -1 Then
    
        objMensagem.TipoMsg = "A"
        objMensagem.Descricao = "Ponta 2 obrigatória."
        objFormCadastrarTronco.cmbPonta2.SetFocus
    
        GoTo Finalizar
        
    ElseIf objFormCadastrarTronco.cmbPonta1.ListIndex = objFormCadastrarTronco.cmbPonta2.ListIndex Then
    
        objMensagem.TipoMsg = "A"
        objMensagem.Descricao = "A Ponta 1 e Ponta 2 não podem ser iguais."
        objFormCadastrarTronco.cmbPonta1.SetFocus
    
        GoTo Finalizar
        
    
    ElseIf objFormCadastrarTronco.cmbVeloc.ListIndex = -1 Then
    
        objMensagem.TipoMsg = "A"
        objMensagem.Descricao = "Velocidade obrigatória."
        objFormCadastrarTronco.cmbVeloc.SetFocus
    
        GoTo Finalizar
    
    
    ElseIf Len(objFormCadastrarTronco.txtSequencial.Text) = 0 Then
    
        objMensagem.TipoMsg = "A"
        objMensagem.Descricao = "Número sequencial obrigatório."
        objFormCadastrarTronco.txtSequencial.SetFocus
    
        GoTo Finalizar
        
    
    ElseIf Len(objFormCadastrarTronco.txtCampo1.Text) = 0 Then 'Nome do No
    
        objMensagem.TipoMsg = "A"
        objMensagem.Descricao = objFormCadastrarTronco.texto1.Caption & " obrigatório."
        objFormCadastrarTronco.txtCampo1.SetFocus
    
        GoTo Finalizar
    
    ElseIf Len(objFormCadastrarTronco.txtCampo2.Text) = 0 Then 'Numero do No
    
        objMensagem.TipoMsg = "A"
        objMensagem.Descricao = objFormCadastrarTronco.texto2.Caption & " obrigatório."
        objFormCadastrarTronco.txtCampo2.SetFocus
    
        GoTo Finalizar
        
    ElseIf Len(objFormCadastrarTronco.txtCampo3.Text) = 0 Then 'Slot
    
        objMensagem.TipoMsg = "A"
        objMensagem.Descricao = objFormCadastrarTronco.texto3.Caption & " obrigatório."
        objFormCadastrarTronco.txtCampo3.SetFocus
    
        GoTo Finalizar
    
    ElseIf Len(objFormCadastrarTronco.txtCampo4.Text) = 0 And objFormCadastrarTronco.txtCampo4.Visible = True Then 'Porta
    
        objMensagem.TipoMsg = "A"
        objMensagem.Descricao = objFormCadastrarTronco.texto4.Caption & " obrigatório."
        objFormCadastrarTronco.txtCampo4.SetFocus
    
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
                            "Houve um erro de execução no método: CriticarCamposTela " & _
                            "da classe clsgenOC_CadastrarTronco"
    
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

Public Function AlterouDadosTela() As Boolean

On Error GoTo ErrorHandler:

    AlterouDadosTela = True
    
    Set objMensagem = New clsgenMensagem
    
    If Trim(objCadastrarTronco.Ponta1) <> Trim(objFormCadastrarTronco.cmbPonta1) Then
    
        GoTo Finalizar
        
    End If

    If Trim(objCadastrarTronco.Velocidade) <> Trim(objFormCadastrarTronco.cmbVeloc) Then
    
        GoTo Finalizar
        
    End If
    
    If Trim(objCadastrarTronco.Ponta2) <> Trim(objFormCadastrarTronco.cmbPonta2) Then
    
        GoTo Finalizar
        
    End If
    
    If Trim(objCadastrarTronco.Sequencial) <> 0 And _
       Trim(objCadastrarTronco.Sequencial) <> Trim(objFormCadastrarTronco.txtSequencial) Then
    
        GoTo Finalizar
        
    End If
    
'    If Trim(objCadastrarTronco.roteador) <> Trim(objFormCadastrarTronco.txtBusca) Then
'
'        GoTo Finalizar
'
'    End If
    
'adeil 25/07/2003
'RN-SGAD-92396FNG - CHANGE REQUEST DNI
'inicio
'    If Trim(objCadastrarTronco.SlotIP) <> Trim(objFormCadastrarTronco.txtSlotIP) Then
'
'        GoTo Finalizar
'
'    End If

'    If Trim(objCadastrarTronco.VipSlot) <> Trim(objFormCadastrarTronco.txtVipSlot) Then
'
'        GoTo Finalizar
'
'    End If

'    If Trim(objCadastrarTronco.PontaIP) <> Trim(objFormCadastrarTronco.txtPontaIP) Then
'
'        GoTo Finalizar
'
'    End If
'fim
    If Trim(objCadastrarTronco.Eqpto_Pos_Campo1) <> Trim(objFormCadastrarTronco.txtCampo1) Then   'Nome
    
        GoTo Finalizar
        
    End If
    
    If Trim(objCadastrarTronco.Eqpto_Pos_Campo2) <> Trim(objFormCadastrarTronco.txtCampo2) Then 'Numero
    
        GoTo Finalizar
        
    End If
    
    If Trim(objCadastrarTronco.Eqpto_Pos_Campo3) <> Trim(objFormCadastrarTronco.txtCampo3) Then 'Slot
    
        GoTo Finalizar
        
    End If
    
    If Trim(objCadastrarTronco.Eqpto_Pos_Campo4) <> Trim(objFormCadastrarTronco.txtCampo4) Then  'Porta
    
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

Private Function PreencheLista() As Boolean

    lQuantidadeRegistros = 0
    
    PreencheLista = False
    
    ' Ler o objeto de coleção da lista de troncos
    For Each objListaTronco In objCadastrarTronco.ListaTronco

        With objFormCadastrarTronco

'            If .optPonta1.Value = True Then

                ' Ponta IP
                .lstTronco.AddItem (objListaTronco.Ponta1 & " " & _
                                            objListaTronco.Ponta2 & " " & _
                                            objListaTronco.Velocidade & " " & _
                                            objListaTronco.Sequencial & "   " & _
                                            "Slots Disponíveis " & _
                                            objListaTronco.Quant_Slot_Disponivel)
                                            
                .lstTronco.ItemData(.lstTronco.NewIndex) = objListaTronco.ID_TRONCO

'            ElseIf objFormCadastrarTronco.optPonta2.Value = True Then

                ' Ponta E1
'                .lstTronco.AddItem (objListaTronco.Ponta2)

'            End If
            
        End With

        lQuantidadeRegistros = lQuantidadeRegistros + 1
    
    Next
    
    If objFormCadastrarTronco.lstTronco.ListCount > 0 Then
    
        objFormCadastrarTronco.lstTronco.ListIndex = 0
    
    End If
    
    PreencheLista = True
    
End Function

Public Sub RefreshDados()

    ' Atualiza os dados de tronco na tela
    With objFormCadastrarTronco
    
        .cmbPonta1 = objCadastrarTronco.ListaTronco(objFormCadastrarTronco.lstTronco.ListIndex + 1).PontaIP
        .cmbVeloc = objCadastrarTronco.ListaTronco(objFormCadastrarTronco.lstTronco.ListIndex + 1).Velocidade
        .cmbPonta2 = objCadastrarTronco.ListaTronco(objFormCadastrarTronco.lstTronco.ListIndex + 1).PontaE1
        .txtSequencial = objCadastrarTronco.ListaTronco(objFormCadastrarTronco.lstTronco.ListIndex + 1).Sequencial

'        .txtSlotIP = objCadastrarTronco.ListaTronco(objFormCadastrarTronco.lstTroncos.ListIndex + 1).SlotIP
'        .txtVipSlot = objCadastrarTronco.ListaTronco(objFormCadastrarTronco.lstTroncos.ListIndex + 1).VipSlot
'        .txtPortaIP = objCadastrarTronco.ListaTronco(objFormCadastrarTronco.lstTroncos.ListIndex + 1).PortaIP
'        .txtNomeNo = objCadastrarTronco.ListaTronco(objFormCadastrarTronco.lstTroncos.ListIndex + 1).NomeNo
'        .txtNumeroNo = objCadastrarTronco.ListaTronco(objFormCadastrarTronco.lstTroncos.ListIndex + 1).NumeroNo
'        .txtSlotE1 = objCadastrarTronco.ListaTronco(objFormCadastrarTronco.lstTroncos.ListIndex + 1).SlotE1
'        .txtPortaE1 = objCadastrarTronco.ListaTronco(objFormCadastrarTronco.lstTroncos.ListIndex + 1).PortaE1
    
'        objCadastrarTronco.ID_TRONCO = objCadastrarTronco.ListaTronco(objFormCadastrarTronco.lstTroncos.ListIndex + 1).ID_TRONCO
'        objCadastrarTronco.PontaIP = objCadastrarTronco.ListaTronco(objFormCadastrarTronco.lstTroncos.ListIndex + 1).PontaIP
'        objCadastrarTronco.Velocidade = objCadastrarTronco.ListaTronco(objFormCadastrarTronco.lstTroncos.ListIndex + 1).Velocidade
'        objCadastrarTronco.PontaE1 = objCadastrarTronco.ListaTronco(objFormCadastrarTronco.lstTroncos.ListIndex + 1).PontaE1
'        objCadastrarTronco.Sequencial = objCadastrarTronco.ListaTronco(objFormCadastrarTronco.lstTroncos.ListIndex + 1).Sequencial
'        objCadastrarTronco.roteador = objCadastrarTronco.ListaTronco(objFormCadastrarTronco.lstTroncos.ListIndex + 1).roteador
'        objCadastrarTronco.SlotIP = objCadastrarTronco.ListaTronco(objFormCadastrarTronco.lstTroncos.ListIndex + 1).SlotIP
'        objCadastrarTronco.VipSlot = objCadastrarTronco.ListaTronco(objFormCadastrarTronco.lstTroncos.ListIndex + 1).VipSlot
'        objCadastrarTronco.PortaIP = objCadastrarTronco.ListaTronco(objFormCadastrarTronco.lstTroncos.ListIndex + 1).PortaIP
'        objCadastrarTronco.NomeNo = objCadastrarTronco.ListaTronco(objFormCadastrarTronco.lstTroncos.ListIndex + 1).NomeNo
'        objCadastrarTronco.NumeroNo = objCadastrarTronco.ListaTronco(objFormCadastrarTronco.lstTroncos.ListIndex + 1).NumeroNo
'        objCadastrarTronco.SlotE1 = objCadastrarTronco.ListaTronco(objFormCadastrarTronco.lstTroncos.ListIndex + 1).SlotE1
'        objCadastrarTronco.PortaE1 = objCadastrarTronco.ListaTronco(objFormCadastrarTronco.lstTroncos.ListIndex + 1).PortaE1
    
    End With

End Sub

Public Sub HabilitaDesabilitaObjetos()

    ' Coloca cor de fundo como default para habilitação
    If (sTipoOperacao = "I") Or (sTipoOperacao = "A") Then ' Inclusão ou alteração
        
        ' Muda cor de fundo dos objetos de entrada de dados para branco
        objFormCadastrarTronco.cmbPonta1.BackColor = &H80000005
        objFormCadastrarTronco.cmbVeloc.BackColor = &H80000005
        objFormCadastrarTronco.cmbPonta2.BackColor = &H80000005
        objFormCadastrarTronco.txtSequencial.BackColor = &H80000005
        objFormCadastrarTronco.txtCampo1.BackColor = &H80000005
        objFormCadastrarTronco.txtCampo2.BackColor = &H80000005
        objFormCadastrarTronco.txtCampo3.BackColor = &H80000005
        objFormCadastrarTronco.txtCampo4.BackColor = &H80000005
        
        ' Muda cor de escrita dos objetos de entrada de dados para preto
        objFormCadastrarTronco.cmbPonta1.ForeColor = &H80000008
        objFormCadastrarTronco.cmbVeloc.ForeColor = &H80000008
        objFormCadastrarTronco.cmbPonta2.ForeColor = &H80000008
        objFormCadastrarTronco.txtSequencial.ForeColor = &H80000008
        objFormCadastrarTronco.txtCampo1.ForeColor = &H80000008
        objFormCadastrarTronco.txtCampo2.ForeColor = &H80000008
        objFormCadastrarTronco.txtCampo3.ForeColor = &H80000008
        objFormCadastrarTronco.txtCampo4.ForeColor = &H80000008
        
        ' Habilita o frame dos objetos de dados
        
        objFormCadastrarTronco.frameTronco.Enabled = True
        
        objFormCadastrarTronco.framePlacaE1.Enabled = True
        
        ' Desabilita botões de inclusão, alteração, exlusão e frames de busca e lista
        objFormCadastrarTronco.cmdIncluir.Enabled = False
        objFormCadastrarTronco.cmdAlterar.Enabled = False
        objFormCadastrarTronco.cmdExcluir.Enabled = False
        objFormCadastrarTronco.frameBusca.Enabled = False
'        objFormCadastrarTronco.frameLista.Enabled = False
        
        ' Habilita os botões gravar, limpar e cancelar
        objFormCadastrarTronco.cmdSalvar.Enabled = True
        objFormCadastrarTronco.cmdLimpar.Enabled = True
        objFormCadastrarTronco.cmdCancelar.Enabled = True
        
    Else
    
        ' Muda cor de fundo dos objetos de entrada de dados para cinza
        objFormCadastrarTronco.cmbPonta1.BackColor = &H80000004
        objFormCadastrarTronco.cmbVeloc.BackColor = &H80000004
        objFormCadastrarTronco.cmbPonta2.BackColor = &H80000004
        objFormCadastrarTronco.txtSequencial.BackColor = &H80000004
        objFormCadastrarTronco.txtCampo1.BackColor = &H80000004
        objFormCadastrarTronco.txtCampo2.BackColor = &H80000004
        objFormCadastrarTronco.txtCampo3.BackColor = &H80000004
        objFormCadastrarTronco.txtCampo4.BackColor = &H80000004

        
        ' Muda cor de escrita dos objetos de entrada de dados azul
        objFormCadastrarTronco.cmbPonta1.ForeColor = &HC00000
        objFormCadastrarTronco.cmbVeloc.ForeColor = &HC00000
        objFormCadastrarTronco.cmbPonta2.ForeColor = &HC00000
        objFormCadastrarTronco.txtSequencial.ForeColor = &HC00000
        objFormCadastrarTronco.txtCampo1.ForeColor = &HC00000
        objFormCadastrarTronco.txtCampo2.ForeColor = &HC00000
        objFormCadastrarTronco.txtCampo3.ForeColor = &HC00000
        objFormCadastrarTronco.txtCampo4.ForeColor = &HC00000
        
        ' Desabilita o frame dos objetos de dados
'        objFormCadastrarTronco.frameDados.Enabled = False
    
        ' Habilita botões de inclusão, alteração, exlusão e frames de busca e lista
        objFormCadastrarTronco.cmdIncluir.Enabled = True
        
        If objFormCadastrarTronco.lstTronco.ListCount > 0 Then
            
            objFormCadastrarTronco.cmdAlterar.Enabled = True
            objFormCadastrarTronco.cmdExcluir.Enabled = True
        
        End If
        
        objFormCadastrarTronco.frameBusca.Enabled = True
'        objFormCadastrarTronco.frameLista.Enabled = True
        
        ' Desabilita os botões gravar, limpar e cancelar
        objFormCadastrarTronco.cmdSalvar.Enabled = False
        objFormCadastrarTronco.cmdLimpar.Enabled = False
        objFormCadastrarTronco.cmdCancelar.Enabled = False
        
        ' Desabilita os FRAMES
        objFormCadastrarTronco.frameTronco.Enabled = False
        objFormCadastrarTronco.framePlacaE1.Enabled = False
    
    End If

End Sub

Public Sub LimparDadosTela()
    
    objFormCadastrarTronco.cmbPonta1.ListIndex = -1
    objFormCadastrarTronco.cmbPonta2.ListIndex = -1
    objFormCadastrarTronco.cmbVeloc.ListIndex = -1
    objFormCadastrarTronco.txtSequencial.Text = ""
    objFormCadastrarTronco.cmbPosEquip.ListIndex = -1
    objFormCadastrarTronco.txtCampo1.Text = ""
    objFormCadastrarTronco.txtCampo2.Text = ""
    objFormCadastrarTronco.txtCampo3.Text = ""
    objFormCadastrarTronco.txtCampo4.Text = ""
'    objFormCadastrarTronco.lstTronco.ListIndex = -1
    
    Exit Sub
        
End Sub

Public Sub Saida()
    
    Unload objFormCadastrarTronco
        
End Sub

Public Function RecuperarDados() As Boolean
        
        Dim bRetorno
        
        Dim iDigito As Integer
        Dim sLetra As String
        
        Dim lngCMB As Long
        
    bRetorno = objCadastrarTronco.RecuperarDados(objFormCadastrarTronco.lstTronco.ItemData(objFormCadastrarTronco.lstTronco.ListIndex()))
        
    
    For lngCMB = 0 To objFormCadastrarTronco.cmbPonta1.ListCount - 1
        If objFormCadastrarTronco.cmbPonta1.ItemData(lngCMB) = CInt(objCadastrarTronco.Id_Ponta1) Then
            objFormCadastrarTronco.cmbPonta1.ListIndex = lngCMB
            Exit For
        End If
    Next lngCMB
    
    For lngCMB = 0 To objFormCadastrarTronco.cmbPonta2.ListCount - 1
        If objFormCadastrarTronco.cmbPonta2.ItemData(lngCMB) = CInt(objCadastrarTronco.Id_Ponta2) Then
            objFormCadastrarTronco.cmbPonta2.ListIndex = lngCMB
            Exit For
        End If
    Next lngCMB
    
    'objFormCadastrarTronco.cmbPonta1.ListIndex = CInt(objCadastrarTronco.Id_Ponta1)
       
    'objFormCadastrarTronco.cmbPonta2.ListIndex = CInt(objCadastrarTronco.Id_Ponta2)
    
    bRetorno = obtemDigitoVelocidade(objCadastrarTronco.Velocidade, iDigito, sLetra)
    
    For lngCMB = 0 To objFormCadastrarTronco.cmbVeloc.ListCount - 1
        If objFormCadastrarTronco.cmbVeloc.ItemData(lngCMB) = CInt(iDigito) Then
            objFormCadastrarTronco.cmbVeloc.ListIndex = lngCMB
            Exit For
        End If
    Next lngCMB
    
    objFormCadastrarTronco.txtSequencial.Text = objCadastrarTronco.Sequencial
    
    For lngCMB = 0 To objFormCadastrarTronco.cmbPosEquip.ListCount - 1
        If objFormCadastrarTronco.cmbPosEquip.ItemData(lngCMB) = CInt(objCadastrarTronco.Id_Pos_Eqpto) Then
            objFormCadastrarTronco.cmbPosEquip.ListIndex = lngCMB
            Exit For
        End If
    Next lngCMB
    
    If objCadastrarTronco.Internacional = "S" Then
    
        objFormCadastrarTronco.chkInternacional.Value = 1
    
    Else
        
        objFormCadastrarTronco.chkInternacional.Value = 0
    
    End If
    
    objFormCadastrarTronco.cmbPosEquip.Text = objCadastrarTronco.Pos_Eqpto
    
    objFormCadastrarTronco.txtCampo1.Text = objCadastrarTronco.Eqpto_Pos_Campo1
    
    objFormCadastrarTronco.txtCampo2.Text = objCadastrarTronco.Eqpto_Pos_Campo2
    
    objFormCadastrarTronco.txtCampo3.Text = objCadastrarTronco.Eqpto_Pos_Campo3
    
    objFormCadastrarTronco.txtCampo4.Text = objCadastrarTronco.Eqpto_Pos_Campo4
                
    
End Function

Public Sub PesquisarTroncos()

On Error GoTo ErrorHandler:

    Dim bRetorno As Boolean
    
    Screen.MousePointer = vbArrowHourglass
    
    lQuantidadeRegistros = 0

    ' Intância a classe de mensagem
    Set objMensagem = New clsgenMensagem
    
    ' Informar o nome do objeto de mensagem
    objCadastrarTronco.InformaMensagem objMensagem
    
    ' Intância a classe da lista de tronco
    Set objListaTronco = Nothing
    Set objListaTronco = New clsgenListaTroncos
    
    ' Limpa a lista
    objFormCadastrarTronco.lstTronco.Clear

    ' Executar o método de montagem da lista na classe de analisar order entry
    bRetorno = objCadastrarTronco.MostrarTroncos(objFormCadastrarTronco.optPonta1.Value, _
                                                                     objFormCadastrarTronco.optPonta2.Value, _
                                                                     objFormCadastrarTronco.txtBusca.Text)
    
    ' Se ocorreu algum erro na DLL exibi-lo
    If Not bRetorno Then
        
        GoTo Finalizar
    
    End If
    
    ' Preenche a lista de troncos
    bRetorno = PreencheLista()
    
    If Not bRetorno Then
        
        GoTo Finalizar
    
    End If
    
Finalizar:

    ' Se não encontrou registro desabilitar os botôes de alteração e exclusão
    If lQuantidadeRegistros = 0 Then
    
        objFormCadastrarTronco.cmdAlterar.Enabled = False
        objFormCadastrarTronco.cmdExcluir.Enabled = False
        
        Call LimparDadosTela
        
    Else
    
        objFormCadastrarTronco.cmdAlterar.Enabled = True
        objFormCadastrarTronco.cmdExcluir.Enabled = True
        
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
                            "da classe clsGenOC_CadastrarTronco"

    GoTo Finalizar

End Sub

Public Sub MudaCamposPosicaoEquipamento()

    Select Case objFormCadastrarTronco.cmbPosEquip.Text
    
        Case "NBC"
            
            With objFormCadastrarTronco
            
                .texto1.Caption = "Nome do nó"
                .texto2.Caption = "Nº do Nó"
                .texto3.Caption = "Slot"
                .texto4.Caption = "Porta"
            
                .texto1.Visible = True
                .texto2.Visible = True
                .texto3.Visible = True
                .texto4.Visible = True
            
                .txtCampo4.Visible = True
                
                .cmbEstBrasil.Visible = False
            
            End With
            
        Case "DXC"
                        
            With objFormCadastrarTronco
            
                .texto1.Caption = "Nº do DAC"
                .texto2.Caption = "NPC"
                .texto3.Caption = "TS"
                
                .texto1.Visible = True
                .texto2.Visible = True
                .texto3.Visible = True
                
                .texto4.Visible = False
                
                .txtCampo4.Visible = False
                
                .cmbEstBrasil.Visible = False
            
            End With
        
        Case "MCP"
        
            With objFormCadastrarTronco
            
                .texto1.Caption = "BAST"
                .texto2.Caption = "SIST"
                .texto3.Caption = "TRIB"
                
                .texto1.Visible = True
                .texto2.Visible = True
                .texto3.Visible = True
                
                .texto4.Visible = False
                
                .txtCampo4.Visible = False
                
                .cmbEstBrasil.Visible = False
            
            End With
        
        Case "IPX"
            
            With objFormCadastrarTronco
            
                .texto1.Caption = "Nome do nó"
                .texto2.Caption = "SLOT"
                .texto3.Caption = "PORTA"
                
                .texto1.Visible = True
                .texto2.Visible = True
                .texto3.Visible = True
                
                .texto4.Visible = False
                
                .txtCampo4.Visible = False
            
                .cmbEstBrasil.Visible = False
                
            End With
            
        Case "PSS"
            
            With objFormCadastrarTronco
            
                .texto1.Caption = "Nome do nó"
                .texto2.Caption = "LP"
                .texto3.Caption = "P"
                
                .texto1.Visible = True
                .texto2.Visible = True
                .texto3.Visible = True
                
                .texto4.Visible = False
                
                .txtCampo4.Visible = False
                
                .cmbEstBrasil.Visible = False
            
            End With
            
        Case "DPN"
        
            With objFormCadastrarTronco
            
                .texto1.Caption = "Nome do nó"
                .texto2.Caption = "PL"
                .texto3.Caption = "PO"
                
                .texto1.Visible = True
                .texto2.Visible = True
                .texto3.Visible = True
                
                .texto4.Visible = False
                
                .txtCampo4.Visible = False
                
                .cmbEstBrasil.Visible = False
            
            End With
            
        Case "SMA"
            
            With objFormCadastrarTronco
            
                .texto1.Caption = "Estação no Brasil"
                
                .cmbEstBrasil.Left = 2310
                .cmbEstBrasil.Top = 525
                .cmbEstBrasil.Visible = True
                .texto1.Visible = False
                .texto2.Visible = False
                .texto3.Visible = False
                .texto4.Visible = False
            
            End With
            
    End Select


End Sub


