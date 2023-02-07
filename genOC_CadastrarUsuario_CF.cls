VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsgenOC_CadastrarUsuario_CF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Declara as classes utilizadas
Private objMensagem                 As clsgenMensagem
Private objCadastrarUsuario_CF      As clsgenUsuario_CF

' Declara os formulários usados nesta classe
Private objFormCadastrarUsuario_CF  As frmgenCadastroUsuarios_CF

' Declara variáveis comuns
Private sLogin_CF                   As String
Private lID_Login_CF                As Long
Private iTipoOrigem                 As Integer
Private sTipoOperacao               As String
Private lQuantidadeRegistros        As Long
Private objListaUsuarios            As clsgenListaUsuarios

Private aListaCF()                   As Variant

Public Sub IncluirNoGrid(ByVal sTextoCbo As String, _
                         ByVal sLogin_CF As String, _
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
    Set objCadastrarUsuario_CF = New clsgenUsuario_CF
    
    
    
    If objFormCadastrarUsuario_CF.cboUsuCF.ListIndex < 0 Then
    
        Call ProcuraTextoNaCombo(objFormCadastrarUsuario_CF.cboUsuCF, objFormCadastrarUsuario_CF.cboUsuCF.Text)
        
        If objFormCadastrarUsuario_CF.cboUsuCF.ListIndex < 0 Then
        
            objMensagem.TipoMsg = "A"
            objMensagem.Descricao = "Usuário não cadastrado."
            
            GoTo Finalizar
        
        End If
        
    End If
    
    
    ' Informar o nome do banco desejado
    objCadastrarUsuario_CF.InformaBanco frmgenMDI_SGAPlus.objBanco
    
    'Verifica origem - 1 Tela Cad. Usuarios
    '                  2 Tela Cad. Centros F.
    If TipoOrigem = 1 Then
        'Se Tela Cad usuarios - Mostra so coluna CF
        iColuna = 2
    ElseIf TipoOrigem = 2 Then
        'Se Tela Centros Func. - Mostra so Coluna Usu
        iColuna = 1
    End If
    
    'Verifca se grid esta vazio
    If objFormCadastrarUsuario_CF.grdAlocacoes.DataRowCnt <> 0 Then
              
        'Verifica se item do combo já existe no grid
        For LcRow = 1 To iLinha
            
            objFormCadastrarUsuario_CF.grdAlocacoes.Col = iColuna
            objFormCadastrarUsuario_CF.grdAlocacoes.Row = LcRow
            
            If objFormCadastrarUsuario_CF.grdAlocacoes.Text = UCase(sTextoCbo) Then
                ' Atualiza a mensagem de retorno
                objMensagem.TipoMsg = "A"
                objMensagem.Descricao = "Informação já cadastrada."
                
                GoTo Finalizar
                
            End If
        Next
        objFormCadastrarUsuario_CF.grdAlocacoes.MaxRows = iLinha + 1
    Else
        'Não ha dados no grid
        lmax = 1000
        ReDim aListaCF(4, lmax)
        objFormCadastrarUsuario_CF.grdAlocacoes.MaxRows = iLinha
    End If
    
    'Inclui info no grid e no array
    objFormCadastrarUsuario_CF.grdAlocacoes.Row = objFormCadastrarUsuario_CF.grdAlocacoes.MaxRows
    
    lRetorno = objCadastrarUsuario_CF.ObtemID(sLogin_CF, TipoOrigem)
    
    If TipoOrigem = 1 Then 'Tela Cad Usuario
        'Coluna 1 Usuario - Hidden
        objFormCadastrarUsuario_CF.grdAlocacoes.Col = 1
        objFormCadastrarUsuario_CF.grdAlocacoes.Text = sLogin_CF
        
        'Coluna 2 CFs
        objFormCadastrarUsuario_CF.grdAlocacoes.Col = 2
        If Len(objFormCadastrarUsuario_CF.grdAlocacoes.Text) = 0 Then
        
            objFormCadastrarUsuario_CF.grdAlocacoes.Text = UCase(sTextoCbo)

        End If
        
        ' Atualiza o array
        'USU
        aListaCF(0, objFormCadastrarUsuario_CF.grdAlocacoes.Row) = "NOVO"
         
        'Inclui Descricao referente ao CF selecionado no combo
        aListaCF(1, objFormCadastrarUsuario_CF.grdAlocacoes.Row) = objFormCadastrarUsuario_CF.cboUsuCF.Text
        
        'ID USU
        aListaCF(2, objFormCadastrarUsuario_CF.grdAlocacoes.Row) = lRetorno
        
        'Inclui ID CF0 referente ao CF selecionado no combo
        aListaCF(3, objFormCadastrarUsuario_CF.grdAlocacoes.Row) = objFormCadastrarUsuario_CF.cboUsuCF.ItemData(objFormCadastrarUsuario_CF.cboUsuCF.ListIndex)
        
    ElseIf TipoOrigem = 2 Then 'Tela Cad Centros Func
        
        'Coluna 2 CFs - Hidden
        objFormCadastrarUsuario_CF.grdAlocacoes.Col = 2
        objFormCadastrarUsuario_CF.grdAlocacoes.Text = sLogin_CF
        
        'Coluna 1 Usuario
        objFormCadastrarUsuario_CF.grdAlocacoes.Col = 1
        If Len(objFormCadastrarUsuario_CF.grdAlocacoes.Text) = 0 Then
        
            objFormCadastrarUsuario_CF.grdAlocacoes.Text = UCase(sTextoCbo)

        End If
        
        ' Atualiza o array
        'USU
        aListaCF(0, objFormCadastrarUsuario_CF.grdAlocacoes.Row) = objFormCadastrarUsuario_CF.cboUsuCF.Text
         
        'Inclui Codigo referente ao CF
        aListaCF(1, objFormCadastrarUsuario_CF.grdAlocacoes.Row) = "NOVO"
        
        'ID USU
        aListaCF(2, objFormCadastrarUsuario_CF.grdAlocacoes.Row) = objFormCadastrarUsuario_CF.cboUsuCF.ItemData(objFormCadastrarUsuario_CF.cboUsuCF.ListIndex)
        
        'Inclui ID CF0 referente ao CF selecionado no combo
        aListaCF(3, objFormCadastrarUsuario_CF.grdAlocacoes.Row) = lRetorno
    End If
    
Finalizar:

    Screen.MousePointer = vbDefault

    If objMensagem.TipoMsg <> "" Then

        objMensagem.ExibeMensagem

    End If
     
End Sub

Public Function CriticarCamposTela() As Boolean

On Error GoTo ErrorHandler:

    CriticarCamposTela = False
    
    If Len(objFormCadastrarUsuario_CF.cboUsuCF.Text) = "" Then
        objMensagem.TipoMsg = "A"
        objMensagem.Descricao = "Selecione informação no obrigatório."
        objFormCadastrarUsuario_CF.cboUsuCF.SetFocus
        
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
                            "da classe clsgenOC_CadastrarUsuario_CF"
    
    GoTo Finalizar

End Function


Public Sub ExcluirDoGrid()
    
    Dim iContador As Integer
    Dim sCampo1, sCampo2 As String
    Dim iLinha As Integer
    
    'Verifica se grid esta vazio
    If objFormCadastrarUsuario_CF.grdAlocacoes.MaxRows = 0 Then
    
        Exit Sub
        
    End If

    'Seta variavel numero de linhas do grid
    'Pode haver linhas escondidas, que marcam os regs para exclusão
    iLinha = objFormCadastrarUsuario_CF.grdAlocacoes.MaxRows
    
    'no array marca com NULL os registros a serem excluidos
    'login
    aListaCF(0, objFormCadastrarUsuario_CF.grdAlocacoes.ActiveRow) = Null
    'CF
    aListaCF(1, objFormCadastrarUsuario_CF.grdAlocacoes.ActiveRow) = Null
    
    objFormCadastrarUsuario_CF.grdAlocacoes.Row = objFormCadastrarUsuario_CF.grdAlocacoes.ActiveRow
    objFormCadastrarUsuario_CF.grdAlocacoes.RowHidden = True
    
    'Atualiza texto combo
    objFormCadastrarUsuario_CF.grdAlocacoes.Row = objFormCadastrarUsuario_CF.grdAlocacoes.ActiveRow + 1
    
    If objFormCadastrarUsuario_CF.grdAlocacoes.Text <> "" Then
        objFormCadastrarUsuario_CF.cboUsuCF.Text = objFormCadastrarUsuario_CF.grdAlocacoes.Text
    End If
    
    End Sub

Private Sub PreencheGrid()

    Dim iRow As Integer
    
    ' Ler o array para preencher o GRID
   For iRow = objFormCadastrarUsuario_CF.grdAlocacoes.DataRowCnt + 1 To UBound(aListaCF, 2)
        
        If Not IsEmpty(aListaCF(0, iRow)) Then
            'Usuario
            objFormCadastrarUsuario_CF.grdAlocacoes.Col = 1
            objFormCadastrarUsuario_CF.grdAlocacoes.Row = iRow
            objFormCadastrarUsuario_CF.grdAlocacoes.Text = aListaCF(0, iRow)
            
            'CF
            objFormCadastrarUsuario_CF.grdAlocacoes.Col = 2
            objFormCadastrarUsuario_CF.grdAlocacoes.Text = aListaCF(1, iRow)
            
            objFormCadastrarUsuario_CF.grdAlocacoes.MaxRows = iRow + 1
        Else
            objFormCadastrarUsuario_CF.grdAlocacoes.MaxRows = objFormCadastrarUsuario_CF.grdAlocacoes.DataRowCnt
            Exit For
            
        End If
    Next
    
End Sub




Public Property Get Login_CF() As String

    Login_CF = sLogin_CF
    
End Property
Public Property Let Login_CF(ByVal sNewValue As String)

    sLogin_CF = sNewValue
    
End Property
Public Property Get ID_Login_CF() As Long

    ID_Login_CF = lID_Login_CF
    
End Property
Public Property Let ID_Login_CF(ByVal lNewValue As Long)

    lID_Login_CF = lNewValue
      
End Property


Public Property Get TipoOrigem() As Integer

    TipoOrigem = iTipoOrigem
    
End Property

Public Property Let TipoOrigem(ByVal iNewValue As Integer)

    iTipoOrigem = iNewValue
    
    ' Chama  método que controla a habilitação e desabilitação dos objetos,
    ' conforme o tipo de operação
    
End Property

Public Sub TelaCadastrarUsuarios_CF(ByVal pOrigem As Integer, ByVal sTexto As String, ByVal lID As Long)
'Podera receber parametros de duas telas diferentes
'Parametro Origem = 1 --> Tela Cad Usuarios
'Parametro Origem = 2 --> Tela Cad Centros Funcionais
'Parametros sTexto e lID serão descricao e ID passado pelas telas do Usuario ou da CF

    Dim bRetorno As Boolean

On Error GoTo ErrorHandler:
    
    ' Instância o formulário de pré ativação e chama em seguida
    Set objFormCadastrarUsuario_CF = New frmgenCadastroUsuarios_CF
    
    ' Intância a classe de mensagem
    Set objMensagem = New clsgenMensagem
    
    ' Informar o objeto controlador
    objFormCadastrarUsuario_CF.InformaClasseOC Me
    
    ' Instância a classe de dados para cadastrar usuários
    Set objCadastrarUsuario_CF = New clsgenUsuario_CF
    
    ' Informar o nome do banco desejado
    objCadastrarUsuario_CF.InformaBanco frmgenMDI_SGAPlus.objBanco
            
    ' Informar o nome do objeto de mensagem
    objCadastrarUsuario_CF.InformaMensagem objMensagem
               
    ' Chama o método de exibição do GRID
    objFormCadastrarUsuario_CF.lblUsuCF.Caption = sTexto
     
    TipoOrigem = pOrigem
    
    If pOrigem = 1 Then
        'Monta Combo com CFs
        Call ExibirListaCentroFuncional(objFormCadastrarUsuario_CF.cboUsuCF)
        
        objFormCadastrarUsuario_CF.grdAlocacoes.Col = 1
        objFormCadastrarUsuario_CF.grdAlocacoes.ColHidden = True
        objFormCadastrarUsuario_CF.grdAlocacoes.Col = 2
        objFormCadastrarUsuario_CF.grdAlocacoes.ColHidden = False
        
        'Atribui valores parametros nas propriedades para tela Usu
        objCadastrarUsuario_CF.Login = sTexto
        objCadastrarUsuario_CF.ID_USUARIO_CF = lID
        
        objFormCadastrarUsuario_CF.lblOrigem = "Usuário"
        objFormCadastrarUsuario_CF.Label2.Caption = "Centro Funcional"
        
        'Seleciona CFs do Usuario corrente
        bRetorno = ExibirListaGrid(1, lID)
        
        objFormCadastrarUsuario_CF.grdAlocacoes.Col = 2
        
        objFormCadastrarUsuario_CF.grdAlocacoes.Row = objFormCadastrarUsuario_CF.grdAlocacoes.ActiveRow
        
        If objFormCadastrarUsuario_CF.grdAlocacoes.Text <> "" Then
            objFormCadastrarUsuario_CF.cboUsuCF = objFormCadastrarUsuario_CF.grdAlocacoes.Text
        End If
    
    Else
        Call ExibirListaUsuario(objFormCadastrarUsuario_CF.cboUsuCF)
        
        objFormCadastrarUsuario_CF.grdAlocacoes.Col = 1
        objFormCadastrarUsuario_CF.grdAlocacoes.ColHidden = False
        objFormCadastrarUsuario_CF.grdAlocacoes.Col = 2
        objFormCadastrarUsuario_CF.grdAlocacoes.ColHidden = True
    
        objCadastrarUsuario_CF.Centro_Funcional = sTexto
        objCadastrarUsuario_CF.ID_CF = lID
        
        objFormCadastrarUsuario_CF.lblOrigem = "Centro Funcional"
        objFormCadastrarUsuario_CF.Label2.Caption = "Usuários"
        
        'Seleciona CFs do Usuario corrente
        bRetorno = ExibirListaGrid(2, lID)
        
        objFormCadastrarUsuario_CF.grdAlocacoes.Col = 1
        
        objFormCadastrarUsuario_CF.grdAlocacoes.Row = objFormCadastrarUsuario_CF.grdAlocacoes.ActiveRow
        If objFormCadastrarUsuario_CF.grdAlocacoes.Text <> "" Then
            objFormCadastrarUsuario_CF.cboUsuCF = objFormCadastrarUsuario_CF.grdAlocacoes.Text
        End If
        
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
    
    Set objMensagem = Nothing
    
    On Error GoTo 0
    
    Exit Sub
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: TelaCadastrarUsuarios_CF " & _
                            "da classe clsgenOC_CadastrarUsuario_CF"
    GoTo Finalizar

End Sub

Public Function GravarDados() As Boolean

On Error GoTo ErrorHandler:

    Dim bRetorno As Boolean
    Dim iRow As Integer
    Dim iCol As Integer
    
    GravarDados = False
    
    ' Instância a classe de dados para cadastrar usuários
    Set objCadastrarUsuario_CF = New clsgenUsuario_CF

    ' Informar o nome do banco desejado
    objCadastrarUsuario_CF.InformaBanco frmgenMDI_SGAPlus.objBanco

    ' Intância a classe de mensagem
    Set objMensagem = New clsgenMensagem
    
    ' Informar o nome do objeto de mensagem
    objCadastrarUsuario_CF.InformaMensagem objMensagem
    
    ' Inicia a transação
    bRetorno = frmgenMDI_SGAPlus.objBanco.Iniciar_Transacao()

    ' Se ocorreu algum erro na DLL
    If Not bRetorno Then

        GoTo Finalizar

    End If
    
    ' Excutar o método para gravar os dados
    'Verifica array para saber se info será excluida
    If TipoOrigem = 1 Then
        iCol = 0
    Else
        iCol = 1
    End If
    
    For iRow = 1 To UBound(aListaCF, 2)
            'Verifica se ainda tem info no array
            If IsEmpty(aListaCF(1, iRow)) Then Exit For
            
            If IsNull(aListaCF(iCol, iRow)) Then
                'Marcado para exclusão
                sTipoOperacao = "E"
                
            ElseIf aListaCF(iCol, iRow) = "NOVO" Then
                'Novo registro
                sTipoOperacao = "I"
            End If
            
            If sTipoOperacao <> "" Then
                'Pega Codigos de Usuario e CF
                objCadastrarUsuario_CF.ID_USUARIO_CF = aListaCF(2, iRow)
                objCadastrarUsuario_CF.ID_CF = aListaCF(3, iRow)
                
                bRetorno = objCadastrarUsuario_CF.GravarDados(sTipoOperacao)
                
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
                            "da classe clsgenOC_CadastrarUsuario"

    GoTo Finalizar

End Function

Private Sub Class_Terminate()

    Set objCadastrarUsuario_CF = Nothing
    Set objMensagem = Nothing
    
End Sub

Public Sub Saida()
    
        Unload objFormCadastrarUsuario_CF
        
    
End Sub

Public Function ExibirListaCentroFuncional(ByRef cbo As ComboBox) As Boolean

On Error GoTo ErrorHandler:

    Dim objCentroFuncional As clsgenCentroFuncional
    Dim objMensagem As clsgenMensagem
    Dim bRetorno As Boolean
    Dim aListaCentroFuncional() As Variant
    
    ExibirListaCentroFuncional = False
    
    ' Limpa a combo
    cbo.Clear
    
    ' Intância a classe de mensagem
    Set objMensagem = New clsgenMensagem
    
    ' Instância a classe de centro funcional
    Set objCentroFuncional = New clsgenCentroFuncional

    ' Informar o nome do banco desejado
    objCentroFuncional.InformaBanco frmgenMDI_SGAPlus.objBanco
                
    ' Informar o nome do objeto de mensagem
    objCentroFuncional.InformaMensagem objMensagem
    
    ' Executar o método de montagem da lista de centro funcional na classe de analisar order entry
    bRetorno = objCentroFuncional.MontarListaCentroFuncional(aListaCentroFuncional)
    
    ' Se ocorreu algum erro na DLL
    If Not bRetorno Then
        
        Exit Function
    
    End If
    
    If objMensagem.TipoMsg = "" Then
    
        PreencheCombo cbo, aListaCentroFuncional
        
    End If
    
    ExibirListaCentroFuncional = True
    
Finalizar:

    Screen.MousePointer = vbDefault
    
    If objMensagem.TipoMsg <> "" Then
    
        objMensagem.ExibeMensagem
        
    End If
    
    Set objCentroFuncional = Nothing
    Set objMensagem = Nothing
    
    Exit Function
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: ExibirListaCentroFuncional " & _
                            "da classe clswidOC_AtividadesAbertas"
    
    GoTo Finalizar

End Function
Public Function ExibirListaUsuario(ByRef cbo As ComboBox) As Boolean

On Error GoTo ErrorHandler:

    Dim objMensagem As clsgenMensagem
    Dim bRetorno As Boolean
    Dim iLinha As Integer
    
    ExibirListaUsuario = False
    
    ' Limpa a combo
    cbo.Clear
    
    ' Intância a classe de mensagem
    Set objMensagem = New clsgenMensagem
    
    ' Informar o nome do banco desejado
    objCadastrarUsuario_CF.InformaBanco frmgenMDI_SGAPlus.objBanco
                
    ' Informar o nome do objeto de mensagem
    objCadastrarUsuario_CF.InformaMensagem objMensagem
    
    ' Executar o método de montagem da lista de centro funcional na classe de analisar order entry
    bRetorno = objCadastrarUsuario_CF.MontarListaUsuarios()
    
    ' Se ocorreu algum erro na DLL
    If Not bRetorno Then
        
        Exit Function
    
    End If
    
    If objMensagem.TipoMsg = "" Then
    
        ' Preencher a combo
        iLinha = 0
        For Each objListaUsuarios In objCadastrarUsuario_CF.ListaUsuario_CF
    
            cbo.AddItem objListaUsuarios.Login, iLinha
            cbo.ItemData(iLinha) = objListaUsuarios.ID_USUARIO
            
            iLinha = iLinha + 1
            
        Next
        
    End If
    
    ExibirListaUsuario = True
    
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
                            "Houve um erro de execução no método: ExibirListaCentroFuncional " & _
                            "da classe clswidOC_AtividadesAbertas"
    
    GoTo Finalizar

End Function

Public Function ExibirListaGrid(ByVal pOrigem As Integer, ByVal plID_USU_CF As Long) As Boolean

On Error GoTo ErrorHandler:

    Dim objMensagem As clsgenMensagem
    Dim bRetorno As Boolean
        
    ExibirListaGrid = False
    
    ' Intância a classe de mensagem
    Set objMensagem = New clsgenMensagem
    
    ' Informar o nome do banco desejado
    objCadastrarUsuario_CF.InformaBanco frmgenMDI_SGAPlus.objBanco
                
    ' Informar o nome do objeto de mensagem
    objCadastrarUsuario_CF.InformaMensagem objMensagem
    
    ' Executar o método de montagem da lista de centro funcional relacionados ao usuário
    bRetorno = objCadastrarUsuario_CF.MontaListaGrid(plID_USU_CF, aListaCF, pOrigem)
    
   ' Se ocorreu algum erro na DLL exibi-lo
    If Not bRetorno Then
        
        GoTo Finalizar
    
    End If
    
    ' Preenche a lista de usuários
    Call PreencheGrid
    
    ExibirListaGrid = True
    
Finalizar:

    Screen.MousePointer = vbDefault
    
    If objMensagem.TipoMsg <> "" Then
    
        objMensagem.ExibeMensagem
        
    End If
    
    Set objCadastrarUsuario_CF = Nothing
    Set objMensagem = Nothing
    
    Exit Function
        
ErrorHandler:

    objMensagem.TipoMsg = "E"
    objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Houve um erro de execução no método: ExibirListaGrid " & _
                            "da classe clsGenOC_CadastrarUsuario_CF"
    
    GoTo Finalizar

End Function






