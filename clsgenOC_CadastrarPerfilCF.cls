VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsgenOC_CadastrarPerfilCF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private objMensagem             As clsgenMensagem
Private objCadastrarPerfil_CF   As ClsgenPerfilCF

' Declara os formulários usados nesta classe
Private objFormCadastrarPerfil_CF  As frmgenCadastroPefilCF

'Array referente as informações do grid
Private aPerfilCF()             As Variant

Private lID_Perfil_CF           As Long
Private lID_CF                  As Long
Private sCentro_Funcional       As String
Private lQuantidadeRegistros    As Long

Private bSalvouAlteracao          As Boolean
Public Sub Saida()
        
        Unload objFormCadastrarPerfil_CF
    
End Sub
Public Function SaiTela() As Boolean

On Error GoTo ErrorHandler:

    
    SaiTela = False
    
    Set objMensagem = New clsgenMensagem
    
    ' Chama o método para verificar se foram alterados alguns dados
    If Not bSalvouAlteracao Then
       
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
                            "da classe clsgenOC_CadastrarPerfilCF"
    
    GoTo Finalizar

End Function


Private Sub Preenche_Grid(ByRef paPerfilCF() As Variant)
    Dim i As Integer
    
    With objFormCadastrarPerfil_CF
    
       .SpdPerfil.MaxRows = 0
        
        For i = 1 To UBound(paPerfilCF, 2)
            'If IsEmpty(paPerfilCF(0, i)) Then Exit For
            
            If IsEmpty(paPerfilCF(0, i)) Then Exit For
            
            .SpdPerfil.MaxRows = .SpdPerfil.MaxRows + 1
            Call .SpdPerfil.SetText(1, i, paPerfilCF(1, i))
            
            'Marca linha corrente
            .SpdPerfil.Row = i
            
            'Verifica Flag ATUALIZA
            If paPerfilCF(2, i) Then
                .SpdPerfil.Col = 2
                .SpdPerfil.Text = paPerfilCF(2, i)
            End If
            
            'Verifica Flag EXCLUI
            If paPerfilCF(3, i) Then
                .SpdPerfil.Col = 3
                .SpdPerfil.Text = paPerfilCF(3, i)

            End If
            
            'Verifica Flag CONSULTA
            If paPerfilCF(4, i) Then
                .SpdPerfil.Col = 4
                .SpdPerfil.Text = paPerfilCF(4, i)
            End If
           
        Next
    End With
End Sub

Private Sub Class_Initialize()
    Set objMensagem = New clsgenMensagem
End Sub

Private Sub Class_Terminate()
    
    Set objCadastrarPerfil_CF = Nothing
    Set objMensagem = Nothing


End Sub


Public Function GravarDados() As Boolean
'Grava as informações do grid
    Dim bRetorno        As Boolean
    Dim iRow            As Integer
    Dim sTipoOperacao   As String
    Dim iAtualiza       As Integer
    Dim iExclui         As Integer
    Dim iConsulta       As Integer
    Dim iCentroF        As Long
    Dim iIDFase         As Long
    
    GravarDados = False

    ' Intância a classe de mensagem
    Set objMensagem = New clsgenMensagem

    'Seta variaveis que marcam a posicao dos campos de check no array
    iAtualiza = 2
    iExclui = 3
    iConsulta = 4
    iCentroF = 5
    
    ' Inicia a transação
    bRetorno = frmgenMDI_SGAPlus.objBanco.Iniciar_Transacao()

        ' Se ocorreu algum erro na DLL
        If Not bRetorno Then

            GoTo Finalizar

        End If
    
        ' Informar o nome do objeto de mensagem
        objCadastrarPerfil_CF.InformaMensagem objMensagem
        
        'Percorre array e verifica dados e tipo de operacao
        For iRow = 1 To UBound(aPerfilCF, 2)
                'Verifica se ainda tem info no array
                If IsEmpty(aPerfilCF(1, iRow)) Then Exit For
                
                'Verifica se coluna CENTRO_FUNCIONAL_ID do array é <> zero
                If aPerfilCF(iCentroF, iRow) = 0 Then
                    'Novo Registro
                    
                    If aPerfilCF(iAtualiza, iRow) <> 0 Or _
                       aPerfilCF(iExclui, iRow) <> 0 Or _
                       aPerfilCF(iConsulta, iRow) <> 0 Then
                       
                       'Registro marcado para inclusão
                       sTipoOperacao = "I"
                        
                    End If
                Else
                    'Alteração CF
                    If aPerfilCF(iAtualiza, iRow) = 0 And _
                       aPerfilCF(iExclui, iRow) = 0 And _
                       aPerfilCF(iConsulta, iRow) = 0 Then
                       
                       'Registro marcado para exclusao
                       sTipoOperacao = "E"
                        
                    Else
                        'Houve alteração em alguma das opções
                        sTipoOperacao = "A"
                    End If
                    
                End If
                
                If sTipoOperacao <> "" Then
                '
                    bRetorno = objCadastrarPerfil_CF.GravarDadosPerfil(sTipoOperacao, _
                                                     Trim(aPerfilCF(iIDFase, iRow)), ID_CF, _
                                                     aPerfilCF(iAtualiza, iRow), _
                                                     aPerfilCF(iExclui, iRow), _
                                                     aPerfilCF(iConsulta, iRow))
                    ' Se ocorreu algum erro na DLL
                    If Not bRetorno Then
                    
                        GoTo Finalizar
                    
                    End If
                    sTipoOperacao = ""
                End If
                
        Next
        
        
    ' Finaliza a transação
    frmgenMDI_SGAPlus.objBanco.Finalizar_Transacao
            
    'Inicializa flag de Alteração
    bSalvouAlteracao = True
    
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
                            "da classe clsgenOC_CadastrarPerfilCF"

    GoTo Finalizar
End Function

Public Function AtualizaArray(ByVal pbCheck As Integer, ByVal piCol As Integer, ByVal piRow As Integer)
'Grava as informações do grid
    On Error GoTo 0
    
    aPerfilCF(piCol, piRow) = pbCheck
    
    bSalvouAlteracao = False
    
End Function


Public Sub TelaCadastrarPerfil_CF(ByVal pOrigem As Integer, ByVal sTexto As String, ByVal lid As Long)
'Tela Cad Centros Funcionais
'Parametros sTexto e lID serão descricao e ID passado pelas telas do Usuario ou da CF

On Error GoTo ErrorHandler:
    
    ' Instância o formulário de pré ativação e chama em seguida
    Set objFormCadastrarPerfil_CF = New frmgenCadastroPefilCF
    
    ' Intância a classe de mensagem
    Set objMensagem = New clsgenMensagem
    
    ' Informar o objeto controlador
    objFormCadastrarPerfil_CF.InformaClasseOC Me
    
    ' Instância a classe de dados para cadastrar usuários
    Set objCadastrarPerfil_CF = New ClsgenPerfilCF
    
    ' Informar o nome do banco desejado
    objCadastrarPerfil_CF.InformaBanco frmgenMDI_SGAPlus.objBanco
            
    ' Informar o nome do objeto de mensagem
    objCadastrarPerfil_CF.InformaMensagem objMensagem
               
    objFormCadastrarPerfil_CF.lblCF.Caption = sTexto
    
    ID_CF = lid
    Centro_Funcional = sTexto
    objCadastrarPerfil_CF.Centro_Funcional = sTexto
    objCadastrarPerfil_CF.ID_CF = lid
    
    ' Chama o método de exibição do GRID
    Call ExibirListaGrid
        
    'Marca flag de controle de alteração
    bSalvouAlteracao = True
    
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
                            "Houve um erro de execução no método: TelaCadastrarPerfilCF " & _
                            "da classe clsgenOC_CadastrarPerfilCF"
    GoTo Finalizar

End Sub



Public Function ExibirListaGrid() As Boolean

On Error GoTo ErrorHandler:

    Dim aLista() As Variant
    Dim objMensagem As clsgenMensagem
    Dim bRetorno As Boolean
    
    ExibirListaGrid = False
    
    lQuantidadeRegistros = 0
    
    ' Intância a classe de mensagem
    Set objMensagem = New clsgenMensagem

    ' Informar o nome do banco desejado
    objCadastrarPerfil_CF.InformaBanco frmgenMDI_SGAPlus.objBanco
            
    ' Informar o nome do objeto de mensagem
    objCadastrarPerfil_CF.InformaMensagem objMensagem
      
    ' Executar o método de montagem da lista na classe de duração da fase
    bRetorno = objCadastrarPerfil_CF.MontarListaFases(aPerfilCF)
        
    If Not bRetorno Then
        
        GoTo Finalizar
    
    End If
        
   ' Preenche a lista
    Call Preenche_Grid(aPerfilCF)
    
    If Not bRetorno Then
        
        GoTo Finalizar
    
    End If
    
    ExibirListaGrid = True
    
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
                            "Houve um erro de execução no método: ExibirListaGrid " & _
                            "da classe clsOC_CadastrarPerfilCF"
    
    GoTo Finalizar

End Function
Public Property Get ID_CF() As String

    ID_CF = lID_CF
    
End Property
Public Property Let ID_CF(ByVal sNewValue As String)

    lID_CF = sNewValue
    
End Property
Public Property Get ID_Perfil_CF() As Long

    ID_Perfil_CF = lID_Perfil_CF
    
End Property
Public Property Let ID_Perfil_CF(ByVal lNewValue As Long)

    lID_Perfil_CF = lNewValue
      
End Property

Public Property Get Centro_Funcional() As String

    Centro_Funcional = sCentro_Funcional
    
End Property
Public Property Let Centro_Funcional(ByVal sNewValue As String)

    sCentro_Funcional = sNewValue
    
End Property

