VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGenOC_OperadoraPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Explicit declare variables
Option Explicit

' Ponteiros para classe de banco
Private objBanco As clsBanco 'As SGAPLUSBanco.clsBanco

' Novas estruturas de controle
Private objMsg As New Cls_Msg
Private objGenOperadora As New clsGenOperadora
Private objOC_OperadoraSecundario As New clsGenOC_OperadoraSecundario

' Cria um array para conter as informacoes
Private Operadoras() As Variant
'
'
' Carrega a visão do usuário para as Operadoras ( form )
'
Public Function CarregarForm() As Boolean
    
    ' Inicializa o retorno da função como FALSE
    CarregarForm = False
   
' Inicializa o trap de erros para o método
On Error GoTo ErrorHandler
    
    ' Cria uma nova instância da classe de dados de Operadoras
    'Set objGenOperadora = New clsGenOperadora
    ' Cria uma nova instância da classe de Mensagens
    'Set objMsg = New Cls_Msg

    ' Atribui ao ponteiro local do banco o ponteiro para o banco aberto
    Call objGenOperadora.InformaBanco(frmgenMDI_SGAPlus.objBanco, objMsg)
    
    ' Se ocorrer erro, pula para o trap
    If objMsg.TipoMsg <> "" Then GoTo ErrorHandler
    
    ' Se conseguir executar o método para o retrieve das informações...
    If objGenOperadora.ListarDados(Operadoras, objMsg) Then
        
        ' ... chama o método para popular a visão do usuário ( form )
        Call frmGenManterOperadorasPrincipal.PreencheGrid(Operadoras)
        
        ' Seta o retorno como TRUE
        CarregarForm = True
        
    End If
    
    ' Se NÃO ocorrer erro, sai da função.
    If objMsg.TipoMsg = "" Then Exit Function

ErrorHandler:

    ' utilizando o objeto de mensagens...
    With objMsg
    
        ' ... se NÃO existir tipo de mensagem associado...
        If .TipoMsg = "" Then
        
            ' ... setar o tipo da mensagem...
            .TipoMsg = "E"
            
            ' ... setar a mensagem...
            .TituloMsg = "clsgenOC_OperadoraPrincipal.CarregarForm"
            
            ' ... e uma descrição
            .Descricao = Err.Description
        End If
        
        ' Exibir a mensagem para o usuário
        .ExibeMensagem
    End With
    
End Function

Public Function MostrarDados( _
        ByVal piID As Integer, _
        ByRef sTipoGravacao As String _
        ) As Boolean

' Inicializa o TRAP de erros para o método
On Error GoTo ErrorHandler:

    MostrarDados = False
    
    ' Instância a classe de pré ativação
    'Set objPreAtivacao = New WID.clswidPreAtivacao
    
    ' Informar o nome do banco desejado
    'objPreAtivacao.InformaBanco frmgenMDI_SGAPlus.objBanco
            
    ' Excutar o método
    'bRetorno = objPreAtivacao.RecuperarDados(piID_Servico, paDadosPreAtivacao)
    
    ' Se ocorreu algum erro na DLL exibi-lo
    'If Not bRetorno Then
        
        'MsgBox ("Erro (MostrarDados)")
        
        'Exit Function
    
    'End If
    
    sTipoGravacao = "I"
    
    'If UBound(paDadosPreAtivacao) > 0 Then
    
'        dDataPedidoTeste = paDadosPreAtivacao(1, 1) ' Data do Pedido de Teste
'        dDataPrevistaRealizacao = paDadosPreAtivacao(2, 1) ' Data Prevista para Realização
'        sResponsavel = paDadosPreAtivacao(3, 1) ' Responsável
'        dDataExecucaoTeste = paDadosPreAtivacao(4, 1) ' Data de Execução do Teste
        
        sTipoGravacao = "A"
        
    'End If
       
    MostrarDados = True
    
    Exit Function
        
ErrorHandler:

    MsgBox ("Erro (MostraDados)")
    
    Exit Function

End Function
'
Public Function Alterar( _
        ByVal piID As Integer _
        ) As Boolean

    ' Inicializa o retorno do método
    Alterar = False
    
    ' Executa método de Alteração da Classe de controle Secundária
    Alterar = objOC_OperadoraSecundario.Carrega_Dados_Alteracao(piID)
    
    ' Chamar a visao do usuário ( FORM ) secundaria para Alteração/Cadastramento
    Call frmGenManterOperadorasSecundario.Alterar(objOC_OperadoraSecundario)
    
    ' Sai informando TRUE
    Alterar = True
    
End Function

Public Function GravarDados( _
        ByVal psTipoGravacao As String, _
        ByVal piID_Servico As Integer _
        ) As Boolean

On Error GoTo ErrorHandler:

    'Dim objPreAtivacao As WID.clswidPreAtivacao
    Dim bRetorno As Boolean
    
    GravarDados = False
    
    ' Instância a classe de pré ativação
    'Set objPreAtivacao = New WID.clswidPreAtivacao
    
    ' Informar o nome do banco desejado
    'objPreAtivacao.InformaBanco frmgenMDI_SGAPlus.objBanco
            
    ' Excutar o método
'    bRetorno = objPreAtivacao.GravarDados(dDataPedidoTeste, _
'                                          dDataPrevistaRealizacao, _
'                                          sResponsavel, _
'                                          dDataExecucaoTeste, _
'                                          psTipoGravacao, _
'                                          piID_Servico)
    
    ' Se ocorreu algum erro na DLL exibi-lo
    If Not bRetorno Then
    
        GoTo ErrorHandler
        
    End If
    
    'Set objPreAtivacao = Nothing
    
    GravarDados = True
    
    Exit Function
        
ErrorHandler:

    MsgBox ("Erro (GravarDados)")
    
    Exit Function

End Function

Public Function ExcluirDados(ByVal piID_Servico As Integer) As Boolean

On Error GoTo ErrorHandler:

    'Dim objPreAtivacao As WID.clswidPreAtivacao
    Dim bRetorno As Boolean
    
    ExcluirDados = False
    
    ' Instância a classe de pré ativação
    'Set objPreAtivacao = New WID.clswidPreAtivacao
    
    ' Informar o nome do banco desejado
    'objPreAtivacao.InformaBanco frmgenMDI_SGAPlus.objBanco
            
    ' Excutar o método
    'bRetorno = objPreAtivacao.ExcluirDados(piID_Servico)
    
    ' Se ocorreu algum erro na DLL exibi-lo
    If Not bRetorno Then
    
        GoTo ErrorHandler
        
    End If
    
    'Set objPreAtivacao = Nothing
    
    ExcluirDados = True

    Exit Function
        
ErrorHandler:

    MsgBox ("Erro (ExcluirDados)")
    
    Exit Function

End Function

Private Sub Class_Terminate()
    ' Limpa as estruturas de classes criadas
    Set objOC_OperadoraSecundario = Nothing
    Set objGenOperadora = Nothing
    Set objMsg = Nothing
End Sub

