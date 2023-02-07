VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsgenOC_AlteraCentroFuncional"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim objMensagem      As New clsgenMensagem
Dim lServicoID       As Long



Public Function CarregarForm() As Boolean

    Dim objOC_SelecionarOTS      As New clsgenOC_SelecionarOTS
    Dim objForm                  As New frmgenAlteraCentroFuncional
    
    On Error GoTo ErrorHandler
    
    Load objForm
    
    Call objOC_SelecionarOTS.ExibirListaCentroFuncional(objForm.CboCentroFuncionalAtual)
    
    Call objOC_SelecionarOTS.ExibirListaCentroFuncional(objForm.cboCentroFuncionalAlterado, _
                                                        GbCodigoFaseAnalisarOE)
    
    objForm.informa_OC Me
    
    objForm.Show vbModal
        
Finalizar:

    Set objOC_SelecionarOTS = Nothing
    
    Exit Function
    
ErrorHandler:

    If Err.Number <> 0 Then
       
       objMensagem.TipoMsg = "E"
       objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                   "Houve um erro de execução no Método: " & vbCrLf & vbCrLf & "CarregarForm " & _
                   "da classe clsblfOC_AlteraCentroFuncional"
    End If
       
    objMensagem.ExibeMensagem
       
    GoTo Finalizar

End Function

Private Sub Class_Terminate()

    Set objMensagem = Nothing

End Sub

Public Function VerificaOTS(objForm As frmgenAlteraCentroFuncional) As Boolean

    
    Dim objdadosServico    As New clsGenServico
    Dim objDadosFase       As New clsgenFase
    
    Dim lIdCentroFuncional As Long
    Dim lServicoID         As Long
    Dim lSequenciaAcao     As Long
    
On Error GoTo ErrorHandler

    objdadosServico.InformaMensagem objMensagem
    objdadosServico.InformaBanco frmgenMDI_SGAPlus.ObjBanco
    
    lIdCentroFuncional = objForm.CboCentroFuncionalAtual.ItemData( _
                         objForm.CboCentroFuncionalAtual.ListIndex)
    
    Servico_ID = objdadosServico.Obtem_IDServico(lIdCentroFuncional, _
                                            CLng(objForm.SpdNumeroOTS.Text), _
                                            CLng(objForm.SpdAnoOTS.Text))
    
    If objMensagem.TipoMsg <> "" Then
    
       GoTo ErrorHandler
       
    End If
        
    If Servico_ID = 0 Then
    
        objMensagem.TipoMsg = "A"
        objMensagem.Descricao = "Serviço informado não encontrado! "
        objMensagem.ExibeMensagem
        
        objForm.CboCentroFuncionalAtual.SetFocus
        
        GoTo Finalizar
    
    End If
        
    objDadosFase.InformaBanco frmgenMDI_SGAPlus.ObjBanco
    objDadosFase.InformaMensagem objMensagem
    
    lSequenciaAcao = objDadosFase.ObtemSequenciadaFase(Servico_ID, GbCodigoFaseAlocacao)
    
    If objDadosFase.ObtemSequenciaUltimaFaseRealizada(Servico_ID) < lSequenciaAcao Then
        
        objForm.cmdSalvar.Enabled = True
        objForm.cboCentroFuncionalAlterado.Enabled = True
        objForm.cboCentroFuncionalAlterado.SetFocus
        
    Else
    
        objMensagem.TipoMsg = "A"
        objMensagem.Descricao = "O Centro Funcional deste Serviço não poderá ser alterado!"
        objMensagem.ExibeMensagem
        
    End If
    
Finalizar:

    Set objdadosServico = Nothing
    Set objDadosFase = Nothing
    
    Exit Function

ErrorHandler:

    If Err.Number <> 0 Then
       
       objMensagem.TipoMsg = "E"
       objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                   "Houve um erro de execução no Método: " & vbCrLf & vbCrLf & "VerificarOTS " & _
                   "da classe clsblfOC_AlteraCentroFuncional"
    End If
       
    objMensagem.ExibeMensagem
       
    GoTo Finalizar


End Function


Public Function AtualizaCentroFuncional(objForm As frmgenAlteraCentroFuncional) As Boolean

    Dim objdadosServico    As New clsGenServico
    Dim objDadosFase       As New clsgenFase
    
    Dim lIdCentroFuncional As Long
    Dim lServicoID         As Long
    Dim lSequenciaAcao     As Long
    Dim lIDFase            As Long
    Dim bRetorno           As Boolean
    
On Error GoTo ErrorHandler

    objdadosServico.InformaMensagem objMensagem
    objdadosServico.InformaBanco frmgenMDI_SGAPlus.ObjBanco
    
    lIdCentroFuncional = objForm.cboCentroFuncionalAlterado.ItemData( _
                         objForm.cboCentroFuncionalAlterado.ListIndex)
    
    If objMensagem.TipoMsg <> "" Then
    
       GoTo ErrorHandler
       
    End If
    
    objDadosFase.InformaBanco frmgenMDI_SGAPlus.ObjBanco
    objDadosFase.InformaMensagem objMensagem
    
    lIDFase = objDadosFase.ObtemIDFase(Servico_ID, GbCodigoFaseAlocacao)

    'Atualiza o Centro funcional selecionado para a fase Centro Local
    bRetorno = objDadosFase.GravarIDCentroFuncional(lIDFase, lIdCentroFuncional)
    
    If objMensagem.TipoMsg <> "" Then
    
       GoTo ErrorHandler
       
    End If

Finalizar:

    Set objdadosServico = Nothing
    Set objDadosFase = Nothing
    
    Exit Function

ErrorHandler:

    If Err.Number <> 0 Then
       
       objMensagem.TipoMsg = "E"
       objMensagem.Descricao = "Erro: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
                   "Houve um erro de execução no Método: " & vbCrLf & vbCrLf & "AtualizaCentroFuncional " & _
                   "da classe clsblfOC_AlteraCentroFuncional"
    End If
       
    objMensagem.ExibeMensagem
       
    GoTo Finalizar

End Function

Public Property Get Servico_ID() As Long

Servico_ID = lServicoID

End Property

Public Property Let Servico_ID(ByVal lNewValue As Long)

lServicoID = lNewValue

End Property
