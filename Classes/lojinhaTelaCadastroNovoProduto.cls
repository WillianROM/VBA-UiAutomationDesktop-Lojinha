VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "lojinhaTelaCadastroNovoProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    Dim oAutomation                 As New CUIAutomation
    Dim janela                      As UIAutomationClient.IUIAutomationElement
    Dim tPageControl                As UIAutomationClient.IUIAutomationElement
    Dim novoProduto                 As UIAutomationClient.IUIAutomationElement
    Dim tPainelNovoProduto          As UIAutomationClient.IUIAutomationElement
    Dim oPattern                    As UIAutomationClient.IUIAutomationLegacyIAccessiblePattern

Public Function localizarAplicativoDaLojinha()
    Set janela = WalkEnabledElements(oAutomation, oAutomation.GetRootElement, "TfrmPrincipal")
End Function

Public Function localizarOGrupoTPageControl()
    Set tPageControl = janela.FindFirst(TreeScope_Children, PropCondition(oAutomation, "TPageControl", "ClsName"))
End Function

Public Function localizarOGrupoNovoProduto()
    Dim i                           As Integer
    Dim TPainels                    As UIAutomationClient.IUIAutomationElementArray
    
    Set novoProduto = tPageControl.FindFirst(TreeScope_Children, PropCondition(oAutomation, "NOVO PRODUTO", "Name"))
    Set TPainels = novoProduto.FindAll(TreeScope_Children, PropCondition(oAutomation, "TPanel", "ClsName"))
    
    For i = 0 To TPainels.Length - 1
        If TPainels.GetElement(i).FindAll(TreeScope_Children, PropCondition(oAutomation, "TButton", "ClsName")).Length > 0 Then
            Set tPainelNovoProduto = TPainels.GetElement(i)
            Exit For
        End If
    Next i

End Function

Public Function informarNomeDoProduto(ByVal nomeDoProduto As String)
    Dim camposEditar                As UIAutomationClient.IUIAutomationElementArray
    
    Set camposEditar = tPainelNovoProduto.FindAll(TreeScope_Children, PropCondition(oAutomation, "TEdit", "ClsName"))

    Set oPattern = camposEditar.GetElement(2).GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
    oPattern.SetValue (nomeDoProduto)
    
End Function

Public Function informarValorDoProduto(ByVal valorDoProduto As String)
    Dim camposEditar                As UIAutomationClient.IUIAutomationElementArray
    
    Set camposEditar = tPainelNovoProduto.FindAll(TreeScope_Children, PropCondition(oAutomation, "TEdit", "ClsName"))

    Set oPattern = camposEditar.GetElement(1).GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
    oPattern.SetValue (valorDoProduto)
    
End Function

Public Function informarCoresDoProduto(ByVal CoresDoProduto As String)
    Dim camposEditar                As UIAutomationClient.IUIAutomationElementArray
    
    Set camposEditar = tPainelNovoProduto.FindAll(TreeScope_Children, PropCondition(oAutomation, "TEdit", "ClsName"))

    Set oPattern = camposEditar.GetElement(0).GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
    oPattern.SetValue (CoresDoProduto)
    
End Function

Public Function clicarNoBotaoSalvar()
    Dim BtnSalvar                   As UIAutomationClient.IUIAutomationElement
    Dim BtnSalvarClick              As UIAutomationClient.IUIAutomationInvokePattern
    
    Set BtnSalvar = tPainelNovoProduto.FindFirst(TreeScope_Children, PropCondition(oAutomation, "SALVAR", "Name"))
    
    Set BtnSalvarClick = BtnSalvar.GetCurrentPattern(UIAutomationClient.UIA_InvokePatternId)
    BtnSalvarClick.Invoke
    
End Function
