VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "lojinhaAdicionarComponente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
    Dim oAutomation                 As New CUIAutomation
    Dim janela                      As UIAutomationClient.IUIAutomationElement
    Dim oPattern                    As UIAutomationClient.IUIAutomationLegacyIAccessiblePattern
    Dim tPageControl                As UIAutomationClient.IUIAutomationElement
    Dim tPainelEditarProduto        As UIAutomationClient.IUIAutomationElement
    
Public Function localizarAplicativoDaLojinha()
    Set janela = WalkEnabledElements(oAutomation, oAutomation.GetRootElement, "TfrmPrincipal")
End Function

Public Function localizarOGrupoTPageControl()
    Set tPageControl = janela.FindFirst(TreeScope_Children, PropCondition(oAutomation, "TPageControl", "ClsName"))
End Function

Public Function localizarOGrupoAdicionarComponenteAoProduto()
    Dim i                           As Integer
    Dim TPainels                    As UIAutomationClient.IUIAutomationElementArray
    Dim editarProduto               As UIAutomationClient.IUIAutomationElement
 
    
    Set editarProduto = tPageControl.FindFirst(TreeScope_Children, PropCondition(oAutomation, "EDITAR PRODUTO", "Name"))
    Set TPainels = editarProduto.FindAll(TreeScope_Children, PropCondition(oAutomation, "TPanel", "ClsName"))
    
    For i = 0 To TPainels.Length - 1
    
    Debug.Print TPainels.GetElement(i).FindAll(TreeScope_Children, PropCondition(oAutomation, "TButton", "ClsName")).Length
        If TPainels.GetElement(i).FindAll(TreeScope_Children, PropCondition(oAutomation, "TButton", "ClsName")).Length = 2 Then
            Set tPainelEditarProduto = TPainels.GetElement(i)
            Exit For
        End If
    Next i

End Function


Public Function informarNovoComponente(ByVal novoComponente As String)
    Dim camposEditar                As UIAutomationClient.IUIAutomationElementArray
    
    Set camposEditar = tPainelEditarProduto.FindAll(TreeScope_Children, PropCondition(oAutomation, "TEdit", "ClsName"))

    Set oPattern = camposEditar.GetElement(1).GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
    oPattern.SetValue (novoComponente)
    
End Function

Public Function informarQuantidadeNovoComponente(ByVal quantidadeComponente As Integer)
    Dim camposEditar                As UIAutomationClient.IUIAutomationElementArray
    
    Set camposEditar = tPainelEditarProduto.FindAll(TreeScope_Children, PropCondition(oAutomation, "TEdit", "ClsName"))

    Set oPattern = camposEditar.GetElement(0).GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
    oPattern.SetValue (quantidadeComponente)
    
End Function

Public Function clicarNoBotaoSalvarComponente()
    Dim BtnSalvarComponente         As UIAutomationClient.IUIAutomationElement
    Dim BtnSalvarComponenteClick    As UIAutomationClient.IUIAutomationInvokePattern
    
    Set BtnSalvarComponente = tPainelEditarProduto.FindFirst(TreeScope_Children, PropCondition(oAutomation, "SALVAR COMPONENTE", "Name"))
    
    Debug.Print BtnSalvarComponente.CurrentName
    
    Set BtnSalvarComponenteClick = BtnSalvarComponente.GetCurrentPattern(UIAutomationClient.UIA_InvokePatternId)
    BtnSalvarComponenteClick.Invoke
    
End Function
