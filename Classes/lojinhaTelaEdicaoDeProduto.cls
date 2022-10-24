VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "lojinhaTelaEdicaoDeProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    Dim oAutomation                 As New CUIAutomation
    Dim janela                      As UIAutomationClient.IUIAutomationElement
    Dim tPageControl                As UIAutomationClient.IUIAutomationElement
    Dim editarProduto               As UIAutomationClient.IUIAutomationElement
    Dim tPainelEditarProduto        As UIAutomationClient.IUIAutomationElement

Public Function localizarAplicativoDaLojinha()
    Set janela = WalkEnabledElements(oAutomation, oAutomation.GetRootElement, "TfrmPrincipal")
End Function

Public Function localizarOGrupoTPageControl()
    Set tPageControl = janela.FindFirst(TreeScope_Children, PropCondition(oAutomation, "TPageControl", "ClsName"))
End Function

Public Function localizarOGrupoEditarProduto()
    Dim i                           As Integer
    Dim TPainels                    As UIAutomationClient.IUIAutomationElementArray
    
    Set editarProduto = tPageControl.FindFirst(TreeScope_Children, PropCondition(oAutomation, "EDITAR PRODUTO", "Name"))
    Set TPainels = editarProduto.FindAll(TreeScope_Children, PropCondition(oAutomation, "TPanel", "ClsName"))
    
    For i = 0 To TPainels.Length - 1
        If TPainels.GetElement(i).FindAll(TreeScope_Children, PropCondition(oAutomation, "TButton", "ClsName")).Length > 0 Then
            Set tPainelEditarProduto = TPainels.GetElement(i)
            Exit For
        End If
    Next i

End Function

Public Function clicarNoBotaoAdicionarComponente()
    Dim BtnAdicionarComponente      As UIAutomationClient.IUIAutomationElement
    Dim BtnAdicionarComponenteClick As UIAutomationClient.IUIAutomationInvokePattern
    
    Set BtnAdicionarComponente = tPainelEditarProduto.FindFirst(TreeScope_Children, PropCondition(oAutomation, "ADICIONAR COMPONENTE", "Name"))
    
    Set BtnAdicionarComponenteClick = BtnAdicionarComponente.GetCurrentPattern(UIAutomationClient.UIA_InvokePatternId)
    BtnAdicionarComponenteClick.Invoke
    
End Function
