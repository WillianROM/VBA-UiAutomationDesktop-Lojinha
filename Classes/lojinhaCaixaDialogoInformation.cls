VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "lojinhaCaixaDialogoInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    Dim oAutomation                 As New CUIAutomation
    Dim janela                      As UIAutomationClient.IUIAutomationElement
    Dim Information                 As UIAutomationClient.IUIAutomationElement
    
Public Function localizarAplicativoDaLojinha()
    Set janela = WalkEnabledElements(oAutomation, oAutomation.GetRootElement, "TfrmPrincipal")
End Function

Public Function localizarACaixaDeDialogoInformation()
    Set Information = janela.FindFirst(TreeScope_Children, PropCondition(oAutomation, "Information", "Name"))
End Function

Public Function clicarNoBotaoOkDaCaixaDeDialogoInformation()
    Dim BtnOk                       As UIAutomationClient.IUIAutomationElement
    Dim BtnOkClick                  As UIAutomationClient.IUIAutomationInvokePattern
    
    Set BtnOk = Information.FindFirst(TreeScope_Children, PropCondition(oAutomation, "OK", "Name"))
    
    Set BtnOkClick = BtnOk.GetCurrentPattern(UIAutomationClient.UIA_InvokePatternId)
    BtnOkClick.Invoke
    
End Function
