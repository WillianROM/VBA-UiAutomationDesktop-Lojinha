VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "lojinhaTelaDeLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    Dim oAutomation                 As New CUIAutomation
    Dim janela                      As UIAutomationClient.IUIAutomationElement
    Dim tPageControl                As UIAutomationClient.IUIAutomationElement
    Dim login                       As UIAutomationClient.IUIAutomationElement
    Dim oPattern                    As UIAutomationClient.IUIAutomationLegacyIAccessiblePattern

Public Function localizarAplicativoDaLojinha()
    Set janela = WalkEnabledElements(oAutomation, oAutomation.GetRootElement, "TfrmPrincipal")
End Function

Public Function localizarOGrupoTPageControl()
    Set tPageControl = janela.FindFirst(TreeScope_Children, PropCondition(oAutomation, "TPageControl", "ClsName"))
End Function

Public Function localizarOGrupoLogin()
    Set login = tPageControl.FindFirst(TreeScope_Children, PropCondition(oAutomation, "LOGIN", "Name"))
End Function

Public Function informarOUsuario(ByVal usuario As String)
    Dim i                           As Integer
    Dim camposEditar                As UIAutomationClient.IUIAutomationElementArray
    
    Set camposEditar = login.FindAll(TreeScope_Children, PropCondition(oAutomation, "TEdit", "ClsName"))

    For i = 0 To camposEditar.Length - 1
    
        If camposEditar.GetElement(i).CurrentHelpText = "Usuário" Then
    
            Set oPattern = camposEditar.GetElement(i).GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
            oPattern.SetValue (usuario)
            
            Exit For
        End If
    Next i
End Function


Public Function informarASenha(ByVal senha As String)
    Dim i                           As Integer
    Dim camposEditar                As UIAutomationClient.IUIAutomationElementArray
    
    Set camposEditar = login.FindAll(TreeScope_Children, PropCondition(oAutomation, "TEdit", "ClsName"))

    For i = 0 To camposEditar.Length - 1
    
        If camposEditar.GetElement(i).CurrentHelpText = "Senha" Then
    
            Set oPattern = camposEditar.GetElement(i).GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
            oPattern.SetValue (senha)
            
            Exit For
        End If
    Next i
End Function

Public Function clicarNoBotaoEntrar()
    Dim BtnEntrar                   As UIAutomationClient.IUIAutomationElement
    Dim BtnEntrarClick              As UIAutomationClient.IUIAutomationInvokePattern
    
    Set BtnEntrar = login.FindFirst(TreeScope_Children, PropCondition(oAutomation, "Entrar", "Name"))
    
    Set BtnEntrarClick = BtnEntrar.GetCurrentPattern(UIAutomationClient.UIA_InvokePatternId)
    BtnEntrarClick.Invoke
    
End Function
