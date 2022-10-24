VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "lojinhaTelaDeProdutos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    Dim oAutomation                 As New CUIAutomation
    Dim janela                      As UIAutomationClient.IUIAutomationElement
    Dim tPageControl                As UIAutomationClient.IUIAutomationElement
    Dim produtos                    As UIAutomationClient.IUIAutomationElement

Public Function localizarAplicativoDaLojinha()
    Set janela = WalkEnabledElements(oAutomation, oAutomation.GetRootElement, "TfrmPrincipal")
End Function

Public Function localizarOGrupoTPageControl()
    Set tPageControl = janela.FindFirst(TreeScope_Children, PropCondition(oAutomation, "TPageControl", "ClsName"))
End Function

Public Function localizarOGrupoProdutos()
    Set produtos = tPageControl.FindFirst(TreeScope_Children, PropCondition(oAutomation, "PRODUTOS", "Name"))
End Function

Public Function clicarNoBotaoAdicionarProduto()
    Dim i                           As Integer
    Dim TPainels                    As UIAutomationClient.IUIAutomationElementArray
    Dim BtnAdicionarProduto         As UIAutomationClient.IUIAutomationElement
    Dim BtnAdicionarProdutoClick    As UIAutomationClient.IUIAutomationInvokePattern

    Set TPainels = produtos.FindAll(TreeScope_Children, PropCondition(oAutomation, "TPanel", "ClsName"))
    
    For i = 0 To TPainels.Length - 1
    
        If TPainels.GetElement(i).FindAll(TreeScope_Children, PropCondition(oAutomation, "TButton", "ClsName")).Length > 0 Then
            'Localizar o botão Gravar
            Set BtnAdicionarProduto = TPainels.GetElement(i).FindFirst(TreeScope_Children, PropCondition(oAutomation, "Adicionar Produto", "Name"))

            'Clicar no botão Gravar
            Set BtnAdicionarProdutoClick = BtnAdicionarProduto.GetCurrentPattern(UIAutomationClient.UIA_InvokePatternId)
            BtnAdicionarProdutoClick.Invoke
            
            Exit For
        End If
    Next i
End Function
