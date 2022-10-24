Attribute VB_Name = "Mod02_abrirAplicativo"
Option Explicit

Sub abrirAplicativo()
    Dim objShell    As Object
    Dim caminho     As String

    Set objShell = CreateObject("Shell.Application")
 
    caminho = ThisWorkbook.Path & "\Lojinha.exe"
    
    objShell.Open (caminho)
    
End Sub
