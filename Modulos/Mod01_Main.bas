Attribute VB_Name = "Mod01_Main"
Option Explicit

Sub Main()

    Dim telaDeLogin         As New lojinhaTelaDeLogin
    Dim telaProdutos        As New lojinhaTelaDeProdutos
    Dim telaNovoProduto     As New lojinhaTelaCadastroNovoProduto
    Dim telaEditarProduto   As New lojinhaTelaEdicaoDeProduto
    Dim telaNovoComponente  As New lojinhaAdicionarComponente
    Dim telaInformation     As New lojinhaCaixaDialogoInformation
    
    Call abrirAplicativo

    Application.Wait (Now + TimeValue("0:00:02"))

    With telaDeLogin
        .localizarAplicativoDaLojinha
        .localizarOGrupoTPageControl
        .localizarOGrupoLogin
        .informarOUsuario ("admin")
        .informarASenha ("admin")
        .clicarNoBotaoEntrar
    End With
    
        Application.Wait (Now + TimeValue("0:00:02"))
        
    With telaProdutos
        .localizarAplicativoDaLojinha
        .localizarOGrupoTPageControl
        .localizarOGrupoProdutos
        .clicarNoBotaoAdicionarProduto
    End With
    
        Application.Wait (Now + TimeValue("0:00:02"))
        
    With telaNovoProduto
        .localizarAplicativoDaLojinha
        .localizarOGrupoTPageControl
        .localizarOGrupoNovoProduto
        .informarNomeDoProduto ("Camiseta")
        .informarValorDoProduto ("5050")
        .informarCoresDoProduto ("preto, amarelo")
        .clicarNoBotaoSalvar
    End With

    Application.Wait (Now + TimeValue("0:00:02"))
    
    With telaInformation
        .localizarAplicativoDaLojinha
        .localizarACaixaDeDialogoInformation
        .clicarNoBotaoOkDaCaixaDeDialogoInformation
    End With
    
    Application.Wait (Now + TimeValue("0:00:02"))

    With telaEditarProduto
        .localizarAplicativoDaLojinha
        .localizarOGrupoTPageControl
        .localizarOGrupoEditarProduto
        .clicarNoBotaoAdicionarComponente
    End With
    
    Application.Wait (Now + TimeValue("0:00:02"))
    
    With telaNovoComponente
        .localizarAplicativoDaLojinha
        .localizarOGrupoTPageControl
        .localizarOGrupoAdicionarComponenteAoProduto
        .informarNovoComponente ("Broche")
        .informarQuantidadeNovoComponente (2)
        .clicarNoBotaoSalvarComponente
    End With
    
    Application.Wait (Now + TimeValue("0:00:02"))
    
    With telaInformation
        .localizarAplicativoDaLojinha
        .localizarAplicativoDaLojinha
        .localizarACaixaDeDialogoInformation
        .clicarNoBotaoOkDaCaixaDeDialogoInformation
    End With

End Sub


