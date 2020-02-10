using NUnit.Framework;
using System;
using System.Text;
using System.Threading;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using Starline;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.IE;

namespace FluxoVendaCartoes
{
    [TestFixture]
    public class VendaCartoes
    {
        private const string login = "TI_SANDRA";
        private const string senha = "Teste@1234";
        private IWebDriver driver;
        private StringBuilder verificationErrors;
        private string baseURL;
        private bool acceptNextAlert = true;
        public bool OSLinux = false;
        public string PathChromeDriver = "";
        public ProcessTest processTest;


        [SetUp]
        public void SetupTest()

        {

            processTest = new ProcessTest();
            var options = new InternetExplorerOptions();

            var os = Environment.OSVersion;
            OSLinux = os.Platform == PlatformID.Unix;
            InternetExplorerDriverService service = null;
            if (OSLinux)
            {
                PathChromeDriver = "/usr/bin/";
                service = InternetExplorerDriverService.CreateDefaultService(PathChromeDriver, "IEDriverServer");
            }
            else
            {
                //PathChromeDriver = @"C:\Mp4-Marisa\CCM_DotNet\CCMNetCore\CCMNetCore02\CCMNetCore02\bin\Debug\netcoreapp2.1";
                //service = ChromeDriverService.CreateDefaultService(PathChromeDriver, "chromedriver.exe");


                PathChromeDriver = @"C:\Projetos\CCMNetCore02\CCMNetCore02\bin\Debug\netcoreapp2.1";
                service = InternetExplorerDriverService.CreateDefaultService(PathChromeDriver, "IEDriverServer.exe");
                //IWebDriver driver = new InternetExplorerDriver(@"C:\Projetos\CCMNetCore02\CCMNetCore02\bin\Debug\netcoreapp2.1\IEDriverServer.exe");
                //driver.Navigate().GoToUrl("http://www.google.com");
            }


            //processTest.CustomerName = "marisa"; //CustomerName evitar Acentuacao e espa�o
            //processTest.SuiteName = "ccm"; //CustomerName evitar Acentuacao e espa�o
            //processTest.ScenarioName = "cadastroaposentado";
            //processTest.ReportID = 271;

            //options.AddArgument("--headless");   
            
            //options.AddArgument("--no-sandbox");
            //options.AddArgument("--disable-dev-shm-usage");
            //options.AddArgument("--touch-events=enabled");
            //options.AddArgument("start-maximized");
            //options.EnableMobileEmulation("iPhone 5/SE");
//            driver = new ChromeDriver(service, options);
            driver = new InternetExplorerDriver(service, options);

            //baseURL = "https://ccmhomolog.marisa.com.br/psfsecurity/paginas/security/login/tela/login.html";
            baseURL = "https://pagseguro.uol.com.br/";
            verificationErrors = new StringBuilder();


            //// Executar testes em background
            //var options = new ChromeOptions();
            //options.AddArguments(new List<string>() { "headless" });
            //options.AddArgument("--touch-events=enabled");
            //options.AddArgument("start-maximized");
            //options.EnableMobileEmulation("iPhone 5/SE");
            //options.AddArgument("--disable-plugins");
            //var chromeDriverService = ChromeDriverService.CreateDefaultService();
            //driver = new ChromeDriver(chromeDriverService, options);


            //// Executar teste com browser aberto Chrome
            //ChromeOptions options = new ChromeOptions();
            //options.AddArgument("--touch-events=enabled");
            //options.AddArgument("--no-sandbox");
            //options.AddArgument("--disable-dev-shm-usage");
            //options.AddArgument("start-maximized");
            //options.EnableMobileEmulation("iPhone 5/SE");
            //driver = new ChromeDriver(@"C:\MARISA\CCM\NUnit_NetCore\NUnit_NetCore\bin\Debug\netcoreapp2.1", options);
            //driver.Manage().Window.Maximize();
            //baseURL = "https://ccmhomolog.marisa.com.br/psfsecurity/paginas/security/login/tela/login.html";
            //verificationErrors = new StringBuilder();



        }

        public void wait(By elemento)
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(60));
            wait.Until(ExpectedConditions.ElementToBeClickable(elemento));
        }



        //public void Login()
        //{ //Validar Se��o de Login



        //    //Efetuar Login
        //    driver.FindElement(By.Id("login_username")).Clear();
        //    driver.FindElement(By.Id("login_username")).SendKeys(login);
        //    driver.FindElement(By.Id("login_password")).Clear();
        //    driver.FindElement(By.Id("login_password")).SendKeys(senha);
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.XPath("//button[@type='submit']")).Click();
        //    Thread.Sleep(2000);

        //    wait(By.Id("text-login-msg"));

        //}


        [TearDown]
        public void TeardownTest()
        {

            try
            {
                driver.Quit();
            }
            catch
            {

            }
        }

        [Test]
        public void CT001_TST01()
        {
            //Actions touchActions = new Actions(driver);

            //processTest.TestName = System.Reflection.MethodBase.GetCurrentMethod().Name;
            //processTest.StepNumber = 1;
            //processTest.StepTurn = 1;


            //Acessar P�gina do CCM
            //driver.Manage().Window.Minimize();
            driver.Navigate().GoToUrl(baseURL);
            //processTest.PrintPageComSelenium(driver, false);

            //Validar P�gina Inicial
            Assert.AreEqual("PagSeguro: Venda muito pelas máquinas de cartão ou pela internet", driver.Title);

        }
        //[Test]
        //public void CT001_TST02()
        //{
        //    Actions touchActions = new Actions(driver);

        //    processTest.TestName = System.Reflection.MethodBase.GetCurrentMethod().Name;
        //    processTest.StepNumber = 1;
        //    processTest.StepTurn = 2;


        //    //Acessar P�gina do CCM
        //    driver.Navigate().GoToUrl(baseURL);
        //    driver.FindElement(By.XPath("//*[@id='ps']/div[1]/section/div/div[4]/p/a[2]")).Click();
        //    processTest.PrintPageComSelenium(driver, false);

        //    //Validar P�gina Inicial
        //    Assert.AreEqual("Prévia da", driver.FindElement(By.XPath("//*[@id='localnav']/div/div[2]/div[1]/span")));

        //}

        //[Test]
        //public void CT001_TST03()
        //{
        //    Actions touchActions = new Actions(driver);

        //    processTest.TestName = System.Reflection.MethodBase.GetCurrentMethod().Name;
        //    processTest.StepNumber = 1;
        //    processTest.StepTurn = 3;


        //    //Acessar P�gina do CCM
        //    driver.Navigate().GoToUrl(baseURL);
        //    driver.FindElement(By.XPath("//*[@id='ps']/div[1]/section/div/div[4]/p/a[2]")).Click();
        //    processTest.PrintPageComSelenium(driver, false);

        //    //Validar P�gina Inicial
        //    Assert.AreEqual("Préviaa", driver.FindElement(By.XPath("//*[@id='localnav']/div/div[2]/div[1]/span")));

        //}

        //[Test]
        //public void CT001_Cadastro_de_Clientes_Aposentados()
        //{

        //    Actions touchActions = new Actions(driver);

        //    processTest.TestName = System.Reflection.MethodBase.GetCurrentMethod().Name;
        //    processTest.StepNumber = 1;
        //    processTest.StepTurn = 1;


        //    //Acessar P�gina do CCM
        //    driver.Navigate().GoToUrl(baseURL);
        //    processTest.PrintPageComSelenium(driver, false);

        //    //Validar P�gina Inicial
        //    Assert.AreEqual("PagSeguro: Venda muito pelas máquinas de cartão ou pela internet", driver.Title);



        //Testando 

        //wait(By.XPath("//button[@type='submit']"));
        //wait(By.CssSelector("img.png_bg"));
        ////print.PrintScreen();
        //processTest.StepTurn = 2;
        //processTest.PrintPageComSelenium(driver, false);

        //try
        //{
        //    NUnit.Framework.Assert.AreEqual("Login", driver.FindElement(By.CssSelector("strong")).Text);
        //}
        //catch (AssertionException e)
        //{
        //    verificationErrors.Append(e.Message);
        //}
        //NUnit.Framework.Assert.IsTrue(IsElementPresent(By.Id("login_username")));
        //NUnit.Framework.Assert.IsTrue(IsElementPresent(By.Id("login_password")));
        //NUnit.Framework.Assert.IsTrue(IsElementPresent(By.XPath("//button[@type='submit']")));



        ////Validar Tela de Login
        //NUnit.Framework.Assert.AreEqual("PSF Security", driver.Title);
        //NUnit.Framework.Assert.IsTrue(IsElementPresent(By.CssSelector("img.png_bg")));
        //try
        //{
        //    NUnit.Framework.Assert.AreEqual("Portal Lojas", driver.FindElement(By.Id("tituloPagina")).Text);
        //}
        //catch (AssertionException e)
        //{
        //    verificationErrors.Append(e.Message);
        //}

        //processTest.StepTurn = 3;
        //processTest.PrintPageComSelenium(driver, false);

        ////Realizar login
        //Login();


        //wait(By.Id("btnSearch"));
        //processTest.StepTurn = 4;
        //processTest.PrintPageComSelenium(driver, false);


        //////Alterar Filial
        //Thread.Sleep(2000);

        //driver.FindElement(By.Id("btnYesOk")).Click();
        //Thread.Sleep(2000);

        //driver.FindElement(By.Id("labelFilial")).Click();
        //Thread.Sleep(2000);

        //new SelectElement(driver.FindElement(By.Id("changeFilial"))).SelectByValue("54");
        //driver.FindElement(By.Id("changeFilial")).SendKeys(OpenQA.Selenium.Keys.Tab);
        //driver.FindElement(By.XPath("//div[2]/div/div/div[3]/button")).Click();

        //Thread.Sleep(5000);

        ////Validar P�gina Inicial
        //Assert.AreEqual("PSF Security", driver.Title);

        //try
        //{
        //    Assert.AreEqual("Buscar_pornCPFnCARTAO", processTest.Slugify(driver.FindElement(By.Id("tipoSearch")).Text));
        //}
        //catch (AssertionException e)
        //{
        //    verificationErrors.Append(e.Message);
        //}
        //Assert.IsTrue(IsElementPresent(By.Id("cpfSearch")));
        //Assert.IsTrue(IsElementPresent(By.Id("btnSearch")));
        //try
        //{
        //    Assert.AreEqual("Concessao do Cartao", driver.FindElement(By.Id("menuCollapse3162")).Text);
        //}
        //catch (AssertionException e)
        //{
        //    verificationErrors.Append(e.Message);
        //}
        //try
        //{
        //    Assert.AreEqual("Atendimento", driver.FindElement(By.Id("menuCollapse3164")).Text);
        //}
        //catch (AssertionException e)
        //{
        //    verificationErrors.Append(e.Message);
        //}


        ////Cenario 1 - Cadastro de Clientes

        ////Consultar Cliente
        //driver.FindElement(By.Id("cpfSearch")).Clear();
        //driver.FindElement(By.Id("cpfSearch")).SendKeys("5339718379");
        //Thread.Sleep(2000);
        //processTest.StepTurn = 5;
        //processTest.PrintPageComSelenium(driver, false);
        //driver.FindElement(By.Id("btnSearch")).Click();
        //Thread.Sleep(2000);


        //wait(By.Id("iniciarCadastro"));


        ////Validar P�gina Cadastro de Clientes


        //Assert.IsTrue(IsElementPresent(By.Id("formPreCadastro")));

        //driver.FindElement(By.Id("dataNascCad")).Click();
        //driver.FindElement(By.Id("dataNascCad")).SendKeys("10/06/1992");
        //driver.FindElement(By.Id("iniciarCadastro")).Click();


        ////Validar Mensagem de Retorno CPF

        //Thread.Sleep(8000);


        //wait(By.Id("avancarDadosPessoais"));

        ////Cenario 2  -- Dados pessoais


        ////Validar P�gina Cadastro de Clientes


        //Assert.IsTrue(IsElementPresent(By.Id("formDadosPessoais")));


        //////Nome da M�e
        ////Thread.Sleep(2000);
        ////driver.FindElement(By.Id("nomeMaeDadosPessoais")).Click();
        ////driver.FindElement(By.Id("nomeMaeDadosPessoais")).Clear();
        ////Thread.Sleep(2000);
        ////driver.FindElement(By.Id("nomeMaeDadosPessoais")).SendKeys("Maria Aparecida");


        ////Estado Civil
        //Thread.Sleep(2000);
        //driver.FindElement(By.Id("estadoCivil")).Click();
        //Thread.Sleep(2000);
        ////new SelectElement(driver.FindElement(By.Id("estadoCivil"))).SelectByValue(Cadastro.EstadoCivil);
        //driver.FindElement(By.Id("estadoCivil")).SendKeys("CASADO" + OpenQA.Selenium.Keys.Enter);// PEGAR DO BANCO



        ////Sexo       
        //Thread.Sleep(2000);
        //driver.FindElement(By.Id("sexoDadosPessoais")).Click();
        //Thread.Sleep(2000);
        //driver.FindElement(By.Id("sexoDadosPessoais")).SendKeys("MASCULINO" + OpenQA.Selenium.Keys.Enter); // PEGAR DO BANCO


        ////Deseja Receber a Fatura por email?         
        //Thread.Sleep(2000);
        //driver.FindElement(By.Id("faturaEmailOutros")).Click();
        //driver.FindElement(By.Id("faturaEmailOutros")).SendKeys("N�o"); // PEGAR DO BANCO //


        ////Email
        //Thread.Sleep(2000);
        //driver.FindElement(By.Id("cliEmail")).Click();
        //driver.FindElement(By.Id("cliEmail")).Clear();
        //Thread.Sleep(2000);
        //driver.FindElement(By.Id("cliEmail")).SendKeys("starline@starline.com.br");//PEGAR NO BANCO

        //Thread.Sleep(2000);

        //processTest.StepTurn = 6;
        //processTest.PrintPageComSelenium(driver, false);

        ////Bot�o avan�ar
        //driver.FindElement(By.Id("avancarDadosPessoais")).Click();

        ////Cenario 3 -- Dados Residenciais 

        ////Validar P�gina Dados Residenciais

        //wait(By.Id("avancarDadosResidenciais"));


        //Assert.IsTrue(IsElementPresent(By.Id("formDadosResidenciais")));


        ////Cep
        //Thread.Sleep(2000);
        //driver.FindElement(By.Id("cepDadosResidenciais")).Click();
        //driver.FindElement(By.Id("cepDadosResidenciais")).Clear();
        //Thread.Sleep(2000);
        //driver.FindElement(By.Id("cepDadosResidenciais")).SendKeys("06033080" + OpenQA.Selenium.Keys.Tab);



        ////Numero
        //Thread.Sleep(2000);
        //driver.FindElement(By.Id("numeroDadosResidenciais")).Click();
        //driver.FindElement(By.Id("numeroDadosResidenciais")).Clear();
        //Thread.Sleep(2000);
        //driver.FindElement(By.Id("numeroDadosResidenciais")).SendKeys("1010");

        ////Tipo Residencia
        //Thread.Sleep(2000);
        //driver.FindElement(By.Id("tipoResidencia")).Click();
        //Thread.Sleep(2000);
        //driver.FindElement(By.Id("tipoResidencia")).SendKeys("ALUGADA" + OpenQA.Selenium.Keys.Enter); // PEGAR DO BANCO


        ////Telefone Residencia
        //Thread.Sleep(2000);
        //driver.FindElement(By.Id("telefoneResidencialDadosResidenciais")).Click();
        //driver.FindElement(By.Id("telefoneResidencialDadosResidenciais")).Clear();
        //Thread.Sleep(2000);
        //driver.FindElement(By.Id("telefoneResidencialDadosResidenciais")).SendKeys("1135926538" + OpenQA.Selenium.Keys.Tab);
        //Thread.Sleep(2000);


        //processTest.StepTurn = 7;
        //processTest.PrintPageComSelenium(driver, false);

        ////Bot�o avan�ar

        //driver.FindElement(By.Id("avancarDadosResidenciais")).Click();

        ////Cenario 4 -- Dados Comerciais 

        ////Validar P�gina Dados Comerciais

        //wait(By.Id("avancarDadosComerciais"));


        //Assert.IsTrue(IsElementPresent(By.Id("formDadosComerciais")));


        ////Classe Profissisional
        //Thread.Sleep(2000);
        //driver.FindElement(By.Id("classeProfDadosComerciais")).Click();
        //Thread.Sleep(2000);
        //driver.FindElement(By.Id("classeProfDadosComerciais")).SendKeys("Aposentados" + OpenQA.Selenium.Keys.Enter); // PEGAR DO BANCO


        ////Atividade
        //Thread.Sleep(2000);
        //driver.FindElement(By.Id("atividadeDadosComerciais")).Click();
        //Thread.Sleep(2000);
        //driver.FindElement(By.Id("atividadeDadosComerciais")).SendKeys("Aposentados" + OpenQA.Selenium.Keys.Enter); // PEGAR DO BANCO


        ////Renda Mensal
        //Thread.Sleep(2000);
        //driver.FindElement(By.Id("rendaMensalDadosComerciais")).Click();
        //driver.FindElement(By.Id("rendaMensalDadosComerciais")).Clear();
        //Thread.Sleep(2000);
        //driver.FindElement(By.Id("rendaMensalDadosComerciais")).SendKeys("10000");


        //////Tempo de Empresa (ANO)
        ////Thread.Sleep(2000);
        ////driver.FindElement(By.Id("tempoEmpresaAnoDadosComerciais")).Click();
        ////driver.FindElement(By.Id("tempoEmpresaAnoDadosComerciais")).Clear();
        ////Thread.Sleep(2000);
        ////driver.FindElement(By.Id("tempoEmpresaAnoDadosComerciais")).SendKeys("10");



        //////Tempo de Empresa (MES)
        ////Thread.Sleep(2000);
        ////driver.FindElement(By.Id("tempoEmpresaMes")).Click();
        ////driver.FindElement(By.Id("tempoEmpresaMes")).Clear();
        ////Thread.Sleep(2000);
        ////driver.FindElement(By.Id("tempoEmpresaMes")).SendKeys("10");


        ////CEP
        //Thread.Sleep(2000);
        //driver.FindElement(By.Id("cepDadosComerciais")).Click();
        //driver.FindElement(By.Id("cepDadosComerciais")).Clear();
        //Thread.Sleep(2000);
        //driver.FindElement(By.Id("cepDadosComerciais")).SendKeys("06033050");

        ////Numero
        //Thread.Sleep(2000);
        //driver.FindElement(By.Id("numeroDadosComerciais")).Click();
        //driver.FindElement(By.Id("numeroDadosComerciais")).Clear();
        //Thread.Sleep(2000);
        //driver.FindElement(By.Id("numeroDadosComerciais")).Click();
        //driver.FindElement(By.Id("numeroDadosComerciais")).SendKeys("11");


        ////Telefone Comercial
        //Thread.Sleep(2000);
        //driver.FindElement(By.Id("telefoneComercialDadosComercial")).Click();
        //driver.FindElement(By.Id("telefoneComercialDadosComercial")).Clear();
        //Thread.Sleep(2000);
        //driver.FindElement(By.Id("telefoneComercialDadosComercial")).Click();
        //driver.FindElement(By.Id("telefoneComercialDadosComercial")).SendKeys("1136811452" + OpenQA.Selenium.Keys.Tab);
        //Thread.Sleep(2000);

        ////Bot�o avan�ar

        //driver.FindElement(By.Id("avancarDadosComerciais")).Click();

        ////Cenario 5 -- Outros 

        ////Validar P�gina Outros

        //wait(By.Id("avancarOutros"));


        //Assert.IsTrue(IsElementPresent(By.Id("formOutros")));

        ////Vencimento Melhor Dia
        //Thread.Sleep(2000);
        //driver.FindElement(By.Id("vencimentoOutros")).Click();
        //driver.FindElement(By.Id("vencimentoOutros")).SendKeys("10");
        //Thread.Sleep(2000);

        ////Confirma��o de Vencimento
        //driver.FindElement(By.XPath("(//div[@id='modalContent']/div[3]/button)[3]")).Click();

        //wait(By.CssSelector("#tarifaOverlimitOutros"));

        ////Cobrar tarifa de overlimit
        //driver.FindElement(By.CssSelector("#tarifaOverlimitOutros")).Click();
        //driver.FindElement(By.CssSelector("#tarifaOverlimitOutros")).SendKeys("N�o");

        ////Conta B�nus?
        //Thread.Sleep(2000);
        //driver.FindElement(By.Id("tipoCartaoOutros")).Click();
        //driver.FindElement(By.Id("tipoCartaoOutros")).SendKeys("N�o");


        ////Campanha
        //Thread.Sleep(2000);
        //driver.FindElement(By.Id("campanhaOutros")).Click();
        //driver.FindElement(By.Id("campanhaOutros")).SendKeys("931 - Desc Primeira Compra Marisa Itaucard 10%");
        //Thread.Sleep(2000);


        ////Bot�o avan�ar
        //touchActions = new Actions(driver);
        //touchActions.Click(driver.FindElement(By.Id("avancarOutros")));
        //touchActions.Build();
        //touchActions.Perform();

        ////Cenario 6 -- Camera 

        ////Validar P�gina Outros

        //wait(By.Id("avancarBiometriaFacial"));

        ////driver.FindElement(By.Id("avancarBiometriaFacial")).Click();

        //touchActions = new Actions(driver);
        //touchActions.Click(driver.FindElement(By.Id("avancarBiometriaFacial")));
        //touchActions.Build();
        //touchActions.Perform();

        //driver.SwitchTo().ParentFrame();
        //driver.SwitchTo().Frame("assinaturaEletronica");
        //wait(By.XPath("/html/body/form/footer/input"));
        //touchActions = new Actions(driver);
        //touchActions.Click(driver.FindElement(By.XPath("/html/body/form/footer/input")));
        //touchActions.Build();
        //touchActions.Perform();


        ////anexar imagem
        //driver.SwitchTo().ParentFrame();
        //driver.SwitchTo().Frame("assinaturaEletronica");
        //driver.FindElement(By.Id("photo_37")).SendKeys(@"C:\Users\Rafael Leal\Pictures\teste.jpg");

        //Thread.Sleep(20000);

        //driver.SwitchTo().ParentFrame();
        //driver.SwitchTo().Frame("assinaturaEletronica");
        //touchActions = new Actions(driver);
        //touchActions.DragAndDrop(driver.FindElement(By.Id("imgOrientation")), driver.FindElement(By.Id("signatureLine")));
        //touchActions.Build();
        //touchActions.Perform();
        //processTest.StepTurn = 8;
        //processTest.PrintPageComSelenium(driver, false);


        ////Aceitar a assinatura
        //driver.FindElement(By.Id("btnSave")).Click();
        //Thread.Sleep(2000);

        //wait(By.CssSelector("btnInicioCadastro"));

        ////Resultado.
        //driver.FindElement(By.Id("btnInicioCadastro")).Click();

        //processTest.StepTurn = 9;
        //processTest.PrintPageComSelenium(driver, false);



        //driver.Quit();

        //}



        //[Test]
        //public void CT002_Cadastro_de_Clientes_Autonomo()
        //{



        //    //M�todo de Print
        //    ScreenShot print = new ScreenShot("VendaCartoes", "CT001_Cadastro_de_Clientes");

        //    //M�todo simulador de teclado
        //    InputSimulator simulator = new InputSimulator();

        //    //Execu��o

        //    //Acessar P�gina do CCM
        //    driver.Navigate().GoToUrl(baseURL);


        //    wait(By.XPath("//button[@type='submit']"));
        //    wait(By.CssSelector("img.png_bg"));
        //    print.PrintScreen();

        //    try
        //    {
        //        Assert.AreEqual("Login", driver.FindElement(By.CssSelector("strong")).Text);
        //    }
        //    catch (AssertionException e)
        //    {
        //        verificationErrors.Append(e.Message);
        //    }
        //    Assert.IsTrue(IsElementPresent(By.Id("login_username")));
        //    Assert.IsTrue(IsElementPresent(By.Id("login_password")));
        //    Assert.IsTrue(IsElementPresent(By.XPath("//button[@type='submit']")));



        //    //Validar Tela de Login
        //    Assert.AreEqual("PSF Security", driver.Title);
        //    Assert.IsTrue(IsElementPresent(By.CssSelector("img.png_bg")));
        //    try
        //    {
        //        Assert.AreEqual("Portal Lojas", driver.FindElement(By.Id("tituloPagina")).Text);
        //    }
        //    catch (AssertionException e)
        //    {
        //        verificationErrors.Append(e.Message);
        //    }

        //    print.PrintScreen();

        //    //Realizar login
        //    Login();


        //    wait(By.Id("btnSearch"));
        //    print.PrintScreen();


        //    ////Alterar Filial
        //    Thread.Sleep(2000);
        //    SendKeys.SendWait("{ESC}");
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("labelFilial")).Click();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("changeFilial")).SendKeys("54 - L 054 JD / SP");
        //    Thread.Sleep(2000);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);
        //    Thread.Sleep(2000);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);
        //    Thread.Sleep(5000);


        //    //Validar P�gina Inicial
        //    Assert.AreEqual("PSF Security", driver.Title);

        //    try
        //    {
        //        Assert.AreEqual("Buscar por:\r\nCPF\r\nCART�O", driver.FindElement(By.Id("tipoSearch")).Text);
        //    }
        //    catch (AssertionException e)
        //    {
        //        verificationErrors.Append(e.Message);
        //    }
        //    Assert.IsTrue(IsElementPresent(By.Id("cpfSearch")));
        //    Assert.IsTrue(IsElementPresent(By.Id("btnSearch")));
        //    try
        //    {
        //        Assert.AreEqual("Concessao do Cartao", driver.FindElement(By.Id("menuCollapse3162")).Text);
        //    }
        //    catch (AssertionException e)
        //    {
        //        verificationErrors.Append(e.Message);
        //    }
        //    try
        //    {
        //        Assert.AreEqual("Atendimento", driver.FindElement(By.Id("menuCollapse3164")).Text);
        //    }
        //    catch (AssertionException e)
        //    {
        //        verificationErrors.Append(e.Message);
        //    }


        //    //Cenario 1 - Cadastro de Clientes

        //    //Consultar Cliente
        //    driver.FindElement(By.Id("cpfSearch")).Clear();
        //    driver.FindElement(By.Id("cpfSearch")).SendKeys("15774699846");
        //    Thread.Sleep(2000);
        //    print.PrintScreen();
        //    driver.FindElement(By.Id("btnSearch")).Click();
        //    Thread.Sleep(2000);



        //    wait(By.Id("iniciarCadastro"));


        //    //Validar P�gina Cadastro de Clientes


        //    Assert.IsTrue(IsElementPresent(By.Id("formPreCadastro")));

        //    driver.FindElement(By.Id("dataNascCad")).Click();
        //    driver.FindElement(By.Id("dataNascCad")).SendKeys("30/10/1970");
        //    driver.FindElement(By.Id("iniciarCadastro")).Click();


        //    //Validar Mensagem de Retorno CPF

        //    Thread.Sleep(8000);


        //    //if (this.IsElementPresent(By.Id("dialog-message")))

        //    //{
        //    //    wait(By.CssSelector("#modalContent > div.modal-body.row"));
        //    //    string Retorno = driver.FindElement(By.CssSelector("#modalContent > div.modal-body.row")).Text;
        //    //    // var setMensagemRetorno =
        //    //    Marisa.TestDataAccess.ExcelDataAccess.SetMesagemRetorno(Retorno, Cadastro.CPF);
        //    //    Thread.Sleep(2000);
        //    //    print.PrintScreen();

        //    //    //refazer o contador
        //    //    //int counter = 0;
        //    //    //int NumeroCPF = Marisa.TestDataAccess.ExcelDataAccess.GetSemMensagem().total;

        //    //    //while (counter < NumeroCPF) ;

        //    //}


        //    wait(By.Id("avancarDadosPessoais"));

        //    //Cenario 2  -- Dados pessoais


        //    //Validar P�gina Cadastro de Clientes


        //    Assert.IsTrue(IsElementPresent(By.Id("formDadosPessoais")));


        //    //Nome da M�e
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("nomeMaeDadosPessoais")).Click();
        //    driver.FindElement(By.Id("nomeMaeDadosPessoais")).Clear();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("nomeMaeDadosPessoais")).SendKeys("Maria Aparecida");


        //    //Estado Civil
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("estadoCivil")).Click();
        //    Thread.Sleep(2000);
        //    //new SelectElement(driver.FindElement(By.Id("estadoCivil"))).SelectByValue(Cadastro.EstadoCivil);
        //    driver.FindElement(By.Id("estadoCivil")).SendKeys("CASADO");// PEGAR DO BANCO
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);



        //    //Sexo       
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("sexoDadosPessoais")).Click();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("sexoDadosPessoais")).SendKeys("MASCULINO"); // PEGAR DO BANCO
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);


        //    //Deseja Receber a Fatura por email?         
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("faturaEmailOutros")).Click();
        //    driver.FindElement(By.Id("faturaEmailOutros")).SendKeys("N�o"); // PEGAR DO BANCO //


        //    //Email
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("cliEmail")).Click();
        //    driver.FindElement(By.Id("cliEmail")).Clear();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("cliEmail")).SendKeys("starline@starline.com.br");//PEGAR NO BANCO

        //    Thread.Sleep(2000);

        //    print.PrintScreen();

        //    //Bot�o avan�ar
        //    driver.FindElement(By.Id("avancarDadosPessoais")).Click();

        //    //Cenario 3 -- Dados Residenciais 

        //    //Validar P�gina Dados Residenciais

        //    wait(By.Id("avancarDadosResidenciais"));


        //    Assert.IsTrue(IsElementPresent(By.Id("formDadosResidenciais")));

        //    //Cep
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("cepDadosResidenciais")).Click();
        //    driver.FindElement(By.Id("cepDadosResidenciais")).Clear();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("cepDadosResidenciais")).SendKeys("06033050");
        //    Thread.Sleep(2000);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);



        //    //Numero
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("numeroDadosResidenciais")).Click();
        //    driver.FindElement(By.Id("numeroDadosResidenciais")).Clear();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("numeroDadosResidenciais")).SendKeys("1010");

        //    //Tipo Residencia
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("tipoResidencia")).Click();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("tipoResidencia")).SendKeys("ALUGADA"); // PEGAR DO BANCO
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);


        //    //Telefone Residencia
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("telefoneResidencialDadosResidenciais")).Click();
        //    driver.FindElement(By.Id("telefoneResidencialDadosResidenciais")).Clear();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("telefoneResidencialDadosResidenciais")).SendKeys("1135926538");
        //    Thread.Sleep(2000);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);



        //    print.PrintScreen();

        //    //Bot�o avan�ar
        //    //Thread.Sleep(2000);
        //    //driver.FindElement(By.Id("avancarDadosResidenciais"));

        //    //Cenario 4 -- Dados Comerciais 

        //    //Validar P�gina Dados Comerciais

        //    wait(By.Id("avancarDadosComerciais"));


        //    Assert.IsTrue(IsElementPresent(By.Id("formDadosComerciais")));


        //    //Classe Profissisional
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("classeProfDadosComerciais")).Click();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("classeProfDadosComerciais")).SendKeys("Autonomo"); // PEGAR DO BANCO
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);

        //    //Atividade
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("atividadeDadosComerciais")).Click();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("atividadeDadosComerciais")).SendKeys("Autonomo"); // PEGAR DO BANCO
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);


        //    //Renda Mensal
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("rendaMensalDadosComerciais")).Click();
        //    driver.FindElement(By.Id("rendaMensalDadosComerciais")).Clear();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("rendaMensalDadosComerciais")).SendKeys("10000");


        //    ////Tempo de Empresa (ANO)
        //    //Thread.Sleep(2000);
        //    //driver.FindElement(By.Id("tempoEmpresaAnoDadosComerciais")).Click();
        //    //driver.FindElement(By.Id("tempoEmpresaAnoDadosComerciais")).Clear();
        //    //Thread.Sleep(2000);
        //    //driver.FindElement(By.Id("tempoEmpresaAnoDadosComerciais")).SendKeys("10");



        //    ////Tempo de Empresa (MES)
        //    //Thread.Sleep(2000);
        //    //driver.FindElement(By.Id("tempoEmpresaMes")).Click();
        //    //driver.FindElement(By.Id("tempoEmpresaMes")).Clear();
        //    //Thread.Sleep(2000);
        //    //driver.FindElement(By.Id("tempoEmpresaMes")).SendKeys("10");


        //    //CEP
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("cepDadosComerciais")).Click();
        //    driver.FindElement(By.Id("cepDadosComerciais")).Clear();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("cepDadosComerciais")).SendKeys("06033050");

        //    //Numero
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("numeroDadosComerciais")).Click();
        //    driver.FindElement(By.Id("numeroDadosComerciais")).Clear();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("numeroDadosComerciais")).SendKeys("365");

        //    //Telefone Comercial
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("telefoneComercialDadosComercial")).Click();
        //    driver.FindElement(By.Id("telefoneComercialDadosComercial")).Clear();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("telefoneComercialDadosComercial")).SendKeys("1136811452");
        //    Thread.Sleep(2000);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);


        //    //Bot�o avan�ar
        //    //Thread.Sleep(2000);
        //    //driver.FindElement(By.Id("avancarDadosComerciais"));

        //    //Cenario 5 -- Outros 

        //    //Validar P�gina Outros

        //    wait(By.Id("avancarOutros"));


        //    Assert.IsTrue(IsElementPresent(By.Id("formOutros")));

        //    //Vencimento Melhor Dia
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("vencimentoOutros")).Click();
        //    driver.FindElement(By.Id("vencimentoOutros")).SendKeys("10");
        //    Thread.Sleep(2000);

        //    //Confirma��o de Vencimento
        //    SendKeys.SendWait("{ESC}");
        //    Thread.Sleep(2000);
        //    SendKeys.SendWait("{ESC}");
        //    Thread.Sleep(2000);
        //    //Cobrar tarifa de overlimit


        //    wait(By.CssSelector("#tarifaOverlimitOutros"));

        //    driver.FindElement(By.CssSelector("#tarifaOverlimitOutros")).Click();
        //    driver.FindElement(By.CssSelector("#tarifaOverlimitOutros")).SendKeys("N�o");

        //    //Conta B�nus?
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("tipoCartaoOutros")).Click();
        //    driver.FindElement(By.Id("tipoCartaoOutros")).SendKeys("N�o");


        //    //Campanha
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("campanhaOutros")).Click();
        //    driver.FindElement(By.Id("campanhaOutros")).SendKeys("931 - Desc Primeira Compra Marisa Itaucard 10%");
        //    Thread.Sleep(2000);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);


        //    //Bot�o avan�ar
        //    //Thread.Sleep(2000);
        //    driver.FindElement(By.Id("avancarOutros")).Click();

        //    //Cenario 6 -- Camera 

        //    //Validar P�gina Outros

        //    wait(By.Id("avancarBiometriaFacial"));






        //    //   Assert.IsTrue(IsElementPresent(By.Id("fieldSetBiometria")));

        //    //Ligar Camera
        //    Thread.Sleep(2000);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);


        //    driver.Quit();


        //}

        //////Tirar foto
        ////driver.SwitchTo();
        ////    driver.FindElement(By.Id("btnCamera")).Click();
        ////    Thread.Sleep(200);
        ////    driver.FindElement(By.CssSelector("#pnlCapturePreview > div > div.section-content.buttons > div > div.imagebutton.accept.dark")).Click();
        ////    Thread.Sleep(20000);

        ////    //Bot�o Avan�ar
        ////    driver.FindElement(By.Id("avancarBiometriaFacial")).Click();

        ////    //Cenario 7 -- Formaliza��o 

        ////    //Validar P�gina Outros

        ////    wait(By.Id("avancarBiometriaFacial"));


        ////    Assert.IsTrue(IsElementPresent(By.Id("panelAssinaturaEletronica")));

        ////    driver.Quit();


        ////}

        //[Test]
        //public void CT003_Cadastro_de_Clientes_Assalariado()
        //{




        //    //M�todo de Print
        //    ScreenShot print = new ScreenShot("VendaCartoes", "CT001_Cadastro_de_Clientes");

        //    //M�todo simulador de teclado
        //    InputSimulator simulator = new InputSimulator();

        //    //Execu��o

        //    //Acessar P�gina do CCM
        //    driver.Navigate().GoToUrl(baseURL);


        //    wait(By.XPath("//button[@type='submit']"));
        //    wait(By.CssSelector("img.png_bg"));
        //    print.PrintScreen();

        //    try
        //    {
        //        Assert.AreEqual("Login", driver.FindElement(By.CssSelector("strong")).Text);
        //    }
        //    catch (AssertionException e)
        //    {
        //        verificationErrors.Append(e.Message);
        //    }
        //    Assert.IsTrue(IsElementPresent(By.Id("login_username")));
        //    Assert.IsTrue(IsElementPresent(By.Id("login_password")));
        //    Assert.IsTrue(IsElementPresent(By.XPath("//button[@type='submit']")));



        //    //Validar Tela de Login
        //    Assert.AreEqual("PSF Security", driver.Title);
        //    Assert.IsTrue(IsElementPresent(By.CssSelector("img.png_bg")));
        //    try
        //    {
        //        Assert.AreEqual("Portal Lojas", driver.FindElement(By.Id("tituloPagina")).Text);
        //    }
        //    catch (AssertionException e)
        //    {
        //        verificationErrors.Append(e.Message);
        //    }

        //    print.PrintScreen();

        //    //Realizar login
        //    Login();


        //    wait(By.Id("btnSearch"));
        //    print.PrintScreen();


        //    ////Alterar Filial
        //    Thread.Sleep(2000);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.ESCAPE);
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("labelFilial")).Click();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("changeFilial")).SendKeys("54 - L 054 JD / SP");
        //    Thread.Sleep(2000);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);
        //    Thread.Sleep(2000);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);
        //    Thread.Sleep(5000);


        //    //Validar P�gina Inicial
        //    Assert.AreEqual("PSF Security", driver.Title);

        //    try
        //    {
        //        Assert.AreEqual("Buscar por:\r\nCPF\r\nCART�O", driver.FindElement(By.Id("tipoSearch")).Text);
        //    }
        //    catch (AssertionException e)
        //    {
        //        verificationErrors.Append(e.Message);
        //    }
        //    Assert.IsTrue(IsElementPresent(By.Id("cpfSearch")));
        //    Assert.IsTrue(IsElementPresent(By.Id("btnSearch")));
        //    try
        //    {
        //        Assert.AreEqual("Concessao do Cartao", driver.FindElement(By.Id("menuCollapse3162")).Text);
        //    }
        //    catch (AssertionException e)
        //    {
        //        verificationErrors.Append(e.Message);
        //    }
        //    try
        //    {
        //        Assert.AreEqual("Atendimento", driver.FindElement(By.Id("menuCollapse3164")).Text);
        //    }
        //    catch (AssertionException e)
        //    {
        //        verificationErrors.Append(e.Message);
        //    }


        //    //Cenario 1 - Cadastro de Clientes

        //    //Consultar Cliente
        //    driver.FindElement(By.Id("cpfSearch")).Clear();
        //    driver.FindElement(By.Id("cpfSearch")).SendKeys("46777196838"); // PEGAR DO BANCO
        //    Thread.Sleep(2000);
        //    print.PrintScreen();
        //    driver.FindElement(By.Id("btnSearch")).Click();
        //    Thread.Sleep(2000);



        //    wait(By.Id("iniciarCadastro"));


        //    //Validar P�gina Cadastro de Clientes


        //    Assert.IsTrue(IsElementPresent(By.Id("formPreCadastro")));

        //    driver.FindElement(By.Id("dataNascCad")).Click();
        //    driver.FindElement(By.Id("dataNascCad")).SendKeys("10/05/1996");
        //    driver.FindElement(By.Id("iniciarCadastro")).Click();


        //    //Validar Mensagem de Retorno CPF

        //    Thread.Sleep(8000);


        //    //if (this.IsElementPresent(By.Id("dialog-message")))

        //    //{
        //    //    wait(By.CssSelector("#modalContent > div.modal-body.row"));
        //    //    string Retorno = driver.FindElement(By.CssSelector("#modalContent > div.modal-body.row")).Text;
        //    //    // var setMensagemRetorno =
        //    //    Marisa.TestDataAccess.ExcelDataAccess.SetMesagemRetorno(Retorno, Cadastro.CPF);
        //    //    Thread.Sleep(2000);
        //    //    print.PrintScreen();

        //    //    //refazer o contador
        //    //    //int counter = 0;
        //    //    //int NumeroCPF = Marisa.TestDataAccess.ExcelDataAccess.GetSemMensagem().total;

        //    //    //while (counter < NumeroCPF) ;

        //    //}


        //    wait(By.Id("avancarDadosPessoais"));

        //    //Cenario 2  -- Dados pessoais


        //    //Validar P�gina Cadastro de Clientes


        //    Assert.IsTrue(IsElementPresent(By.Id("formDadosPessoais")));


        //    //Nome da M�e
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("nomeMaeDadosPessoais")).Click();
        //    driver.FindElement(By.Id("nomeMaeDadosPessoais")).Clear();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("nomeMaeDadosPessoais")).SendKeys("Maria Aparecida");


        //    //Estado Civil
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("estadoCivil")).Click();
        //    Thread.Sleep(2000);
        //    //new SelectElement(driver.FindElement(By.Id("estadoCivil"))).SelectByValue(Cadastro.EstadoCivil);
        //    driver.FindElement(By.Id("estadoCivil")).SendKeys("CASADO");// PEGAR DO BANCO
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);



        //    //Sexo       
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("sexoDadosPessoais")).Click();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("sexoDadosPessoais")).SendKeys("MASCULINO"); // PEGAR DO BANCO
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);


        //    //Deseja Receber a Fatura por email?         
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("faturaEmailOutros")).Click();
        //    driver.FindElement(By.Id("faturaEmailOutros")).SendKeys("N�o"); // PEGAR DO BANCO //


        //    //Email
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("cliEmail")).Click();
        //    driver.FindElement(By.Id("cliEmail")).Clear();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("cliEmail")).SendKeys("starline@starline.com.br");//PEGAR NO BANCO

        //    Thread.Sleep(2000);

        //    print.PrintScreen();

        //    //Bot�o avan�ar
        //    driver.FindElement(By.Id("avancarDadosPessoais")).Click();

        //    //Cenario 3 -- Dados Residenciais 

        //    //Validar P�gina Dados Residenciais

        //    wait(By.Id("avancarDadosResidenciais"));


        //    Assert.IsTrue(IsElementPresent(By.Id("formDadosResidenciais")));

        //    //Cep
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("cepDadosResidenciais")).Click();
        //    driver.FindElement(By.Id("cepDadosResidenciais")).Clear();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("cepDadosResidenciais")).SendKeys("06033050");
        //    Thread.Sleep(2000);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);



        //    //Numero
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("numeroDadosResidenciais")).Click();
        //    driver.FindElement(By.Id("numeroDadosResidenciais")).Clear();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("numeroDadosResidenciais")).SendKeys("1010");

        //    //Tipo Residencia
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("tipoResidencia")).Click();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("tipoResidencia")).SendKeys("ALUGADA"); // PEGAR DO BANCO
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);


        //    //Telefone Residencia
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("telefoneResidencialDadosResidenciais")).Click();
        //    driver.FindElement(By.Id("telefoneResidencialDadosResidenciais")).Clear();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("telefoneResidencialDadosResidenciais")).SendKeys("1135926538");
        //    Thread.Sleep(2000);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);



        //    print.PrintScreen();

        //    //Bot�o avan�ar
        //    //Thread.Sleep(2000);
        //    //driver.FindElement(By.Id("avancarDadosResidenciais"));

        //    //Cenario 4 -- Dados Comerciais 

        //    //Validar P�gina Dados Comerciais

        //    wait(By.Id("avancarDadosComerciais"));


        //    Assert.IsTrue(IsElementPresent(By.Id("formDadosComerciais")));


        //    //Classe Profissisional
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("classeProfDadosComerciais")).Click();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("classeProfDadosComerciais")).SendKeys("Assalariado"); // PEGAR DO BANCO
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);

        //    //Atividade
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("atividadeDadosComerciais")).Click();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("atividadeDadosComerciais")).SendKeys("Assalariado"); // PEGAR DO BANCO
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);


        //    //Renda Mensal
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("rendaMensalDadosComerciais")).Click();
        //    driver.FindElement(By.Id("rendaMensalDadosComerciais")).Clear();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("rendaMensalDadosComerciais")).SendKeys("10000");


        //    //Tempo de Empresa (ANO)
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("tempoEmpresaAnoDadosComerciais")).Click();
        //    driver.FindElement(By.Id("tempoEmpresaAnoDadosComerciais")).Clear();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("tempoEmpresaAnoDadosComerciais")).SendKeys("10");



        //    //Tempo de Empresa (MES)
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("tempoEmpresaMes")).Click();
        //    driver.FindElement(By.Id("tempoEmpresaMes")).Clear();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("tempoEmpresaMes")).SendKeys("10");


        //    //CEP
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("cepDadosComerciais")).Click();
        //    driver.FindElement(By.Id("cepDadosComerciais")).Clear();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("cepDadosComerciais")).SendKeys("06033050");

        //    //Numero
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("numeroDadosComerciais")).Click();
        //    driver.FindElement(By.Id("numeroDadosComerciais")).Clear();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("numeroDadosComerciais")).SendKeys("365");

        //    //Telefone Comercial
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("telefoneComercialDadosComercial")).Click();
        //    driver.FindElement(By.Id("telefoneComercialDadosComercial")).Clear();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("telefoneComercialDadosComercial")).SendKeys("1136811452");
        //    Thread.Sleep(2000);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);

        //    //Bot�o avan�ar
        //    //Thread.Sleep(2000);
        //    //driver.FindElement(By.Id("avancarDadosComerciais"));

        //    //Cenario 5 -- Outros 

        //    //Validar P�gina Outros

        //    wait(By.Id("avancarOutros"));


        //    Assert.IsTrue(IsElementPresent(By.Id("formOutros")));

        //    //Vencimento Melhor Dia
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("vencimentoOutros")).Click();
        //    driver.FindElement(By.Id("vencimentoOutros")).SendKeys("10");
        //    Thread.Sleep(2000);

        //    //Confirma��o de Vencimento
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.ESCAPE);
        //    Thread.Sleep(2000);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.ESCAPE);
        //    Thread.Sleep(2000);
        //    //Cobrar tarifa de overlimit


        //    wait(By.CssSelector("#tarifaOverlimitOutros"));

        //    driver.FindElement(By.CssSelector("#tarifaOverlimitOutros")).Click();
        //    driver.FindElement(By.CssSelector("#tarifaOverlimitOutros")).SendKeys("N�o");

        //    //Conta B�nus?
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("tipoCartaoOutros")).Click();
        //    driver.FindElement(By.Id("tipoCartaoOutros")).SendKeys("N�o");


        //    //Campanha
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("campanhaOutros")).Click();
        //    driver.FindElement(By.Id("campanhaOutros")).SendKeys("931 - Desc Primeira Compra Marisa Itaucard 10%");
        //    Thread.Sleep(2000);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);


        //    //Bot�o avan�ar
        //    //Thread.Sleep(2000);
        //    driver.FindElement(By.Id("avancarOutros")).Click();

        //    //Cenario 6 -- Camera 

        //    //Validar P�gina Outros

        //    wait(By.Id("avancarBiometriaFacial"));



        //    //   Assert.IsTrue(IsElementPresent(By.Id("fieldSetBiometria")));

        //    //Ligar Camera
        //    Thread.Sleep(2000);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);


        //    Thread.Sleep(2000);


        //    driver.Quit();


        //}

        //////Tirar foto
        ////driver.SwitchTo();
        ////    driver.FindElement(By.Id("btnCamera")).Click();
        ////    Thread.Sleep(200);
        ////    driver.FindElement(By.CssSelector("#pnlCapturePreview > div > div.section-content.buttons > div > div.imagebutton.accept.dark")).Click();
        ////    Thread.Sleep(20000);

        ////    //Bot�o Avan�ar
        ////    driver.FindElement(By.Id("avancarBiometriaFacial")).Click();

        ////    //Cenario 7 -- Formaliza��o 

        ////    //Validar P�gina Outros

        ////    wait(By.Id("avancarBiometriaFacial"));


        ////    Assert.IsTrue(IsElementPresent(By.Id("panelAssinaturaEletronica")));

        ////    driver.Quit();


        ////}


        //[Test]
        //public void CT004_Cadastro_de_Clientes_Estabelecidos()
        //{


        //    //M�todo de Print
        //    ScreenShot print = new ScreenShot("VendaCartoes", "CT001_Cadastro_de_Clientes");

        //    //M�todo simulador de teclado
        //    InputSimulator simulator = new InputSimulator();

        //    //Execu��o

        //    //Acessar P�gina do CCM
        //    driver.Navigate().GoToUrl(baseURL);


        //    wait(By.XPath("//button[@type='submit']"));
        //    wait(By.CssSelector("img.png_bg"));
        //    print.PrintScreen();

        //    try
        //    {
        //        Assert.AreEqual("Login", driver.FindElement(By.CssSelector("strong")).Text);
        //    }
        //    catch (AssertionException e)
        //    {
        //        verificationErrors.Append(e.Message);
        //    }
        //    Assert.IsTrue(IsElementPresent(By.Id("login_username")));
        //    Assert.IsTrue(IsElementPresent(By.Id("login_password")));
        //    Assert.IsTrue(IsElementPresent(By.XPath("//button[@type='submit']")));



        //    //Validar Tela de Login
        //    Assert.AreEqual("PSF Security", driver.Title);
        //    Assert.IsTrue(IsElementPresent(By.CssSelector("img.png_bg")));
        //    try
        //    {
        //        Assert.AreEqual("Portal Lojas", driver.FindElement(By.Id("tituloPagina")).Text);
        //    }
        //    catch (AssertionException e)
        //    {
        //        verificationErrors.Append(e.Message);
        //    }

        //    print.PrintScreen();

        //    //Realizar login
        //    Login();


        //    wait(By.Id("btnSearch"));
        //    print.PrintScreen();


        //    ////Alterar Filial
        //    Thread.Sleep(2000);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.ESCAPE);
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("labelFilial")).Click();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("changeFilial")).SendKeys("54 - L 054 JD / SP");
        //    Thread.Sleep(2000);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);
        //    Thread.Sleep(2000);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);
        //    Thread.Sleep(5000);


        //    //Validar P�gina Inicial
        //    Assert.AreEqual("PSF Security", driver.Title);

        //    try
        //    {
        //        Assert.AreEqual("Buscar por:\r\nCPF\r\nCART�O", driver.FindElement(By.Id("tipoSearch")).Text);
        //    }
        //    catch (AssertionException e)
        //    {
        //        verificationErrors.Append(e.Message);
        //    }
        //    Assert.IsTrue(IsElementPresent(By.Id("cpfSearch")));
        //    Assert.IsTrue(IsElementPresent(By.Id("btnSearch")));
        //    try
        //    {
        //        Assert.AreEqual("Concessao do Cartao", driver.FindElement(By.Id("menuCollapse3162")).Text);
        //    }
        //    catch (AssertionException e)
        //    {
        //        verificationErrors.Append(e.Message);
        //    }
        //    try
        //    {
        //        Assert.AreEqual("Atendimento", driver.FindElement(By.Id("menuCollapse3164")).Text);
        //    }
        //    catch (AssertionException e)
        //    {
        //        verificationErrors.Append(e.Message);
        //    }


        //    //Cenario 1 - Cadastro de Clientes

        //    //Consultar Cliente
        //    driver.FindElement(By.Id("cpfSearch")).Clear();
        //    driver.FindElement(By.Id("cpfSearch")).SendKeys("7007739830");
        //    Thread.Sleep(2000);
        //    print.PrintScreen();
        //    driver.FindElement(By.Id("btnSearch")).Click();
        //    Thread.Sleep(2000);



        //    wait(By.Id("iniciarCadastro"));


        //    //Validar P�gina Cadastro de Clientes


        //    Assert.IsTrue(IsElementPresent(By.Id("formPreCadastro")));

        //    driver.FindElement(By.Id("dataNascCad")).Click();
        //    driver.FindElement(By.Id("dataNascCad")).SendKeys("14/10/1967");
        //    driver.FindElement(By.Id("iniciarCadastro")).Click();


        //    //Validar Mensagem de Retorno CPF

        //    Thread.Sleep(8000);


        //    //if (this.IsElementPresent(By.Id("dialog-message")))

        //    //{
        //    //    wait(By.CssSelector("#modalContent > div.modal-body.row"));
        //    //    string Retorno = driver.FindElement(By.CssSelector("#modalContent > div.modal-body.row")).Text;
        //    //    // var setMensagemRetorno =
        //    //    Marisa.TestDataAccess.ExcelDataAccess.SetMesagemRetorno(Retorno, Cadastro.CPF);
        //    //    Thread.Sleep(2000);
        //    //    print.PrintScreen();

        //    //    //refazer o contador
        //    //    //int counter = 0;
        //    //    //int NumeroCPF = Marisa.TestDataAccess.ExcelDataAccess.GetSemMensagem().total;

        //    //    //while (counter < NumeroCPF) ;

        //    //}


        //    wait(By.Id("avancarDadosPessoais"));

        //    //Cenario 2  -- Dados pessoais


        //    //Validar P�gina Cadastro de Clientes


        //    Assert.IsTrue(IsElementPresent(By.Id("formDadosPessoais")));


        //    //Nome da M�e
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("nomeMaeDadosPessoais")).Click();
        //    driver.FindElement(By.Id("nomeMaeDadosPessoais")).Clear();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("nomeMaeDadosPessoais")).SendKeys("Maria Aparecida");


        //    //Estado Civil
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("estadoCivil")).Click();
        //    Thread.Sleep(2000);
        //    //new SelectElement(driver.FindElement(By.Id("estadoCivil"))).SelectByValue(Cadastro.EstadoCivil);
        //    driver.FindElement(By.Id("estadoCivil")).SendKeys("CASADO");// PEGAR DO BANCO
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);



        //    //Sexo       
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("sexoDadosPessoais")).Click();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("sexoDadosPessoais")).SendKeys("MASCULINO"); // PEGAR DO BANCO
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);


        //    //Deseja Receber a Fatura por email?         
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("faturaEmailOutros")).Click();
        //    driver.FindElement(By.Id("faturaEmailOutros")).SendKeys("N�o"); // PEGAR DO BANCO //


        //    //Email
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("cliEmail")).Click();
        //    driver.FindElement(By.Id("cliEmail")).Clear();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("cliEmail")).SendKeys("starline@starline.com.br");//PEGAR NO BANCO

        //    Thread.Sleep(2000);

        //    print.PrintScreen();

        //    //Bot�o avan�ar
        //    driver.FindElement(By.Id("avancarDadosPessoais")).Click();

        //    //Cenario 3 -- Dados Residenciais 

        //    //Validar P�gina Dados Residenciais

        //    wait(By.Id("avancarDadosResidenciais"));


        //    Assert.IsTrue(IsElementPresent(By.Id("formDadosResidenciais")));

        //    //Cep
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("cepDadosResidenciais")).Click();
        //    driver.FindElement(By.Id("cepDadosResidenciais")).Clear();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("cepDadosResidenciais")).SendKeys("06033050");
        //    Thread.Sleep(2000);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);



        //    //Numero
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("numeroDadosResidenciais")).Click();
        //    driver.FindElement(By.Id("numeroDadosResidenciais")).Clear();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("numeroDadosResidenciais")).SendKeys("1010");

        //    //Tipo Residencia
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("tipoResidencia")).Click();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("tipoResidencia")).SendKeys("ALUGADA"); // PEGAR DO BANCO
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);


        //    //Telefone Residencia
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("telefoneResidencialDadosResidenciais")).Click();
        //    driver.FindElement(By.Id("telefoneResidencialDadosResidenciais")).Clear();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("telefoneResidencialDadosResidenciais")).SendKeys("1135926538");
        //    Thread.Sleep(2000);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);



        //    print.PrintScreen();

        //    //Bot�o avan�ar
        //    //Thread.Sleep(2000);
        //    //driver.FindElement(By.Id("avancarDadosResidenciais"));

        //    //Cenario 4 -- Dados Comerciais 

        //    //Validar P�gina Dados Comerciais

        //    wait(By.Id("avancarDadosComerciais"));


        //    Assert.IsTrue(IsElementPresent(By.Id("formDadosComerciais")));


        //    //Classe Profissisional
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("classeProfDadosComerciais")).Click();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("classeProfDadosComerciais")).SendKeys("OUTROS"); // PEGAR DO BANCO
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);

        //    //Atividade
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("atividadeDadosComerciais")).Click();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("atividadeDadosComerciais")).SendKeys("DO LAR"); // PEGAR DO BANCO
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);


        //    //Renda Mensal
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("rendaMensalDadosComerciais")).Click();
        //    driver.FindElement(By.Id("rendaMensalDadosComerciais")).Clear();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("rendaMensalDadosComerciais")).SendKeys("10000");


        //    ////Tempo de Empresa (ANO)
        //    //Thread.Sleep(2000);
        //    //driver.FindElement(By.Id("tempoEmpresaAnoDadosComerciais")).Click();
        //    //driver.FindElement(By.Id("tempoEmpresaAnoDadosComerciais")).Clear();
        //    //Thread.Sleep(2000);
        //    //driver.FindElement(By.Id("tempoEmpresaAnoDadosComerciais")).SendKeys("10");



        //    ////Tempo de Empresa (MES)
        //    //Thread.Sleep(2000);
        //    //driver.FindElement(By.Id("tempoEmpresaMes")).Click();
        //    //driver.FindElement(By.Id("tempoEmpresaMes")).Clear();
        //    //Thread.Sleep(2000);
        //    //driver.FindElement(By.Id("tempoEmpresaMes")).SendKeys("10");


        //    //CEP
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("cepDadosComerciais")).Click();
        //    driver.FindElement(By.Id("cepDadosComerciais")).Clear();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("cepDadosComerciais")).SendKeys("06033050");

        //    //Numero
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("numeroDadosComerciais")).Click();
        //    driver.FindElement(By.Id("numeroDadosComerciais")).Clear();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("numeroDadosComerciais")).SendKeys("365");

        //    //Telefone Comercial
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("telefoneComercialDadosComercial")).Click();
        //    driver.FindElement(By.Id("telefoneComercialDadosComercial")).Clear();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("telefoneComercialDadosComercial")).SendKeys("1136811452");
        //    Thread.Sleep(2000);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);

        //    //Bot�o avan�ar
        //    //Thread.Sleep(2000);
        //    //driver.FindElement(By.Id("avancarDadosComerciais"));

        //    //Cenario 5 -- Outros 

        //    //Validar P�gina Outros

        //    wait(By.Id("avancarOutros"));


        //    Assert.IsTrue(IsElementPresent(By.Id("formOutros")));

        //    //Vencimento Melhor Dia
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("vencimentoOutros")).Click();
        //    driver.FindElement(By.Id("vencimentoOutros")).SendKeys("10");
        //    Thread.Sleep(2000);

        //    //Confirma��o de Vencimento
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.ESCAPE);
        //    Thread.Sleep(2000);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.ESCAPE);
        //    Thread.Sleep(2000);
        //    //Cobrar tarifa de overlimit


        //    wait(By.CssSelector("#tarifaOverlimitOutros"));

        //    driver.FindElement(By.CssSelector("#tarifaOverlimitOutros")).Click();
        //    driver.FindElement(By.CssSelector("#tarifaOverlimitOutros")).SendKeys("N�o");

        //    //Conta B�nus?
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("tipoCartaoOutros")).Click();
        //    driver.FindElement(By.Id("tipoCartaoOutros")).SendKeys("N�o");


        //    //Campanha
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("campanhaOutros")).Click();
        //    driver.FindElement(By.Id("campanhaOutros")).SendKeys("931 - Desc Primeira Compra Marisa Itaucard 10%");
        //    Thread.Sleep(2000);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);


        //    //Bot�o avan�ar
        //    //Thread.Sleep(2000);
        //    driver.FindElement(By.Id("avancarOutros")).Click();

        //    //Cenario 6 -- Camera 

        //    //Validar P�gina Outros

        //    wait(By.Id("avancarBiometriaFacial"));






        //    //   Assert.IsTrue(IsElementPresent(By.Id("fieldSetBiometria")));

        //    //Ligar Camera
        //    Thread.Sleep(2000);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);


        //    Thread.Sleep(2000);

        //    driver.Quit();

        //    SendKeys.SendWait("^+(J)");
        //    Thread.Sleep(2000);
        //    SendKeys.SendWait("acessoFlash.faceDetect('data:image/jpeg;base64,/9j/4AAQSkZJRgABAQEAYABgAAD//gA7Q1JFQVRPUjogZ2QtanBlZyB2MS4wICh1c2luZyBJSkcgSlBFRyB2ODApLCBxdWFsaXR5ID0gOTAK/9sAQwADAgIDAgIDAwMDBAMDBAUIBQUEBAUKBwcGCAwKDAwLCgsLDQ4SEA0OEQ4LCxAWEBETFBUVFQwPFxgWFBgSFBUU/9sAQwEDBAQFBAUJBQUJFA0LDRQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQU/8AAEQgB9AH0AwEiAAIRAQMRAf/EAB8AAAEFAQEBAQEBAAAAAAAAAAABAgMEBQYHCAkKC//EALUQAAIBAwMCBAMFBQQEAAABfQECAwAEEQUSITFBBhNRYQcicRQygZGhCCNCscEVUtHwJDNicoIJChYXGBkaJSYnKCkqNDU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6g4SFhoeIiYqSk5SVlpeYmZqio6Slpqeoqaqys7S1tre4ubrCw8TFxsfIycrS09TV1tfY2drh4uPk5ebn6Onq8fLz9PX29/j5{+}v/EAB8BAAMBAQEBAQEBAQEAAAAAAAABAgMEBQYHCAkKC//EALURAAIBAgQEAwQHBQQEAAECdwABAgMRBAUhMQYSQVEHYXETIjKBCBRCkaGxwQkjM1LwFWJy0QoWJDThJfEXGBkaJicoKSo1Njc4OTpDREVGR0hJSlNUVVZXWFlaY2RlZmdoaWpzdHV2d3h5eoKDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uLj5OXm5{+}jp6vLz9PX29/j5{+}v/aAAwDAQACEQMRAD8A{+}ysEEd6UKMUoFKB6VzmoEZ7UpBHNKBjp1pcHrigBQgxSr{+}tHPrR3680FJAQc04DJx3pcZwe1KR0x{+}dADSMYP508LkDOBSgH604Ad{+}nvQITGf8aAuOOxp2D74pwHAFAhqjAxjinbenrTgABRj0oATZxmlwKUDINKBngUANIPbpTgvJ60/Hy0DNADQm0U4Lx0NL1Jo6CgBAvB7Um3jjpTyKMZ7UAIeeKMY5pQRz60uc9eBQA3qKbg9qeRjGO9ZuteIrDQLfzbucKTwkYxvc9gB3NAzR6UY/CvJ/FP7Rvh3w/Ym4jhu53UkPFJF5RXB{+}bO7HTjpnrXjJ/bKtvF3iY2dhetoIiRntpn5hmJwCkwI454BHfrjurlKLPr48LRXzp4U/a10yW/GmeIJYLe5kcxx3ETDZ5mf9U2DwcBiM44BHUZbgF/bwXSfG82k6jbxXNqjGDfapklwcbgSRlTlT2PWlzIfKz7KB9aQ4r5z0b9q2114XF3bqrWMR42LkhDvUFs/dYtG3GemPcjJv/2r/wC15mfS41NqjYJdShAAB388YJBwMcj60udDUGfUQUEZByKTHH614t4P{+}N/9rwLHO8bzkKyrbKcMCOSR1B//AF9K7rT/AB9a3yZgZZl3EBklXY56YDHr{+}HFNSTJcWjrsACgjOQRWfZ63bzIGdhDnqHPH51qDBGTgVRJERwM5FGB16mpCucZ5z0pNo7/SgRGPel28cUpGRSbSQBQA3bgY60DnFPH1oxnkUAMxgCmhRg9vepMU3nsKAGbeKTbgVJ25pMc{+}tAEOM4oPHSpdtNx69KAI/wAqRhkipD1pmMEmgBhXHHQUmNpx6mpaaVLexoAjOCO/4UzGD609htoI5oAjOV7UHJPtTyMmmkcUAMYYNNIAFPAyPfPrTTzmgCNuKYRgZxU2OOcUbaCrkA4xzRjt0z3qRlweKYcfWgLjCoB5NFPwDyOaKBloAk88e1L0ODxS529qXGcZoJsIBgDjFP8A50HhvUUu0EmgBOWpwUKM4z70u3sOtOCY/GgBAM9KcBjHenAY96XHGKBAB6UY/P3py/r2pcZoAAKVVJJPSnKo7Ud8UABUD3oHSnAc{+}1OAoAaAeeAaXHPvS9/0oIA79KAEJ74pRyRijPalFAwxyaNtA4OKO/vQAE88UvTmgAk8Vn6zrlpoVsZruVYx2DMB{+}poC1y3LIkSlnIUVz{+}t{+}NtN0clbi6WJ8Z2Zyx/AfhXknjn45RQG4eKXMcakpCmVY{+}{+}OCM9iR{+}FfL/jz44re3chsSLuWSJlaPYWCH/gWR19qydRLY2jSbPbfip{+}1XcaJqSw6JDcXVt9151Awpx6EjPPvXz547/ah8SeLJlsTd/wBnbyY5JI23YyMfdPAz1z7da4K98WatNJJJPOgWUMrRsVcbDn5cdB17CuJ1Ywx3LyyLuUkEN0z06UlJyNuRJHUf8Jt4zsNGu2GrHUI3fy5bXUHEsTcfKCjcq2A2GBOeQMV5t/bk{+}m6g15ChjLffQEyQyAjBBPcHPIPr69Ld14n0u0EkSQqrMuxnExYFfQjGKxtQ8Q20loI1BRMYXfGOc9cYGa0WpDXYt3V9c6hNJe2kzrI{+}BNCzlnBB4YE9ccZ79euTUGurN50N4gKO3zFASdvHPPX8Kx4vEAiKsrpJtPGVK9a2x4ogeUSbg25Nx45zzkGrsTddTQ0nxLqmk20wW4kRLtQ7KhPY7hkfifzq5c{+}L9Wg0uKBJJ5PtDqzxY{+}V8Y2rjvyP51Ui1rS7pYXnB{+}U79sePmAPAOemeK1LbWZtSjDpbxRpyPMdsAd8etQ0aRV9mdR4a{+}KHiPSdCmtrm{+}nneQEN821FyMFSe4AxhRgdc9Oe2{+}GvxNh026dtVv5ruJcqENx5SIeehPzH8Oeeteax3kKwGKS/2x9dsKbQc9eWIHr2q0jaPIuBdT8AgbQTn8RkYPvWZSjY{+}xvhv8bfA99G9leaxZ285YgK7HbKuONpbkntz6dK9r0bW45reFtIujFnBCRzAwP2GCOV/AAeor8yG0PS7gs/2yNhySkpKN79scVs{+}GPEvijwiBJ4W8TyafEcbYSziFvqCCPrxT5kiXC5{+}p{+}j{+}LJZ3MVzbypKpw6ylQyfXB5Hoe/wBeK6WG4WZcjHX1zzXwB4O/bF1XTHt4vGenxIQAi6lY4ZR/tfK2Meo/IV9U{+}BviQnjGwt7zRruzvkkTdEIZBukHQjacbh9MfnVKSZhODR60cY6UY6VkaV4ijv51tp4JLK9xkwyc7gOpU9xWuDz61ZmJtFJ0PtTu1FAhCAB0zSFSBQcjgUoOetADCuaaUINS{+}1Z3iLVRoHh/U9UaIzLY2styYwdu8IhbGe2cUAXCtJgEHpXmf7O/xwtf2gPhZp3jWLTG0KG9nmgWzmuBKQY5Cn3sLnO3OMV6cyYHWgCM8D1pCvX60u5c7cgsOoB6U1ZElBKMrgHB2nOD6UAN9ecYo9K5vXfAFnrOptqVveXuk6myhWurCYoXA4Acchqoz6V4p0SymlbxXaTwRIXMl/p4GxQMkllYfniueVScW7w08mv{+}Ackq9Sm3zU3bumvxvY6yW4himhieVElmJEaM2C{+}Bk4HfAp2Ac56V8/T3nivU9B1jxzc2uo6/HplrJ9h0jSN0E{+}o4IzHFtUsobAyQMnGO3HsPw88RXfivwRousX2hXXhm7vLZZZNHvcia0P8AzzfIU5H0FKhWdePPy2XQjB4mWKi6jhyxe192u5vDPQDgUh5FKkiSLvRw6k43KcivG/2l/g/42{+}L{+}iaNaeCfiDe/D{+}6s7h5bi5spZozcIVwEPlOpIB55zXSd57Fjk008E{+}1fmR{+}xb4c{+}MHx81h/Ecnxm8R2{+}meG9YtftmmXupXc63sYYSMh/e4wyqVIII5r2P9n/x34q{+}Gn7Z/j/4TeLfEmra7pepxNe6A2r30tyY0XM0aRmRjj9y8gbHUwe1VYm59osDzjFJjOQcYr5N/wCCiPxf1/wP4H8L{+}EvBuo3uneLfFOqLFBJpk7Q3AhjIyqOpBUtI8K9eRuHrX0p8P/Dd54R8DaDo2o6nc6zqNjZRQXWo3kzTS3MwUb5GdiSdzZPPrSsM3Svp0puOPwp/8RpCOwxSGRZC8Y/Kin7d3aigdy2OTz0o6ilC807HpQMQen61IBQq8A07oKCWJjBx39acBRt5zk08AnHFAhqipFX3oAHHrTgcnFACFQKMY/GnhBz0NLkD60AJ0ApwHXNIBxinY5oAQj0NKB0oI3Z7UYoGLSdM8ZzThzRt96BpCY5NAGR6Up9KcBxQDG4xRjLU7O2vOfiv8V7bwJp7QwNHJq0qM0aOSUhUdXfHQDsOM0m7Ak3oi78R/itpXw/s281jdai4Jhs4cNI57cfXuePWvjP4vfHPVtR1JG1K6Tz2fdDpNoykRAcgyOfunPcc{+}nGK8{+}{+}JXxpe8ur2LTbqa{+}v5pMXOpzEHeeDhQOAuQcKMjivHrvV5J2bYfPvJGy1xISVDexHLH{+}tcrk5Ox2QpqKuzuNd12a/RzfX8cClyzRQAqpB67j1J46MfTgVz8{+}viRZE02Jyqr887ybcH1JPA/n6VUsLc3siwRQNq1ymczz/JDEfoD264H41sXljpOgobjxPq4mmABis4R37BYx29249B3qdIvzNldmFGWv4WEk4H{+}3EjON31JXP61y{+}s6ZZacrPdPlucBsKx{+}uc4{+}ldJrPjzV9VRrTQdNj063xta4kO6UL0GWPC/z{+}tefyaG19ds91cyagxOA6ElSe4BPJ/CtaalvLQym09I6mfJqGntI8iQ{+}bjj5cnd6CqV7ezzHesQjJHRu1dBdaZFpcCmZ4rdv7gOSv8An1Nc5e3lvG5EOWH/AD0f19hXWmuhzSutyuqTN/rnGPYVqabYQzXCwuxYAEsQcHODgVlWxN1couCQCM59M1t{+}HrS4u73bGoLP3HNEnoTFXEtNtrclJTjYSBnnHpWjBqwkPlpyx{+}7xnPr9D{+}FVNes/7PuXgkBLZ5Kj8v1rBeZoZw{+}4kAjuR/Kl8SLb5XY7IeLZokeAQwuoYksyq34/dzSDWLW6iCFDCwPXoT/WuQffOzTLCSMljsGRgfyxU0LQSHcztH6FcHB9x3/MUuRIFNna21leLkxs8iHLhW7{+}mfX8c1asfEc1g3lyQBOMNKh2t{+}XXP0rB8P67f6BcpPEPOt9wyUG4EjsQRwfpXf6a{+}heMYN7TRWN2/G9f9WTnow7ZrGUuV{+}8tDogufWL1LlhdnxDbF7O5tp2C4NvLIN2M88Hr{+}FXvDHi7WvhxqkV74d1KfS7tWJNpM{+}YCw7Ecj17VyGp{+}G73wzfEL{+}5lyGWTGVYeoPQj9PXFbdl4i03xIqWNwbfS9UPUSswikbHGRg7CfUcf7OOazdlrEvfSR93fs/wD7U{+}nfE22i0TxMyaZ4lgAXc52rKegKNnr6EHn8a{+}n7C6EkSB33sR8r9Q3/ANevxde5k03Ultr1WsLpMPDKrBSM85Rx1B/EGvrv9m79q680CSHw54uumvdOY4h1E5Z4TgZ3jqV557gc{+}hq1O25hOn1R95Y60mDnsKpaNq0Gr2kcsMqSoyh0dDkOvqCKvkYrY5Rh4IzSNk9KeRnGRSbT68UCGHp1rnviM3/FvPFA/wCoXdf{+}imroyvX6VleJ9IbxD4c1XS0lELXtpLbCUjIQuhXOO{+}M0AfnN{+}xz{+}xF4W{+}PnwDsPEvjTX9buvNmubfSrGxvBHDpiLKwZlUqw3tJuc9sEcZ5rH1Px94qs/2Rv2h/hvrWu3OtzfDzXrDTbHWXkbzXgbURHs3ZzhTAxAJJAfbnAAHsPgj9iD4zfBXwwNG{+}G/xtj0q2vt7apbXWngwiQsQJbfIcxsY9gOMHK53YwF7dP2EbPS/wBl/wATfC7TvEjSeIPEt3Bf6p4mvrcu086Txy58vdkLiMgAsTlmYkkmtLk2Pnzxx4R1n4VfB34OeAfDXi3VINY{+}NOo21zr{+}v3VxmRS8dshiRhhhH/pC5GdzCPGcMRVv9rT9jfQP2cfgJqniXwD4n8QabIj2trq1rc326LU42mQKWVVXDrJscY4wrcd6{+}p/jD{+}yZp3xf{+}B3hXwVeavJpuveGba2Gma/axndDPFEsZbZuB2NtyQGBBCnOVrxrxl{+}xH8ZfjP4RfRPiT8bU1W3stjaXbWungQmUMAZbggI0hEZcDOTls7uCGSYWPrX4WlpPhn4SdiWZtHsySTyT5Kc1yWt6hJ8VvEx8PadIy{+}HLFw{+}pXUZ4nYHiNT6ZH6E9hnM8T6tqWh{+}HdG8A6K7vPaWdvY32qFCkcahFjznnaGxyfwGT06nw1qnhnwBocGl2tw1xs5mmihZvMkJZSScesbADsAK8mrVjWqeyvaK38/L/ADPCr4iGJqvD81oR{+}J9/7q/X7jmf2tIxpn7L3xEWz/0UW{+}iSiLyjt8sADGMdMV8Z32reIPix4f8A2aPgfDr95oOheJNBGo6zeW0pWW7jXzT5ZY9flgkwDkFnUkHaK{+}5PjVosHxT{+}FvjPwfHdTaa{+}oWUtm189nJKkROBkKMF{+}TjA68{+}hrwPxn{+}ylpHib4UfDjRtO8ct4e{+}IHgOBYtM8SxWkkKtgglXUkFRuVSDuypzwckH0VUprS6PX9tSWnMjxP9rv8AZh0r9mjwr4R1DwL4i1u00TUvEFtaX2i3d6ZIpZgkjR3C4AwwVZFPX74xgZB/SwLjtXxL4v8A2IPip8Z7bSL/AOIvxki1vUtLuo5bCC201Vso4ersQnl7pGwnzY4APJyMfSem/D/xfafHTVfF9x40mufBlzpy2tv4UMbeXbzjy8zBt2MnY/b{+}OtHqjoifKf8AwSdGfBnxF/7CsH/otq1/{+}Cg{+}g3Xw68XfDP456PCz3nhvUorLURGOZbcsZIwT2U/voz/12Ar1r9kH9l28/Zh0LxLYXniCDxA2r3cdyskFsYBFtUrggs2etem/GX4Y2Xxj{+}FviPwbfOIYdWtGhScpv8mUHdFJjjO11RscZxRfW4W0Pj/wjd237VX7f0viS0lXUPBfw{+}0{+}I2ky8xSz4JjP186SRwe4tx9K{+}7T3B614h{+}yN{+}zFB{+}zD4I1TSX1OPW9V1K8{+}1XOoRwGEMiqFjjClicL8x69XavczyMmkwRFijocDHNKw9OKTAH1FIBrIc8GilNFAFvoen4U8D1oUZNO2jGTQO4m3JPangY{+}tA7U4DpQIQCnhBnpRgAdOaeCPWgBAuTnPHpTqTG7qKcORzxQAEcgUBR260fdpwoAbjnFOUZpSvB9aUcdaCrCY4FKDn2pQe9GMmgYntShcmlC4NL0oFcQ9MUD1pc84qlq2q2{+}i2E97cvshhUs2Bkn2A7k0CMLx54xg8K6XI5eNbgoXUOeFAIG5vbJH16V{+}cv7Svxxi1nVrvRdNuGkjLZvr/ILv2CL6ep9zjgDn0L9rP42NpSTWFrdJNrF7hrl1{+}ZYQMhY19AmSM/xOSeq18PzXD3E7PLI26QmTLnknu2O/f2GaxfvM6YRUdTXk1J7{+}4dUQQwgn9zv6DPc9f8{+}9b2kaPLeeXJLI0duQTiIhS44yPZRnk1k6PpQeNJbiJlhJyFb7x{+}uOnX1zzgZOTXoFlpn2ezWedTawyR7sv8u1QRh39v7qAelZyfLojpinLVmffaxdafa/2foCN9oC7GnP3VXPb2HqfwFZVl4btrVXmnlN3qLnzJb2V{+}hJ6DJ449OTxW3YW6a9dMI7aSG2Y56ZeUHpu75PPHU/y7p9L0fwbbw6prsSXEykSWlhwyRDpvcZwT7HjnvWTnyaLc0UObXocZD4JiksEurxzbaWBuNxLHgOuOSkfUnHQtXI{+}LvHdjBE9h4XsjbWwXbLfTnfNKPd{+}ij2GB6ZrV8XeI9X8aXE9zP50FmzkpbLyzjt9K4zU/DtxIxMwxHjiBGz{+}eP8{+}9awgr3m9TOUrK1NHK3Km7PmlzPLk5y2fzNZps3BLEhf0roLgJpcbxzRhHf5RGBk/U/p{+}ZrInLzsElAiQc4H3vxNdq8jia7jbOJVDlSSeQuD3rsfAOnvLq0K7Q4Y8kDkD1Brl4QoJRCu4cheR29K7L4eh01KL940bZGVCckVE37rLprVGv8UNKFpqUpSOREIG0tjO3HXHbv9eTXkl4PLcqecHoeRivob4gW9vql2GDxiSQBfndSwOOmMgDHHUivEdY0aS3u5V3R7Qfvhs5/Ks6MlaxpWjrcw4JGRgI3ZOc7c96mluWmdTIoUr8u9RtLH1PbNJ5VuG/eTkt6BDx{+}eKvQWVtOFY5mX{+}8eo/Kui6OZJ9CbSbmQvuV9roMgYBBPv7V0VhewTOjCOOzuAT{+}/jO0P7MOn4HrWDFZz21tK1ukM8C/eVic/lmrui62A6xvZwSbuNsiZYe3PUf5xUO0jSLa3PVtD8Tm8sxp2rwGeA52SLnKdOQDyPw9sg4qDxL4J0{+}{+}soLlvmhlbbDdR4zEf7rcH8ue9clpt5cTTJNZ7XAwTAiDeuD1QfxD26/Tt33h3xJbamxSZI1lCeXLbO21ZV6kg5{+}Vs4wT0PXjmuRxcNUdsZKSszgbuK/8PmO210NqegO{+}Eu4jueA84IP8JHXB4PtQLuTQHt2F695pRYG31KMfNC3bcBzx6fXHfPq2q2YsGgjhn{+}3abcf6ppEXEyn70cg6Bwexrz7U9At9KtpH0uXzoptxl0{+}dThQDz69sHrxj8QJ8wnG2x9RfsrftLXWi6nD4f1WU/ZjgYVsqQSNska9{+}uSo98eg/QHSdUg1W0jnhkVw6hvlOQQehB7ivxJ0m8m08oLZmgER82Fsk7O7LkdRjJ{+}nT3{+}5/2Yf2n5NRW30fWrgR3UbbC0pAHJA5J45zx2/PIuL5dGc9SnfVH22wGM03n86hsL6O/tkmjP3hkrnoas9TW5yEZHFIelSN1ppGRigRHgtk0oHHTBp2OKBjNAARkc9qaVB5p{+}MimkDmgDkV8AWlzqM11qDNchpXkjjV2A{+}YjlueSMDHYY{+}mOktrSCxj2QQpCg/gjUAfpVhl5oK4HFZxpwh8KMadGnSu4ojIzzUckayqVdQynggjINTAY96b0Oa0NrEe0KAoAAA49KYeAe9SkA5GOaaUwKBkZOMdD9KYQc56VKwwOlNJxQCGA8U0/ep{+}M9KaQc89aBjXGabtp/WkYZoEMxRQ3WigRe28YpQM9qAOtOHygDvQAbexpQo496cBk5xTjxQAgXb0p2MZoxmlHIoAAO{+}aXHWjGRSqM59KBoAM9qMd{+}1PIxSdRQOwZ4zSZzSjp1p2BjHagY0LT6Tbgj0p3eglsTHWjGKMc0cHjpQIaxA{+}ueleKfHn4hWfh7Tr6S6neKy0qMT3TRnDNI2BHEn{+}0xIGe3zH{+}HFeua3fR6Zp1xdSnaka54OCT2H4nA/GvzB/az{+}LMmoXg0iG7eZY5GubjaTiS4bJyf90HaB257msqj2SNqcb3Z4X8RfG02ua5d3tyPMvZ5CwgUZUc8KB{+}P6/nj6VZXDOuSrXDENNIfuxjsPqOcD{+}fAqnoto{+}ral5rfNs7ngKD/k/5Fehx2Nn4d003F31XkQMMNuI9D/Ef/HB78DOTUFZbnVFc7v0NDSNOsdDtGvNRCuTtNvCxJBPd29vQc5OfQk1rm{+}n1/Uz5m6O2iJU8End06d2x2/lzWMmqtqlx9ovc7m/1UEfAUY47ccY{+}mPrXQR3MWhW8cXk774LuEanCwAgfMf8AaP8Ah9Kzs4{+}pumpeiOutb9fCGnJHDAI9TdT5KkZaIFR8x9SQKzbPRn1i{+}e61S8kuHxubbllU4yAT3bHp09B3zNHjn1PUJprpmk3ktJOxIbOei/ljFdrd{+}K9H8J2Ky6mUs7dFDtAiZYg92P8AQcmsX7isviZovf1eiHPo1rPAPN/0OwQENcEDZGo/idgMsfYcDjpivJPG3j6CeWew0CARW2djXYULNIenBx8ox{+}PoKp{+}PPitq/wARZmhg3aZ4fiYiGBBteUDoSAcA4xx7CuafR106xguL9njjlz5MKHEkvvnsuep784ranR5bSqbmFSspe7DYoym03NlJJmjGN2/5U9cnHJOc/j{+}VKSK3vFyWBVeRyc49O1SN5{+}ouIY4iFHKwxj5VHqfb61BeWDQxqjOCXYk4HHH{+}TXcuxyO{+}6J7Z97jafKReAAQM/SvR/A5TSQ0wmieZsEhXVmHPrj{+}teX2dvIsqmGIs2eDivTfCVlNuhRYZEkYfflU7cd8g/wCevNZVGkjalFt3I9bmW/uHlLiWXaWHdl9sdvwrk5547rcrbCQeA/b/APVXqXifwzFY2yfuTJLEg3vEhI6jBU9wQea8q1SD{+}0LmUFwJM5HQbvyrOlJSRpUg4swNXs5I58jeVb7u7ncP5GqVte3GnMrhQ0YOCAuM/wCfpW2kVwsTQuBcQ9dpPTtkDsfccVXl0zed3LA9Mjk{+}xxXUpJo5HBp3R0Wh{+}ItN1CFYrvdaSLnbMqjK/wDAhz{+}DAj0Fb1/4ImnsBNAsF9GfmSa3OCw6{+}vX615xNozW/76GTcmOUc/Mv{+}I9DXV{+}AvFdxoV{+}qxXAi38NFIu5H5/iB/mMH3rGUWtYG0JJ{+}7NDrWCXTLxeJLOUDOJE4Y45z3z{+}FdjpgTU7dL2Nle8Ti4RByB2k68g9D0wR1rrW0zQPiPF5bkaZqgUeS0BJRz9DyOfT864jUtA1jw1cMmpQOyRsUS6T7{+}MYwSODwercEd6yVRTWujNXBw9DvNOuEvC9tcI3mEfvLZ{+}OBjGD2IPQ9jjsRkv8ARRPIVMqsCgeKfH3mz7dDg4I/oa57RxJdQE2VxiaEBvKbBJHTABzke3XqPSuwsNSju7cw3AWMmQJvA4jkGcHnv1PuPxrF3TubxtLQ8j8QaLd{+}FJf7StAWs1f/AEi3UYMR/vD/AGD1q9pOpi01O3ngdRHKge3bPyvGeSjZ5BXsfY9sY9J1zQheWi3Fsoa/VWj8liBHOMElQGx1GSBnOQR3ryjVdM/sK1EY3tZb/MiQrseIscMuD6NjH1I7mtVaaM5JwZ{+}nv7LnxQk8WeF7eG5u1uLiJdpAwGAHQ4yfp6Zr6DRwyggg{+}/rX5GfBL4qah8M/EtlewTmS14aSJmOyaPPzKfoOfpzX6i/Dn4gaX4{+}0SO/0yYSRMivtzkqD2P45FVSl9l7nHWp8vvLqdf1xSH2pxGOnIo74roOUY3IAoxTgM9uKUgA0DI1JHHSlK96UqCeKMYxQFhpzjpTSvNSU0jn2oERlSTTSO9SEdfpTSOo7UAR4OfejnGMU8DnNNblsigCNl9utNPAxUjcnHT3ppHNADMZPpTHHOR1FS7SAcDmmFcUDI{+}2Mc0gJXOKkcdwKZjI96BiADFFJkDg8UUEmgq8dKWlAJPtSgY{+}lAC{+}2aOM80AZJpw6CgAPANAFLtz7U480FWGnI7flSgZzSgZNOA4oAQ9PWg9Bx{+}dLtyKXFACKnJpelL2{+}lGM0CEx9KMg5FFLjI4oATtQ5wpINOA46VS1C9W0jlkY/JCmTjrnt{+}NAjwP9sL4qJ4A{+}H81ukgF1dBkCj124A/DIP{+}RX5OazqNx4h1iUMzSSM5LEc45/z{+}Jr6J/a7{+}Ltz8QfHdzZROPsenF4FCnIkmB{+}dh68kDPoorxTwtpo0wJdS4E0mHiJHUDIyPx6HufpXPzWbkd0YaKJteGLaLwlG0jsiyw4ZpGAO1vTB4yP5gDsc5N7q//CR373c5KwrxBE5Lbz3Y{+}pPU/Sq2qXzalcm3XcYi37wjox9Pp2q9penLKzlhtjQDe/YgH7orK9vee5vb7KNOxjTT4F1CVDNcyDEEeBhRuxkjv/8AWqWzEt9MdrSNKRveQdFOQevfA/KotQujdFCGOIsIGUegwPpU8Nz9jtBLO5WNMFkX{+}X1P{+}NJNrXqytH6HSfbYtEtFSEtK6KWBC5APPOPXn6D615zqzz{+}LdXD3qulnE3y8ZcdSOO7Hnr0rfvZLmSeOaRAk0uG8sLkQRkZGR69D{+}IrV0HQjEH1C72QQRjKq43HnnJ9z19TkCrilT1e5Mr1NFsc/LYQ6XbxXD2qG4bi0tCpIAxne31/WsiDw9eazem5vpljThnneQbV77Rjr9BgfhXZLZG9E9/c3rWlmOATzPIPT0Qe3P9a4bxB4n3PsiYxwtwgbJZ{+}eDipU3J2juaeySV3sT6le6ZpVmba1ZpnYAskS7Qx9Wb/61ZVvEb/yFjtUUsDgbN7fePTP{+}FSaJ4e/taTK2iydSSWf169a9u{+}HfgaOyuI7i4s4IVSP5FfdknPufesalaNGL6s6KVCVaWqsjlvDnhBnMDXkCW0ZJyJI{+}uB2AHt6V3lpG9tsdYZU0{+}McM0GAxz0G/rz1Oe1ehWngmS9kSeaaS36kOUHTpwOv{+}RVxPCRuddhjtGlPzpAr3PJRc5JHoeK8aWK5nc96OFUdEeean4f1XxPHF5Ud7FaxjJlPl8tjpnA9OuPz7eSa98PrqNZJyrSQBvMLxjlARnI9ffFfdukeBIdT067t1LBZ8KHT7xQdTk9Mg4/GqHjb4ORQ2VwtrGArHCArkYAA/ofaqp4xxIng4z33PgAxNp8{+}25TzYj9yUH7x{+}v8Aj{+}nWtTTdJtr4vjLSMCD8mHQepXuPcV6j4u{+}GF3pO67Fj/oLtieFUBMZz99R6EYPHWuNuPA7Ws0bQMUVv3ilW4PoVPUcV3/WIz1TOJ4SUdJK5xmo{+}FLuwJnVRPbOCRIBnA75Hp7/n61zWpaHNCouo0zFnG7PzIce3avarGaRn23SBzxlZOrdRk46/X9TRqfg4m2e5tIDKkg{+}eDHUe3/1un1raGLcdJHJUwfMrxPKPDWr3NjNGBdZiU4wCSAfTjJH{+}cV7t4Z8XWfiTSo7PVJHuMkxghh5inHbuR7fQV4jq/hkWzS3VpFNFMv8ACvyt15DKeD{+}GM{+}tRaF4theRLa6TytpwtxGNrBuqn2wa7JRjWjzROGLlSdpI9U8Q{+}F38MznULNvMtS4ZZ4uOeD0z{+}nB4{+}la1vBLrFtJqEQErDas0SqAGXtx2I9faoLXxO95BGtwq3KygKSR{+}6uAQcBwejHnnv9aqafFLoc0V9pbPPbM2DFIfmjP8AccHuOx75rC72kb2W8TYsdWOqW8lncHy5ojtQk43Ln5Tnvzzn9a43xQxsrx0uo/tFhPlJ43XIXPBI/PPHsfWtvUr1kuIr1EQRTMG44aNuMjOMbSc9f6Vo3{+}nReINNunhaEswG2Juq9sc{+}vzDqeo9KV{+}R{+}oviWu55VpswtbybSZz9kuUfNtKcmNmAzj1AbB9eRX0Z{+}zR8XNY8AeK7GG3YTafdt5awPL{+}6yxGUDDgHptz6j3NfOnifRR5dvJHJ5cwO1Wb70cyfwk{+}42/XrUWh6vc28yS20wifIYeXk55z/iPwrWUb{+}9E5094S2P228K{+}IrbxTo1vf2rbo3HQjBB9DWsV5r5K/Yv{+}OK{+}KoDomoTp9t24BJwzN1yR3JwfxU56ivriumEuZXOCpBwk0N7Ubd31p4HOMUoUgetWZoh27ck0m3cKkbnr{+}tIRtPtQVcjbjtSbcdakI3AiowCaCRKaQKcaRgMcigRGVwP6U3BA{+}tS9RTWHFAEeOKaVNPoPPFAEbA5ppGaeR/kU0jHOfwoAjK854puCD1qRgDzTcbgc9KCkxuKKTBPUZooEaABH0pQCSaUdRSnGfegQYOelGM4pRyacBgUDsJjHTijB9aXg04CgYKMe9KBgUUo5oEJR0pSMUdCKBCDoKUdqOATS9hQO4ADFKB6UgpwwOaAEZgiEnNeBftZfFMfDb4d3YgcR31/mKLnqxByfwH8/avc9RuBDbu2Rk4VeO59u9fl3{+}3f8XD40{+}JsukWE26w0jNijB8h5gcyvxxwcL/wABNZzeljWmryuz58uLldY1RrmR5JFJKZHVnJycfiT{+}JHtV3VJPJcRsoE0h{+}crxsAGdqj0A5/Kq{+}nqllB5xjx5ahYUPQnB{+}Y/zP19qoTz3ErCeRmLvlY/U5PLfj/hXLu/I71ojS06JDOkUaY3nLyN/CPTA9{+}P8A9fFtj9sP2a1{+}QZLMw4yO5J9B29/pTLdVsLX92SJSPnkznYg6kn2/U0spj061WWTPny/MV6bEA{+}VPryPxNRd3ujRLRXFunt9Ot0AYKQMk7ulXvD9g19B9vu482kZAhtG/5bSH7uRnpnH6dsVi6DpMniTWPKlZEtYvnmLcDthfzPJ9M{+}lel6RppkYXM6ywQJkRqR07Zwe55x{+}HpW1uVXe5HxPTYqWWlzNcJbB40PzSTyZ{+}VR3OfbB//XWxLYLdwC4kkddItjuCn{+}Nv77e57D0x71fstHazt0nuV{+}RiRb2pXiQqDgn1Rc556nk9BnM8W6{+}2l6etqXS4kdv3cbP{+}7LH{+}JmHbP4n{+}XHObk{+}WJ2U6aj7zPN/GOsTahMqxRh{+}R5Nqik5HZmHpnp6kdhUHhb4ez6/qAmvJN8rnllw2T3xjv61uaPocurXf2iWR5VkHySOu0NkfMUTsOwLdu3GT718NPh8JUjZoSQcfe{+}8w6e3H5VjVxHsY8sTtoYb20uaRneBPhV5EUaWlsUZWyZiu7p2z3r0vR/hu8{+}rwy3SO6RKykAdTkHIJHH8{+}DXp3h/QGtY02xohUAcdh6V1VjoxkVX55JyPxrwJVZTd7n0EacYI5Cx8H2yKAIQqgYJcEkj6mtaDwLvRBbw{+}Rl95ZV2gnGBz{+}P45ru9O0/orREjoM{+}ldRZ6cuc7Dx0yOlQotkTqqJy2keHUsLe3SPBKf3gAMnrnP4Vp6n4dS5tDE3zYHOBx9a6qHTV3Kyqpbp09KtzWCNHtYDJ5rpjS0PPliLs8N8UfDy31eCRGjRSylc7eo9CK{+}aPHnwwn8IyPG8HmaS5yjlSfJYnrn{+}6fXt16V93ahpwJOBjjOVrkNe8J2{+}r2c0M8SOrDoRn6VKbgzrjUU1qfnxq3h8KFWaNnTdlWP3kb1B6g06xN9pdwuzN/C6BSH4PfHTv7{+}2CK9v8AiD8MB4VOxoy2muxMc2M{+}Tnorf7Poe35GvLL7Sm0yeSGX7n8LDtXQp8xbpqxRu/DcGpmW5sURJcAmKRPlOc5DA{+}/9K8o{+}IvwqKCXUbCzkVkOZoI{+}SV4yVPHQ{+}uPwr3XT8XEyOjIlywKuSMrKMd/U{+}1adzpTTBvlQGIbXCjLJx1AHUf0rSniJUZXTOGvho1Vqj5I8OalJpaMt1IskKNslU9BnhXwex9exxnrXpOk6o04LLlkcAOrfd9sjn25HQ89DUvxM{+}FDoJNZ0aNWOGM9vFyrryWKex7r2J9q4Twvqr2ToGkcxMAEdjhl9Vb3/z617kZxxEOaJ4DhOhLlZ6fLCLcSLcFJbKY{+}W8bgAj1B9COuenGe3Oe{+}nzeE9XNlta9spvmgdshivI/qR39agtPEFtbqY70GWBhscY5A5IYeo54PYjHIxW5dWxvrNbaOdJXVPOsLndww6lQfQjP0IrG7jpI1snqjkNb0yOOKW2WQyCdRJEzkkFwMr/ADKn/eHpXnNxJEpW7i3wNGxWdOu0k/MPzAP1Jr07U5IdY08ysxSUqQ4x/q5AecY6A/ex/vV574mt/s1zHdyriO4zDd7P4WHDN{+}ob8RXXRd1ZnFWWt0d38IfH83gTxdaahby4fzElhkGQGCsCf5YwfUV{+}w3gfxVaeNvC2na1YMHt7uISL7eoP0r8NdNuMu9hKyxzQnzIX{+}mQ2P54{+}tfoB{+}wL8dW3t4J1abAkO{+}0d2{+}UHncv1zj86E/Zzs9mY1F7SF10PutV5FOwO1KBn2pQuDmus4SMgHoKZtBqfHHSmMuKBEJTmoyvNTMCaYRg4oHuRMB{+}NNIwcU5v5UNg/WgBnXtTSMnHangYpGzmgREy44pBjHvUrAGmbfUY{+}lAEec/SkK5HNSEcU3GOnSgCIg0EcH0p7LwabyBzQA0ADrRQfpRQO5exgf1pcD1yaKUgYoCwADrS8ijHFKq56mgoRRT6OlL0oJbDFAzQDjrS4x0oEDUgGaX60DOKAF245xRjPNHagelAAOaacucDge9KRztzRLKttA8jsFVBkk9KBnjf7U3xbi{+}EPww1HUYpVXVbhTa2QPXzGHL4/2RzX5EW5l13VpbuUmWMMxLE7t3djn9Pxr6k/b/wDim/iXxhb{+}HIZCqWzHehPIPQ8dsc/jn2r5lVlsdLVCuM5wq9QAf6n{+}QrklK7udsIW0Y3USbh/K3BTIwQAnJAyTj{+}Z/P1qVGWK4ZI8tIoChj6{+}ntgcn61k6WztI1w{+}FMZKgHpn1/T/x0eta9ko{+}0QwBd0svIcjnGev1z{+}tQ1y6HSnc1II44bcvKu/YQ7h{+}jkfcTHpnk/gK5i6vZta1bZE5ZASQxPUDO5z{+}R/Wr/AIhvhHvtV4jiDb3GSS3A6foPoKoaIqNdRjaEUphwpwCB2/kPwqqUftsU5X91Hp/hrSI7fTEWU{+}TC37{+}WQdflHG714PT1PNd5aF7mSzt0jCM/z7JQSsY5Blf1AHQd2PoK4vw3Zi48tZ8yHPnyI3ORnCp7kkDPsM9q9W/sxNE0jN4yte3eJJnfk8dIxj{+}FQQT7kDjnHLXqcqsjqpQTZia5K9zeDbIsdsibVct/AMZ5PHfJOeSfpXl8NxFrPieWRHjunt03eaykW9v/AA5A79uT78ei{+}MvFU{+}r3b2sUotNNjIR3UndKQfuqP4jnjjjJPetLwjo/9pvBbWtr9ntY2Hm4OTJJ15PfHT0rD{+}HDmludcI{+}1mox2O5{+}HHhh9Q1DzCWmVOBLL1b39AK{+}mfCeipFHFvUjHoOwrifh/4aTT4YwQB34xzXsei2mCmAOfWvArVeeWh9NRpqnCxtaZYZ2gDg{+}vauq0/TFBCf0qppkI{+}XAx3zXU2ap8vy59xSppM5q07C22mrHjIAbtkVp28A29CCeeadFEDgHhT3NXoIwPY/rXbCB5c6l0NitBgD1GPxp6W4Jx3X0qyqlTgHHHYf8A1qaSVbIYn/gNdNkc25lXtodvc4PUDrWTdWwAJIBz3xjmujmOd{+}TnJznpWfNHlCdoPsKwnFdDeEmcR4h8OQatYTQyxhgwxgjNfMHxN{+}FsvhyQyRZOnOfkYjIiJ42n0HPB/Cvr{+}86n5SMdK4vxdpseoW00bxhw4wVZeDXE5cp69Kbe58MOtxomobJUG3djaRkEeldVpepQz4eGVY5AOPM4Jx2J9vWtj4i{+}C10{+}SRYwRGW4BBO32{+}leewq9vdhQ/lygbtjnhse/Q962upq5clZnTXNibmbz4jh3fDJ0VyB3A6N7/wAxxXmvxF{+}FhntLnVtLikuEibdc2apiVBn5mX1x179OODgel6Jdx3QVCgYYx5bD/DuMnB{+}vrVu6aW3mbyn35IVZJGzvJ/hb0PbPQ9{+}lXSqTpSvFnFXpRqKzPluxhMskmnSTpdXMMe{+}2lXjzEPVWHX1/XuK6vwJqPm2MmnSTGONX3R7yN0bA9j25A/n0Jrc{+}K/ws/tCy/t7w/G8GoWcjma0yAxxyQPQ9ePf8K8vtNXF2kd/GzRurCO4{+}UKVYcK2Mcccfh7V7sZRxELx/4Y8CUZUZ8r/4c9C8TWBsrldT8s2tu58u52DKA9N3t1B{+}hPpXBatZx29zdWs0fnQTJ91TlUcdDkdiOM/4V6p4c1KHVtJmtLlRcQyIY51K8{+}XyA{+}PUZ56dx3zXmusaHLbzXWnXO5bq0AeGUjl4{+}2fXGcVVKTTs9yKsdLrqec3vm2UysVYXVueGYcnb2x9MH867PwF42uPDviGw1OxmMMiOJI8How6r{+}v41zV/5sllKUZhdWLfvQWJLpnhvyPP0rM0{+}4MMrwhsrIQ8LHqrfw5/kfpXfOPPHTc4U{+}SVmfuX8D/ihZ/FrwDp2tWzfvmjCzxk8q4yD{+}oNeg7D/AI1{+}bn/BPj40/wDCO{+}JZfDGoTbLLUGLorH/VzY5AHoQuePQe1fpMo6f0pU5XRzVYcktNmRtgUx04qYqfSmsM{+}1amJWYYqNgBz2qy4x2qJlA4oAgZMn1FMOM4zipSKYy8ZoGMpCMjilz8tJyBQIaRk4zzSbeOTTxTTyTQAwgmmnr0qTPHFM780ANPJxzTCO1OYfNxSkZIxQBDjPr{+}dFObg0UAXR3p3WjHFABz34oLFH6U4dPajHOaXrQSwAzSYxTl6UEc0CFIzQenFFA5NABjJxR6igjvnFOA5NAxuTjHpSjtijaO1OA44HPvQA1Rl2/AVynxJ8UW/hjw1qeoTsFhsoGlZh1L/wAIH4/0rpp3a3BZnVMjOdvT3618s/tx{+}OY/DHwvfTVm2X{+}qyeXGpfBVBn5uMenP19qznLlia043lY/O3xXqh8ZePNX1u4n3hpDtLryx4zjn0/rXLa3eF7gwLxghMKeP89KvXjrYwNKuCiqdnPOff3/xrE0ovNNLduNywkY6nc56D{+}v5VhFX959DtdlojV08pEuJvltLdC0sndiecfXoKvQ3r2di{+}ozp5M1xwuOPLjA4A/z6VmQI1xdQaefubt0vQ5x1J{+}nP61T8RaodRutkRKxkiOJR2UHrRy8zsNNJXKtxqPnyqWTKK/HvgevfGQPrn0rpvA9k19fidyVtk{+}dwe6AgKo{+}rY49jXGRvv3JGflYhEz6ZwPzPNe0{+}BNHisbaCViGtY9rFc4DPg{+}WGPbu59M1rVkqcLImknOR6T4H0hUeW4vSimIG7ubh/uI38IOOygZx3A9a4r4nfEue7ikaGQ2kdwfLhGzEhjGeSeuSSWOO7AZrZ8QajcSQ2{+}mK7Q2UuLm4QHAkHBUHvg9eey9q8Z8Q6o3iHxPdTxLmKF9sSgcEg/ljPP1{+}leVTXtJ80tj037kbLqXfDVlPq2qxR4LyggMx52DuB74/XNfTXwx8IrYRI5jy45JP{+}NeW/CPwi8I8yYZkYhiSPXqa{+}ovDWjCOGEBFUYAPoK8/GVrtpHt4OjyRUnudV4bslVUyAO3vmu409MbeelY2mWarGu304rpLCHcwGMfSvFPUeiOi0wYQNyMnp2rpdOBKgE9{+}lc/p8fK9OK3bEfvABzg9q6aZ51bU3Lc4yOoHpWhbkuAMD6ZrOhBrRtcg/Nge/pXowZ5MiztDHbnlenFRyoRxjFWEIPfb25FMc9eRnHGe1b2IuUJTtJUgY9CKozkkHaQBjk1emDOenI9sVWkVQp56c1zTuaRMa6UnIyD71z{+}qW6ujbs4HPFdPcgKT8oz2FYWpIHJHBAPB7V580elSep498QPDy39tONgzj8a{+}cvFGlNaXh3pvZeq9M{+}49DX1/rdmLmORWHzEZIFeGePfDIM0rhMEDOQP6VMJcr1PRa5onkmkzpbzmdHKxnCukg6fl6V0l7bGOL7QoM0UqAMQcq49R2BH68VzF5byWt/tXCy9m6A{+}x/x/A1qaZfm3YDZiFwVntie3qp7HrXZ5nI/MqvdPayBpJhLFgAXJU7Ux/BL3Hs3OPUjpwfjz4dRia517SrIlpFKajp0WBlT/y0XqMd8jofY5HW{+}Kr1tCn80fvYJV4mjO3cnqfQj/HrWT4d8app99FZ38i/eCw3cQwkiHGEdB3x/d4PUDNddPmh70Tzqyi/dkebeGtYk0S7khQN5nG0PjLD0xj0x{+}B{+}tdX4g0{+}HxBpFvq2nIfNg/wBYrEk5zhlb/H3GfbovGXw{+}s9TZbqz22V5GytCRwkyEdOB94Z7dR0HauQ0nULnQdSZbyIxwSkpcRn{+}Ejgke3t6Gu7nVRc8d{+}xwOHI{+}SR5Z4jgFhrcV4U2Wsq{+}TIPUHIyf1H/AfeuOu4JbKaRQpzA5Ug9QOeDXuPjTwzHdyywxphbgFADxtk7Z9iRXj{+}rRMrRzujbwnlTKeSxXjP124/HNelRqcyPMqwcWdD4I8TXXh/WNO1qxl8u4tpkkzz1ByG{+}nSv2l{+}A/wATLX4rfDTSdbt2xKYlinj/ALsgGCPzB/Kvw2sZTpl2Yg48mb5onHTOOPwNfb//AATx{+}Nn9g{+}Mh4WvJfLsdWby9hPyxzj7pA7A4x{+}J9KbThO/Qya9pC3VH6XD86jZRzTweuO9DAHg1ucJXfnOahYcVYYcc1E6jg0AQd6aRUpT2qJhjFAEZ44700g5J/nUuKZjGaBjeaaaec5xTWHFAhMH60w49KkPQUzoTzxQAzmmnkelPb0FMPANADD164op23Pc0UAXQNvvTvxpBgGlHIoGw4{+}tH4Y96XHFHJNAWFXpQTg07OBTQc0CFpTx070AUoGTzxQMbtpwOOtLx70UAA4Jp4GB/OmjNK54x{+}H1oApXWZDz9xefXLdvy5r83/ANv7xhHr3xG0rQopSy2sWWx6uwH/ALL{+}vvX6Ma7OLWwmYcYQkMPXBH8zX5AfH7xavir42eLdVMvmwW0rRwuoOMINi8j6fpXNV10Oql1PIfE168dxLArgrDkluuewH86dCV07SoYyMMkW92zgmR84H4L/AD9qoRBJp42uQ2xm86UZwSB0UfoPxqfc2ralHG4AijO99vTJPIH8hWlkrI0WrbL9pKdO0xZOl3dfNhOqxj/Ej/Oa5ydiWMhYD5SoVPQcH{+}ePxNaGs3pmI2/KWACqPuov{+}HU/lWVdsIo9qAlioA7kAHp9SeT9KqmurFOVnY1PDWmnWdegg{+}7HGMu2MbfX9P1xXtKyLAltbbwttbKTIQAMkcuxHucL9Aa4z4caT/ZmnSXk2PtUgJPmDIBz/QAH8BWrJqitCyAbUYfaH3HJKg4RT{+}Iz9Aa4sQ{+}d2R24ePKrsb4i1{+}VbC/u5HdZrk5LFumcjH5cD3D{+}tYPgfSvtl8jOCF35Ix3/z{+}uaxvFGrSXd3bWmclSJCAeenGfpyfqTXpfwu00nY/TGKxmvZ0vU6aT9rV9D3f4e6LHbpGCi4wOnP5ivctB08Ki/L6celeZ{+}CrTaIwmMkD/61exaRbiOMZPK8mvmKmsj6mm9LG7YW2FUBc{+}tdDp9rgZwM4xms7TojjcQM9jmugtVVSuORnrWSi7lyloXrW3BbhcnpW9Zw7MYJH1rOsTlgwH4VtWq85PU11U4nmVZdy/GmUzjPfHpV{+}2VVBByOwPvUEChTu4GOlTySgdRjjk9q9CKsjzm9SUkKuRz7U2QkKMrweetV/NyAc8fXFKsxdfr0HaquhDpkLR54OOKoTLsJAH5VqxspGGOSPeqtwiqDzyex61E1oEGYN0jEFdvOcfSsS8gYghT19K6WdFk65YfpWdeQbV6Dj9K82cT0KcrHHX1qz7gOevHtXEeJdHFwjAISTznGcGvTLy3XJIHQc81z{+}qxxMCvR8YAzWDielCdz5b8beGdk8sgXHPp0NcpCqTt9nmkENwDiOZiNrdtrH8sH/I9/8YaVBLDKoHznPLelfPvjNP7PuJCFKjJ7c12UrtWMqndGZ4gmkhhmsrhAksZKgSchTjkH1B/{+}uK8Z1i6bTYpFt0W6tg3{+}okHMZz8yA/qD9etep/21BrGn/YNRl8iTbsgvmGQnor46p78kfpXinjye/wDD2s3AmDQTxBRMo5DofuuMcMMHqOoxXs4aF3Y8PFzsrnsvw88e2V1ZC2vLnzrKSQKryjMsTntIOTg56/jWz4q8JfaZbeGeSNRKuy2vtwAfP3VfsT6H8Oh4{+}abbxFc6JfC/twsto6jfGTnC91LDkjoQR0H6/RXw58f6d4j0lNOuiJbKX5GSUhmQkZ4Pcdx69RzkUVaMqEvaQ2MaVWNePJLRnP6ppt3PaSW08Zt9YszskV1/1i44Pv0/T1ryzx3poJkvYo9pmkJljByPM9c9sjOD6nHavoLXLA6XdQWVzcvcLtBsNRcjdKnaJz6jsx9gexrzrx1pELKbqB0WKdtssSrkRuBggr6Hr7EVrRmk01szOtC8XfdHhBj8{+}F7ZfnaP54gfvAdcfnXSeEPEV94b1DTtZ06aSK8tpQ5ZTkq6nIOPfFYGoQ/Y710bK{+}UxKleSq9156gH9Kl0{+}4VLuRSA6SDehJ{+}Vu{+}PavVkuaOh5a92Wp{+}53wQ{+}Jlp8W/hjoXia2kV2u4F89V/gmHDr{+}efwIruzyK/N//AIJsfFx9D8Yal4Hvrgta6oguLJd{+}VEoXJGOxKn/x0/3a/SHrRF3RzVFaQxhnrUDjFWTzUUg61ZkVjk54IpjjIqVziozntzQBF6UEDmnNnpTSOOtAxjKM5/lSdqfjBprcnFAiPGTSMB0708jA4OaafWgBoXrTSBTj9aRuhOPxFADePeijBPQ0UAXdmT1pyrkA0DkU4crQAhXikA45OB604dKOhoGJtBAo28{+}1OpV4GaBCgYo28/Wj9KUjFAARim45zTqCMUDFAAFNKFjkt8vbHFPAwRS7cnHSgex518dNf/4Rr4c61qX3RBay49MhcgfiRivxl1e9ZLXUZrjc0kr7nOODznHvk/54r9X/ANszUJk{+}GaaTbz{+}VJfySeZwP9TGhkfr7AfnX5DeKrh7nUYrASZO7c2Octk/y5/Oud{+}9UsdULqnfuUELR6Y9zKczXDfL7KPu/hnn8BVyCBreykU9SvmynuMdB{+}oH1NQn/AEu8SH/l3txwfp/jVu/mSJUzt2AhtpPDYPGfx5P4Cm9XYqOiuYt2s2/c3EsnzFfQHGB{+}PH4D3p{+}k2bapqUNpCC2Xy7E4JUfyzj9arSztO0sxOR97J7nt{+}pJ/Cup8C6U0kMk7Hy2u2Mfmdo4EGZH{+}nb8K2k{+}WNyIrmlY71bqJLGC0TbGHjMrsvQR8KPzGSPYisi9dPLnaUFIYV{+}0TYHOMDav/AHzj/vqrFuol3zMNyyP58qkcLCmFjj/4EcLXM{+}NdTdbH7PkCS8ZpHI6lQe/1Yn8q86KcpWR6UnywuctpkrXusee5JaR9x74Hp{+}HT8K{+}h/ht5cEK546Ee9eBeF7IyXanA4PevZNK1tND00zsocxrjFPFrmSSKwTcPeZ9J{+}HPE1jo9m1zcTxQxR9Xc4x/9ar7/ALQehxOsNrcC7JydyMME{+}vp17V8z{+}DvB3ir4wXbAyy2ml5wXK5XHoBnmvb9I/ZXgt7ZAuqToMcqUAO73rwqtKlTdpPU{+}gpVJyV0tDv8ATv2krMZXyTtXjeTj{+}fb3rc0n9pvSpJf37BFH90jGPrXiOpfs66rayPHDrSyoW3bZY{+}3PA5PFcnrfwm1zRy8kcEUg2/ftm4/75NZxpUpbSNXVkt4n21pfx78NXKDGoRI56DeM/wA66vTfi7ok08Fs17Eks4zEAw5HtX5gatBq1kwRrZuBydmCPetHw/491LSVtUeV0KS7lMhyY/Xb7Edq1eGlFc0ZGMalObtONj9XrTxhZzFSlwrB{+}hBrZTUFdMqwI68dMV{+}eHh/4n397aRNHdSQ3FpKzkhuo2kqPzJ/Ovp/4X/EObW1YSS74oykfB6sUyawU5rSQTw8HrBnuougFbIx6YphnUcFvmPYVjwX4kAII2kdzTzccZyTWnOcfIai3hVD0z2I7VXvNRWHl2JwKyZ9SjhjYkkKBXi/xQ{+}LUml2GptbTqpi2pGP9rqx/AHNZzm7aG1OjzO56vf8AjfT9OSSWWZdqIWyWx0OOteaeIv2kfDumoUWUyztnaic8Y6mvlvxp8TNQ1XyUxJJCA5EeTh8scZ9hmvO7yHVrx/NKvI55d1B4P1ohTcviZ1uMKeyufVmsftS6aZvKjGwEbhLJwP5/zrkdX/aVimjLJKjY4JA6{+}uOa8EtPCerajKzPGuc5LMMD0x2rprD4Oz6haq897hepQPgY9MVo6NJatihOo9FE0PEf7RVxPGxjYnGQWx09s9D{+}IryPxD8Zbu6ldLyEupyQ6YGfqK9zsPgro0UBF3tkcKPufMSfxzWH4q{+}EHhZIJJEsgz4yWHBJ61tSqUYu3KZ1oVZK6keMReJbfVLRngkBweUJ5XHNcneeMbZ5F0rXkabTAxEF1Goaez3ddoP34z3jPHUgg9eg1TwkNAvLh7ffHbsc7G7c15f4qjLuT0y3AFe7h4QbvE{+}cxU5qOu5sXmh3mgXFvYOYpbe5UyabeRnfDcA9EDHqrcjB5B64OcaPhXxB/YNxE8B/cE{+}VPayEhocuflbIHGeh7Z9RWR4P8app2iz6LrMDan4blmBMAOJbZyP9dAx{+}64HUdGwAexG94j8MyALr{+}m3C6pazZzeRqdl4h4YOp5SUDlkPXkjPU9M0vhkcEJfaifQXg3xfp/irS49J1QsbKVfLjncnzLWbnhiO3r9c8jNZ2veDJNDvLnTLwq77Pklydl1FxjH{+}2vp6YxXinhzxNJ4d1E3lu73FnJhL21J5UDOGz7ckN1BB9efofw94g03xtpSabqFyJHKBrS8I547exxwR0IHHSvGq03RldbHs0pqsrPc{+}Z/HGkyWV65cYlgwrYGdyY4auQdntrcGNmSSF9yt7ckD{+}f5ivor4i{+}DZroTwXEZXULVSY5ccXMWeeR1I4P5nvXz7qeny2zMjqSCMBlPyt6EH8sj1Ferh6inFHm16bg2d78M/F194O1/SvFujyGO{+}0m6iuPLBIG0H5gfYjI/Gv24{+}H/jC08feDNI8QWTrJb39skwKnjJHIr8GfAGopb3yJL{+}9hBKvHjllPDAD6E1{+}m/wDwTs{+}IFx/Z/in4f6ldGSTRrnzrISHloWJBI9uFP/A6v4ZWOaouaPMuh9oA802ReOtPApGGe9anKVXT3/SoHBz6VbcYqF0GKAICOlN2g1Ifu4pjDbQAxh/k0hGehB{+}lPPrSDjNAEW3H9aQjAzUhGabigCNlzg03aR1wam6A96YTjjigCMjBop3HrRQBcC96djNIBxQBtNADsbRTR0p45FMHzGgBwHrS0uOBigryTQAEYpRnJpBkmgDPfGKBgMk9qcBnpR0pVHegLCjrUgH40wLT5WWCB5G6KpY0CPjr9t3xMtqlxIJSkdpZTwAZG0yPER{+}vmKf{+}Ae1fllDMTLeahJ8zFisYJzn3H4CvuH9v/wAQtLp0UUblrjUbpiCp4x2B99pXp0r4oCBruCyjVTHApyf7x68/p{+}lc0d5M7baJFi0AstPd3UAv8xOORx/nFYup3b3ThAv71xyqn7o6Bava1fZ2oW3KhGAP4j9P1/KqtgRb3AZ0zKTvO4cgjOB69iTW1NfafUJv7KFNksk0NnCOXbGD2xx/Pcfyr0rS9Mit7G4CtiIFLKMDoFX5pG9{+}ef8AgVcr4U08u82oSKGVPliz/EBx/wCPHA/4FXbTY03S47edVRo/lZT/ABsSGb/2VfoDXNiJ9EdVCF9WU47nLOrq0ayMJ3B6qACI1/4Cpz9WPpXnuu6gl7qN3IOVj2pEvoqtXbeI5DZaPeXDkm425ye7E9fzx{+}FedW6mS9ZMYG3kfgP60YZXvJixTatBHc{+}ENN82WNgBufjnPWvT7Xw0lykVvKNyD7w65Nc74K0tsW{+}AACufxNejW9lPFMGK5HHzY4FcWIqe9Y78NT91NnqPgrVrfw/osVvboqR7QRhsEGte7{+}Kj2seEZiADkjtzXm7LILZWYgNymOueK878e{+}I20iORQVkcpho2JBQ{+}p{+}leUqLqzsj23VjShzSPSfEvxzlsV2y3JAflAOp9sdc1iDxr4z8SRCWw0id4Tt2TTkRqwPAxvI6{+}1fNk3xGm026eaAC8u85W4mXlfYDpipLX4ua7BZG2k1No4925QTll{+}YHjuMHtXrLL5RWh5DzKnd3/AAPoi80vxpFF9pu/D01xEVGWgZJ8f8BUkkY9qTRIrDWFIktVWVThkI2sp9COoridI{+}OfjD4Ua7p8HiK1YfabCG6jhuwYzNDKm6OQZ6qy7SDXsPhnxVoPxqRbi1KaT4kjDHzGUKJFB4Vv7wx3rmrUalL4lodtDEUqy9x69mVbTw5bWUvmwoyKVwyq2MgHP9T{+}det/CjUV0y4GZMq0hkKg4{+}YjGf51gab4flv4DBLBsvI{+}CvY/T1FXNN0u40HUI90bBCeozmvLqao9KLWx9S6HqZmhQg7gRn1rZeYgEgn8RXC{+}CLrfbRAcZXr/APWr0Ge3H2ZWzg1jGTaMKqUZI5rX9RMFrJn7pHPavlr4izR6vqFzbhS4dzwnf8a98{+}IN{+}1raS7AT2ANeOaToE13dXF1LEGAJbcaUHd3Z0q0UeZa1YW{+}jWH2mdVhSNQFwvPsBXIRad4p1adHL23h3TirFJdTcl2UY5ES8jr3xXoHj3VbXw6v9t6ttKRNjTrAAl55Ox29/b8{+}wr50{+}OzfELSTZa1rljNpVtrCSTWqSMcqmV7dj04PavawtF1dTzMViVS3PUdRhm0fRby4ufHVjttyf3aWhDSEEjCjzM/wmuR0n45XtjYxI91Z3jEbUjDMsm7HU7hjA6de1eAaz4itNT8P6Ja2unXFrrNuZ/wC0b97wyR3hZy0ZWIr{+}7KqSp5O7rxXffDP4D6t8TPBWta/HM9uLCUQo3l5SQ7dxB9McfnXr/UqcV754izGrJ3gj2rSvjnFdtHFu8uZhgh/lJGTnHPPGOlbi{+}O4r8SLIxdgueTxmvi5b240W{+}ktmciW3crkMSFIPavXvh/rVzr0UcKK67Ttkkx1H8h/n61yV8DGC5obHXh8xlUfJNHX{+}MZBdRyGMncema8U8YWjLCNqgDd0x0r6ZufBEyaHNNcRqpI{+}Uda8J8baa0dkXQAlZME{+}h9qeDmk7GONhdXPLEcgqueA3{+}NdP8O/HNx4Q1KeJ4vt2lXYAubCRyqy46EEcq47MOQcVy7r{+}/dcYODgVJEuJ0ZRgMMg{+}{+}a9qSUotM8CLcZXR7J4i8JW0{+}mReKvDFw1/pU0ebiNT{+}9tHHDLIo6MAoPHB25HBIWHwN4nFni18zyFc7kCHgEc/L7jqB/jXIeA/HOo{+}C9du/sZR4J2AmtJhmGZc8q49CO4wRwQeK7DWdB0/ULCXXvDbMmnlwJ9Omb99p82fuNgcocHaw7Z6HiuKcbLllquh6NKet1ue9ad4gtfH{+}mQ2N4Fh1ZP9XKvG/0Knse/I788E48E{+}K/hC50DVJZola2DEjZj93v7gg9uvB6fiKuaJrjxpHKHdZY8bwpwR6MPQ9Qfp6EV69Z6npXxN8P/wBnaskct8YfKW4bP7wYxz3B/P8AGvOinh53Wx3ztWj5nyZYzompkmNrV26mE4AI56H/ABFfW37O/jSfwP8AFT4e{+}OVmH9l6oy6JqUrvjcx{+}Ukg8/daFs{+}qmvnnxz8NNU8EakIp4z5II8i5YZRx/CpPTPbNdf8ODLrvgrXtEEbNcWUy3tsyZDK6o{+}OnTK5H1A9K9OTUkpxPMUXG8Wft3BKJUByDkcEd6krzD9nP4hL8Tvg34Z14tm7e3W3u17pOg2uD{+}INeng5UGtU7q5wtWdiJ1zULrirJHNQyjOfSmIrOM800rtGKldeKiZc8mgBgxikJHPFLjHSkzz7UAIeOlNJ6045NJgYwKAGbhTR605hg0YyfSgBpGaKBRQBcAOef0pVGG9aXGTShe2aBiGkwO1OCYGRSjp70DsGMClHSjpzQBk0AHSk6dqdjn3oPFAhtPB6UADNO2jOfagdxyjJFUvEF19k0a8l6eXE8n/fIzWhGAa5j4nX0WmeB9ZnmdooxbOpdDhhuwuR{+}dJ6IErux{+}Vf7aXib{+}0PHlpaAh49Pt1IiJOPMdVA6/7Cxn6mvnSCM6fFI0jhZGy7leu3/6/QfhXonxo11/EvxK17UJ2BSOZs8lgNvygA{+}wGK8u1a5NzMIgQjMfMkJ7emfoP61zwXMrHe9DPuJ2lPm7CshJ8tf7vv8Ah/OrlhZPNN5MeQ4UDJ7Z6n8ADTbKEbHvXxsU7Ioz39P15rY8P2sjeWQMT3B3l84wucL{+}Z/RT610SkkjKKuzstBs1jt7eFFBLuqqoHPBwPbGTn6LTb26OpeIBCMyRROzMynO585c59ug{+}lWhcGxsrm7UbjBAVi9CWGxePYEt{+}NZ{+}nW40qyEsnySMfvN2z0/Xn8DXlP3merH3YmH4yvNzG2U5jjIDY5Hf/AAz9GrnNN/dnzGXLN8oHrzz/AEq9qEhurbzOS0sjSbcc44Cj8sflVO1njaSOLhjGSfzr06UeWFkeVVblO59B{+}ArNLhLUrwVUD3H5V63a{+}HhJDuYYLA5wK8t{+}F6gQRDIII3Zz/KvoPw5brcQRcfK2DkjIr5vEycZM{+}qwqUoI8y8T6XqmmaaDDGGfaSGYZB9OPX/Oe1fK3xM1vVL3UpFuwdn9xVOM1{+}mNt4Gg1jT3jOF3Z5A5NeV63{+}ybY/bJr1R9scncFmbPPsOmK1wmJhTd5rUwxmGqVlaMrHx5{+}z78MYPiF44s4NXYWmlKd8rE/MeeByeM88039oL4Mz/C/4g6nZwqZdImkM9jdp8yyRtyBn1HQivoi9{+}BENtqEkpiaxljPyxxjaMjr09au6d4d13Sba5M2ofb9ibY4byISDcxGMA9B2r0Xjru6OOOWw5bPc{+}VtG{+}FmpeKNc8O2VpctqsmpxRhGSOULB820ozOoHy45Iyo9a{+}8fi18BPhto3hqx/srUzpfiqytUijvtOlOJZFQKN6jIOeeeDzyavaRreqmyfTmjsEit0jMMlvZqjAk/OFx0BwDgY613{+}jaZea2TbWcIED55EYOM47446frXNWxspbI66OAp0le55f8AAXU9W1LU18P{+}JTFevGFEWsWYLRsuBjfkAg9unb2r2jxd4Mh02VogQwRRhwOCCM10{+}i{+}GLjTY1Q7YxxnYAPxJrP8AElv5aspcsRwSSTxXj1He8rWuepFpyST0Kvw{+}R4jEnUg4r2PyC2m9BkjvXk3gtBHcKvbPFexI/wDoW3GBjuKwpJNNkYvRxseL/EK0Z4WUsFG7qKr{+}DLC1GgXEcsAeRiQqg/fOelbnjVFeTGOlV/BjiCXY4G0t37VNO17HRO/KpHimqfBDxNdfElvFM0VhfSwbXsYbpi0dqMHICA4yT368CoPi1oV78W/DsPh3xn9mtoYJs2l5p9m/mK{+}DnbkkbTwDn8uK{+}rrvwpBd2oMJKSEfwniuE1v4UanJIJLO8STBzsfIAz2r1VOpBKzOGM6NR{+}{+}kfC{+}r/s5eDl0TTdLbxVcxQ21zKzQ/Y0MpmaRY8FxGGCnCEBieORjNa1l8Tb3wR4KTwL4R0WaHTBuZ7kplp2P3mY45J4H4V9H6/wDCnWt1wr6RFcM0iyblA5KkEHP4CsCb4QeIb1P3enxWqhyQTx15P86uWKk9JBHD046xSPkKb4C3firUFm1OGG3u5cN{+}4QIGVu{+}ABgjNerfDX9nvVPDN4kVxAjxxH5JlQDcOx{+}vFfT3gr4JxaUI7zUXWe6RcbQuQv0zXZX9jHaxbI41AXoQKwrY2o48t9Ap4WmqnP1Pn7xV4aaz0aZGbLBc7VH9a{+}QPiJaoINQiQEFZAy{+}5zX3j8RLcSafNxksp5xjtXwd8WZGsptQyTjrg96vL5OUzLMIcsDxS8/dXc0p5xweO9RbCohA{+}7nH4Hmn3z7tw5wTuPqeMUkZZ7Xy85aMh{+}B2P/AOqvrNj461mSW7O2pRiRmIZghwexGK6Xw/4yutM/fROokZQkiOMrKO6sO4OK567gaO4icDAYA/Qgkf0FRthJpF4TLMP54qJRU1qaqTi9D27Q9LsvENk2qeHZR9pQZutKmGZIfof44z0z1BAz6nFg1WXSLxJ7csqk8xk/MPUE{+}ua4rQtZntLuK7tHkguYujqcEjH{+}f0rp5PFtl4pUxaggstTQbWuoVxFMo4yyDo3TJXj25rilTab6o7oVE12Z7t4d8TaZ450BtP1tUuoFwFyPnx3Iz355B4P1wayNI{+}F7{+}FfFEj6dOb/Rr{+}IwLKjlXH8SxuOueCozk/NzkYNeTWV1c6PPHuPyOMrLG{+}6OQeoI4Nd/4b{+}JUun3cYmV5Ivly6H515GCR0ZfryPWuVc1N{+}7sbu01rufZH/BPTxulveeOPAlxP{+}8guRqNvGScgN8kuPYOoP419qo24Z4GRz7GvzS{+}COp2ngX41aT4qV/s4nzb3qqcK8cnBY59CQ34c{+}tfoxousW2tW3nWkyTxN8ysjZH{+}c5rupTUkeXXpuEr9GazCo5ORT92QT2xmmtjFbnMV26VExxUz8VGQCaAItvY{+}tNPpUjcGmnkUAM74zSYxTzxSMuaCugx{+}lNx0qQjAHFMIxQIjKk0U8UUAXOQad060Y3fhS47nrQA09KXGcUqrilNA7iEUDrSkZFAGTQICM0pHHTNGDnpQM4OKBAFxT1GaaoJ5JxUqDB{+}tAD0Tj0rxH9rnxoPBHwl1TUvMjX7Oo2RuM{+}bMzqI0x9ckj0r22SUQxM56KM1{+}dX/BR7x9c3es{+}HvDCTlbeCJ9QuVUnDOxKqD2OArY9mFZzdomtJXkfDmtXOBI7vvd28yRs53HsPfkk1zEcbXTkA5knbbnP8I5Zv5flV7WLtrq4MSDdjksD/n6fhVcH7NCFH/Hw3y4AxtX/E5pwXKjplq9C1HAl9qsFlbuVtVwC/8AdUfebP5muy0JISZJpEIVF7D7oIwB9dvH1NcroNs0FrJIB89yfLHHOwHnntkgCurs3/s6yDFQ7oPNck/ekb7i{+}/qfxrCq9LI3prW7LdqH1q8mtpBtgilDTkdAcEsPwAQfgaxvGmpedc{+}UiCN5JAEwOkYAwfyH6muk0{+}EaN4Sdjhbi6c5kfqAcFj{+}QP51xkLNqutC4fO2NS3TPc/0Fc9JXk5dEdFRtR5erKes/6K2xcnCCMgfQk/qSPwrB0{+}I/2jcnH3QVA9MHj{+}VdDLFLeMWADsTuIA5HI/mafBpzWd5cWdzbSQSXSCWF3UjcccgV6CkkrHn8jk7nrXwq1MS29uW4{+}RRwTX094Pl32cXqBzxXxr8LNQNtLHHwArFSv4//AF6{+}vPh5dq9vGzHtng/0r53HxtJ2PpcBLmikz3DwvP5UCq344ru7TSkvVQMoYNxg1514cuUOD1B9q9K0G72bOjY7GvGhLXU9mrD3bopar8OLS/P7yBWHXLCsNvgjpEzgtHIpLZwGOK9at51lXgZz6rTwcnlO3QivQSR5bqS2PPdK{+}E{+}iae6EWXm46eaxau2ttMhs4QkUSKFHG0cYq8WAPIA47VUu7ohDjA/Gm5qKM/em9TP1LbFCVx8x/KvOdfuPMkKjr7Guy1e5PkOK8/1IFmZs4brXnVKrbPUoUrK5q{+}EWLagpGQcjr6V66rZtiD6da8j8ILtvUbtmvVN5EJOcZH4VdHRMyxavJHnXi4hpSpOeT3rO0WRVZVONx9a1fFg/f5xuHNY{+}lMu/PcGsW7SOuKvBI9H0TUf3SwyAHb0Nb8YEwX5u9cPpkp3r0Jrr7CQbBxwe1dtOr0Z5Veiou6Jzp4lZhwT05702XR1UcoOnOec1dUgjj9R0oduDjgY6dq69DjvI5y/gjhjbC4PqK4zXMKGIOSOK7vVDtVgSB7CuB8RH5HPTArzKzPXwybPJfH16rWsynBwD{+}VfCPx4Gye7kLbg5HQdDmvtPx5diKO4JOMA18h{+}PNGHibxFFZld4dtxz0wP8iu/LXyzuzLM4c0LI8JudIm/s03QHyIwGB15FZ8an7Sq5IDDyiPfHFeqah4McM0Cq/wBmI3gDpuAPHv1JrzO5jEN7OD821xyOMEHivq4VFO9j5CvRcLM07u2MiWbEkZGTn9f1FZE{+}8yuc4bllz{+}orsJ7VJtLPy4IOQe21uf55rCvLQRgFVCqATuJ/A/y/WnGRnOJBbOYg7HKrImcemTj9MilnMtu5ZsxyZOG/ut/gahlBhgXGFcAqQepCn/ACr0UkV6j78kqu18enQH8OlN9yYvobXhzX5ba3YDY8IbL20g3Kjdx7exrpLObTtUEnlTGwlI4jkBKKcY69QPrke4rzdJJNOu84O5PlkA/iT1{+}orVFyWmDqSFIDLs4Izzx/ntiuadO{+}qOqE9LM{+}qLG21WXwxoWsXkqW32J7eKSUEMJImAUg9iBtGPr2r7Y8I{+}D9Y0DQ7XWPC2peVNPCjz2n3oQ{+}AWBTkgZzyv5V8IfBTxpDrnwt8X{+}H7x0kuba3S7tw6hmkwy5IB46HrX6L/ALUU8S/DLSb0skhyV86PgqTgnp23FhWMI80rPcVV8sdNjU8P/F5Y1W08SWk2kXI{+}X7Uyb7WQ9Mq4PA9j0rvbLUor23WWGaKaMj78TAg1hat4ZTUfMWUZcjByudw9GHRvyzXAHwpq3hOeaTQ7y5snI3/AGdF8y2bPcow{+}X3/AJ11Lmj5nE1GWx7AJlkfA59{+}tK3B4/GvJbT423ehv5PifRZFVeTfaYnmKQOpMWS/HtkV3/hrxpofjO0NzomqW2oxg4YQSAsh9GXqp9iKpSTIcWtzYYd/0NNxxShge/FB9qokacAdKbj3okcKORRkNgj0zQMD0ph46085prA0DQ00UHAooAvKDgnPNKeB70oGBjtRj05oEGDRg96cBQTQA0gnpSgYoHt0p2CKBCEY75oAJpccetOGFwKADbgDJp{+}KTGe9A{+}c47Dqf6UAZevXht7E7CBI5Cx56fU{+}w6n2Ffjd{+}1x47/wCEr{+}L3ie8hkb7Olz9mhDHHyRgRrwenCZPua/WD4x{+}Jj4Z8G{+}I9Y{+}UJpulyTLvOAZHBVOfoCce4r8OfF{+}pvq2p3kzsSXm3c8nnP/wBYVlL3pJHVS92LZT01YvmkkywHLE/y/H/Glt4jcy5kJXgzSPn7q/54/GqZcrGLdW{+}VW3N/tNV{+}5L28ENtGMXE{+}Cxzx6gfQcmtHozRbaG3pTGZycHyeU{+}XsMc4{+}gIX65rdjt/t1/HZgBlQiWQHpu6Bf5D8/eq2mw2{+}m2ALHCxjHI{+}8{+}OB/NquW7/wBj209xI2ZD8zf7319QD{+}Z9q86b5nZHbTXLFXGeMNT84tZRHGwqi4PYE5P4k5{+}gFHhOxH791Ut8oB9yTt/lmuYtZpr{+}9uro/OpAX8CP/r/pXqXh7RxYaHdXIxkSRjPXACEkfma0a9nFRRMf3knI5zwZopv9dsbEqCZJMkgds9P517X8U/hTbatolk9sTBqVvtkRQOmMHP0rjvgtpQvvH2nNIAcqPlX3wP5mvqbxR4Ul1ad5bWGWf7NanzDEhYKgHJbA4H1rzsRWcaia6Hq4WknS95bnwro9tJo2tXUTIVKzbwPqP6V9MfC/XBJaw5wDjpmvGPH{+}jtp{+}vi52lDKuGwONw/8A1103w41v7PJFG7gZJzRXaq0uYeGXsazgz668OXu2NCG7jvXqGgXQPl7mOK8J8G6qJ1jweAAf0r2PQLoEJz2r5qV4s{+}nS5o2PTrCfai47jGRWoi5yd2SfauW0y6wnBzj0610VpP5ihhx7V2053R4tWnytkk6jZ2Hqc1mXQBQgt9e1bL7ZUPI4PT1rLv7lIYn3EdM1cloRTvdHJ645gjYnhQOoNef6lfqznZjB44rX8aa{+}J7y3tI2ILtg47ism6topbiKNFztTLZ9a4XHU9iOkTrPBimQBgv516QsZe1wy544xXC{+}CIGTZwCenXrXqi6cTp3mZB4xgV2UYOSdjzsVPlkrnlniOPG7vweorjbXUPJu2jxjHWvRPEtsY5XBOfSvOL63jtdTRzht/H0rmnGzsd9N3jdHd{+}H5DMqlW3CuzsIyq4A6DnPGa8j0HXXstVFtkBW{+}ZR9K9S07VhOihm564FaU0upyYiLexvwNuPJ7enaklOMAH/wCvUazqIwUOB61XurgGIkHg98V3XsrHnKN2ZGpzHaSSRnvXCa9LiJskdCM5rsNUk3KSeQK4HxRMEgkweQCa82o9T3MPCyPC/iLcBVm9Rn6Gvl3WrmR/FTpASknKjafpX0J8TNWWIzHdtAJJJNfLtxqlqni61ur3zDaR3kbSLHnLJu{+}YDBBOQCOD37V6mAi7NnBmU17q8z2HXdAtdI{+}Ff9p3AjiESLKD6YBDfmCR9a{+}NRci6ubqdvuSOTjPPXIr0X4xfGS48QQXfhrToprHSYr2YmGUncE8xikfJOAowDyfrXmNhGJBICTg4UH0Pb{+}VfRYSi6cXKfU{+}Wx2JjWnGENl{+}Z6Ho5e90iJic5j2kAf3Sf8TVWezkHmYX5VOB754/{+}J/Or3gVy6ywHkQtv6c4xk/oWP4V0d14fNvNIuB/Eqg8g91H5hRVSfLJoyiueNzzK4sT9qEIyWZWwW6k7Tn9Mfkaz9NvHtpzOOUYlv97nDD9f1ruNR00W2qW0/wDyzDA9MY55H5Y/P2rjbK1DXN9bY5hYyL64HDAfhg/8Broi{+}ZHNKPLJFy{+}tj5aTxMWKD5Wb7zKPX3HSorWdEkVM/u2GVwfufT{+}f/wCqrUDhZjC5zG3Ksf4WzjP45AP4elQXFqLWRonQhJOVbvGfb{+}R9vpUrTQ0sdD4c1m40XVY7i0YJNDkMoOAyH7yn2IJ4/DtX6p/sG{+}K7PW/hVPaW8ikW9wAYgeY89m9/8K/JCBpHjBHEq/Jwev8Aj/8AXr6d/Y7{+}PM3wm8YwXl4z/wDCPai6W2pfN8sQJ{+}WQj2PX2z6CsWuWakU7zi4n62uD85A{+}ZAOvfrTJEinKEorqQfvDkd6i0{+}9jukjmikWWGZA8cikEOpGQQe/HenyjDoo67yD9Of8AGtzg2Me/8I6fqbFniKknIHBU/gQetcZq/wADtHuLxryyggs7zqLm2VoJwe2JFNeodDTSO/Sk0nuUpNHlmlab8QfChZUu7fxJagfLFfERyr7CRRzx3KmtWL4mNZgrrPh3VdKcZywi{+}0Rn3DJkkfgK7pRktnnnv9BSlAynPIPrQlbYfNfdHO2HjfRtSjVobh8MOFlt5IyfwZRWhZv5oLAEqSdoPOBnNXFtYlfIXB9uKkKimLQZjIpMc89qkwMUxhnpQIZtU9s0U4LRQUXMehpV{+}lL19qVQe9BAmCTQVFKcge9KVNACAcUEZ4pQPwpdvORQAijHUU7PrS4prkqOBmgA6nHJ96VuAqD{+}I4pyLtHr7{+}tRXUqwQXEzEDy0JyexxQB8i/8ABQH4hJoHwalsEbZLrt44IHXyocqB{+}JX{+}dfk0JjOtzIRlgQ2R{+}X8zn8a{+}3/8Agpd42km8T{+}G/DaHEdpp/mMAOrSZ/{+}Jz/AMDNfEFjEwJIGBgjd1yTUwV7yOpaJIktbdYoHnlIdA21Qf4mra0Czkvbx710JYnZH/U1nQWMuoXPlxoViViAFHP0H17mu5gtmsLOGHKpIy5B6bEH3m/PP61hVnbRHRTg27vZDoo96EKAViJVc/xMTkt{+}g/KuT8Ta4bu8isYSwjjyXx/E3{+}c/ma6XxDqK6HoysjbLqRMxR45TPQn3PX24rz3RVa51ZNxJJDEnqelKjDRzZVaeqgjvvDuisdOglKnY7gZ/Aj{+}q17BDo/2Lw/LFMjAO0koA4B2xjg/jXK{+}E7QzjSbYgBZXLALxuAZVH{+}favYfiFp66dokbbQNiTEeWvB/dZI/OuGrUbkkdtOHLBs8Ik8b3Xw21jS9UsYRcPvSPYxI3YzwPfgV9PL8dfEnhAT2sk{+}naHd61DaM9rfQx3iNaTRybp1kUlVCk7WB7n2r5R8ZaWb{+}x0e/d4ba2iu2dnnb5QFVWAx1P3vbrWj45i0vxtef8ACVWovNGiUpaRLoyMbdMhnAX7xyx3Ejd60p04T5ebS99SoVakU47rTQ7n4jatDrGmaVeoTHJcyMiBgFZyqA5IHADKyMCOCDx0rlNDvmsrqPL5IY89KZrXw91/S/Amj6he2OoTSLcW9xb6hf3Cp59oxcKsUbHccY5AyQCDgAjM01mUhhnSM7COx5/lWUYxjT5U7nQ6kpz5mrM{+}h/hnrqzrGpb/AD1xX0B4f1JPKQBsn1FfG3w91prWRMMwGc9f0r6T8H{+}IUeFCXyfr/KvCxFOzuj6XD1FOOp7bpl6AoG4qQAM10tnqgAx37{+}9eX6brIZVKnPrk10dnqu7aCeMdaxg7Dq01LU7aXVQkJC8Z/SuH8V{+}KVt0MaEszHA2nkml1LXBFbuQwAAz171z3h/SP{+}Eh1Rr26/wBTGcIhPBPcmtZSb0MIQjDVnDeKNZn0rXtOubsbIpiVBY8A9cf59K15/ElvvWeOUNlccVH{+}0h4IvtY8DySaLzqFqRLDgZOR2/EcV8O{+}IfiR8R7GP7P/AGVcWbR5DSJExJ7cZziuilQ9uvddhOryLmaP0P8ACHj2zSQR{+}epYH16/4V7Jb{+}LFudNUo48rGQQcV{+}QHhX46eL9HvV/tJJLlQwJ3x7HH4jGfxFfSHh/9pC7bR0kWf5duSJOCv1FbuhUw7tujJuniEm9LH1v4s8TQW1u7zyoo6jJ5NeTXfiyG{+}mEu9VTdkfSvlT4ofH/xR4iza6PuMmcB1TcR9B0/OuO8Nx/F/V5UFutxhjgvcEMPyNH1W65pySNo1FH3YK59t{+}FtYTxJ4z8uBxJ9mQyOf0r0htcGl3KxyybC33Q3pXkn7O3w71Pwfoc13rE/2nVrwhppW9B0UegFeoeI/D9tr{+}nNFOWjfqkiH5kbsRXBJpOyNZb6nYab4gMqKTJlenBzWg{+}oZQ4PHtXguheJ77w3q0mk6lICyf6uXBxIvY16Xaa6LmFSGBGOMU1J2IlSW6NjUrlWUgYznpXmvjPVFSKQcH8a6XV9UMUR55xzivGfiJ4iS1gndnAyCBz1rBpydjqptQVzw/4veI1ikmAfr0zXibTXOmeEde1qyFy19CiQxGBAyL5rFGMmeQNuSCOQ23p1roPG2rvrerzBSSgOOe9czqvibS/Di2umavC8mmX{+}BcBGYFQpxvGOSVDsR74r6bC03CFkrs{+}Vx1ZTm23ZHB/FC1Ex0XUNxa4uLRVnTYMIwA6uPvMWLk556daytI00yWtw6j5kdSD9Bk/zp2r{+}Jor3V91hbtFp/lxo9tLIXDsqAM3PTJBPtmuqj0QadpBMXMFwqXUUxP3kY7efcEEH6V7UW4QUZHzziqlSU4kXhi9TTNRikbAjkIQqeMkHGPxXIr2s6JBrOieeGybdwrY{+}8eMBvxwh/wCBV4beogWNQCPLYSk56qMH{+}te7eBb7{+}2dHkttw2Pb7HZQDuZAGXA9wAPqK5sRpaR2YfrE47xdoU0Vs8wiCRoUljZOcrjJH6kfhXkd6h03WZroANtYSYHRsjJH5E19LazHG{+}lwyFVEaLjGMbju6H6Ekf/WrwPxtpv2WKWNeAJdqhu646n8KdGd3YmvBJXXQqT2kcqL5LB18sHPXPyn9cc/WqyA30CxsQJMDaxPfHGfqP1zWlogNxoli4GZIt8fHcjBGfw4/CqcoWxu5G2sIgADGOu3g8e4Nap62M{+}lyrbYVWwQj4OQfUV3fwV8Qxaf4xitb{+}L7Tpd6ypPaE4EgY4K57HkkHscVxWrwqsquOVmAww7nsR/nvVe1kLR{+}dG5WWD5yR7Y4FaP3lYz2Z{+}u37Mnj5vDV9J8L9cuZXezhS70K7ujte5s2AO3P95MlSP9kjqDX0lvKNlv4m{+}8OnTivhXwWs/wAW/wBnvw94t8PFpfGXhlvtFk0bHcZVwZISe4fJwPfGea{+}wPhX8QbT4o{+}BNI8RWQVPtaAyxA/6twCGU{+}hDAj8KiL6GFRdTsTg80mPl7AUgjBGVyp9u/4UPvVSSoYD0rQyGJyT9T/h/Slb07UifKqgg5x6U5gce9AhvFNJFPIyvpSYwKAGDFNIwfapOnbNI3PtQBEMtRTiDn0ooGXjnj{+}VKB3pCOBTsYFAhCO5pRwT6UvajFACYJ6ilIxS4pSM0AJk4pHBK5AzjmnAYpPrQAqtxnqKxvFshi8PSIrhJLiSOBSTjl5FH8ia1m/dqT/A3H0rj/AIkaotk2iRFQ6m5e5cEZASKNmz{+}eKT0RUdz8lP23vEq{+}I/jv4klREnit5o7KJGDBkEaBSDj3/ka8GsbKa7fcpEUIwOOMD2H{+}etelfEqb/hMvGms6lcS{+}Q8moXRlYHOcynjP8RwMZ56Vmwiz0uAOY2lwQiA8mRh2Vfb3Nc0qvJHljuejGlzavYdpukw6Np63MygZOYomHzP7n2qNL9LWO51C5OeRiPs{+}DwP8AdHp3wPTmGW/lZmu72XIPJOenoormNW1N9Va4ABWGMqiL1Gc//WrKlScneRpUqKCsjP1C/uNf1mWWVy2ScA1N4Ti2{+}J7fK5T5sg/7uaTRrDzdUjVh97HGeevaul8HaUx8WWW1Mks3UdD5ecV2VJKMWl2OSEXOSbPZfCGjsdT8MwvwXt1fHTr82P1r1b4oIkWkxpEyLG0NwAM55IYMfwwRXG6Za{+}V4o0RIMFbSzWTHdsNtx{+}uK6f4s6lE0NiYWwTHcSMpGCcu4/mTXz7leaPeStE{+}UvjHG0dhZRM5yjmQKDxgjb/7IK5Dwh8S/Efgaa0fSdQaFbWWSeKKRA6K7p5bNgjqV49u2K7P4z82lg{+}D8yKvI93J/mK8kr6KhFSpJSR83iJSjVbizb1bxrrmu3kNxfancXDwqY4Q0hCxIWLFVA4UbiTgdya9/8LTLqmjQncW82MMD35GRXzKRmvbfg7ra3GkrBISZIWCA57DkfpiscXSXs7xWx1YCq3Uak9z03RIDDcqpUqvRW9a9f8M6o9qEVW{+}UY4Y15lpsYnkJH3l6e9d9olqWVfmHGCD6818zV1PrqMuVaHrWka3IUU7wxI7966iz1t1UAna1cBoSOUAVQR0NdxaaXIbHeAMerda81xs9D0FUbRX1bX57u5hs4fnmmfaiqfzJ9gOtd9oO3SrSKLeWxwW9T3NeAaz8QLL4eXl/rOu/6LaL8ltI4O3aOuD6lqbof7VHh7VYI7mGUTws{+}DIrcD6{+}n41v7GclzRWhKam{+}W{+}p9MmdLxSHQFD1BGaybrwJpWqKTPZRAn1Uc15HaftI6fc4W2EZOepYE1sW37QB6BIHBGRk1LpSR2RwlWS91Frxz{+}zDoeu6c89vaxxSHoyKOvSvCp/2S9UXU2SOaTyC33QK{+}jLD9pBYo/IvLC2miRuhJDfpXT2vxr8LX9vM4011uEUOVWUYIP8QzyfyrsgpJaMwngcTHVwv9x5J4A/ZetNGiWdovNmQZJYf416lpXgm002IIsKIRx06GsHU/2hJDHIlhZ21qhwpcDcw69ycZrjNT{+}NF9Er5vURzk/KAa56kHJ9ztpYCvZ3SR7U1ktvEFU9{+}lUL6d4FAwCOtfN2pftJXNoxT7WzBcbioBI5xXnHjH9sOfRonaeR3kb/V26gGR/wAOw96qGEqz{+}FHLXo/V05VJaH0h8So7TWLAkP5F/bEvFKvUex9Qax/APi{+}4urbyJWcTIdpGODXn/wCzjF4w{+}Ol5NrniO0Gl{+}H0AMELSN5k{+}e56cV7lF4PtdK1B1t4dqoASVrKrD2L5Huc1OXNtsZ2t6tI0JJLAYySO9eAfE3WJbjzVXcAM8D0Ne8{+}L0{+}z2bKAMkdTXz74vxNISADuPHanRV5XFVlaJ5FcWgiLueWY8k9q8W{+}LOpfbPEUcIOVt4gMehPJ/pXt/iy6i06KU7gqIpLegr5n1a/fVdSuruQ/NK5bA/SvrsHF35j4rMJ6KCHafF5soUdx2r0hdWmufh9puXUxWk9xZqe{+}0{+}XIP5HH4159pv7kA/xsCMe2Oa9Zj8MafpnwCh1036PqF74ha2SwCkOkccGWkz3BMijFdFd/C/P/M5MNpzehx0N0JngjIw0hWJiO2QR{+}vH5V678GdXli1EQgqJowUCMMKXTlSf{+}A5/75rwqS6Iu42Und5qEH6E13vgXXDpHiVZ2IULJuJB5BDEfqCfzrOvHmg7G1Cdpq57zrtvFKmoWSjyotguYyTwVY/dH4ZP5{+}1eTfEDTl1XRIZ449s0JeMgDJyOR{+}gPNe3au8bwR3DbNyZQ7f4lPzofb5WQfjXBavpzvql3amHb9otmntwBkEqd2M/Qkfga82jO2vY9OrBNW7nnHg/RgvhKWT{+}JLkOwx0BBH{+}H51z2sWUi3M5I{+}ZGKFc9epB/wDQq9n/ALIgh8PXqQLtWKFeV{+}7nKYHPfANeYeLYQwtblTnzYip7/MjMp/kPzrqhU5pM5p0{+}WCRyaK15ok1ucma2PmRnvt7j8P6iqej30lnd{+}fAQGUCQqQCDgjgg9fmwfoTWmAbFjOjqMYDdwVOB{+}RHFZdswsdQnSMlQ4ePaeoVlPfv1rtj2OKelj7w/4JmeM/tFx4l8HztII3i{+}0wnPCkEcgHuPlOfavqnwdFD8Kvi1c6VHGLTR9dkaQQou2OO{+}xlgo7CRMMPeOTuTX52fsM{+}IpfDH7RfheSWTy470NbPtGN4PyAEfUA/QZr9OfjNojz{+}GrvUrcE3dmhuYAACTLD{+}9h/wDaie4kI71LRk9XY9Wj2hiR0PNLJluOPWsPwRr8fifwrpmqRsHF3bpKNvuAf55rdUFtzE9eAPaqWpg9HYYAP/10Hkn0pzLwKD0piGU054NOI96D0oAaBSdDTwOKaeKAGgUU6igC0OTjpTvu8CjoM0o65oAD05oFBOKPxAoAKCcUCgjOKACgjiilAzQAjsojDE4UEZJ6da{+}Lv2v/ANqTT/A1xqfh3R7xLrWLi3azcRkubMNguDj{+}MnjrxXeftj/tGT/Cjw/FoPh2ZY/E{+}pqds{+}NwtIuhkPv6e9fFvw{+}{+}A3iH4vXtzf6Vb3N2QWe81u6OZZSxOMMx2rxzkZPJ5OawqS6I6qUPtSPnt0eaY3M7FAxLBcY4Pv2Ge/Ws7UtZSKQlmBwMhQMED0A7D9TU3xLN/wCHr82TIIHEkiMxPO5ZXjbv/ejbk8{+}9cYXYylQdwJVS2c7j3/Dj9aIUb6yNZ1fsou3mpzalKWY4RsBUHQY5OKv2MStp9wc4bO8598kfyrO020a4u7ZSCFY4yOxBrZ09Af7QAGYxGjD2yGI/QitnZaGe{+}rJvDVkJte8tiVfyiVI7kf8A1hXoVlpf9n{+}JbGdBtImCsBwCDGR3/D865HQLZI9VtGYlHkdEU/UP{+}nT869RaKRorWdV2PbzkNnkkbcp/7KK82vNqVj0KEfdv2O9sfLF/oBKlZJAys/QlfMBAz9cn8qd8T5kXULZV5MlpMct6mQt/n61kW{+}obdU8PTMv3XZSp/hOEbArQ{+}J1uq{+}IY1k/1flSohHQ4VSM/X/CvLXxo9OXwnzp8U5FurS2iyG8uJGVlHrk/1xXk8i7Dg4yK{+}gr270ObQLrTdfgdLdzE1tfW{+}PNs5MFWbB4dGwAVJHTgr1rzTVvhhqSTqdOkttXsif3dxbzKvGe6sQR{+}o96{+}joVYpKL0Pm8RSlKTlHU4U11nw81Ke01OaKIOyMvmEKCQu3kk47YzVc{+}AtSgVpL8wWMScu8s6Ej8FJ96be61DpNrJp{+}jzOYpMedcn5TJ6ge39K2nJVFyR1MacXSl7Selj6S8IayLkx7369QOvFet6G6lEKFWB6HH5V8p/DvxY08cTFtrA7W56GvpPwbq6TW8a8bcZJr5rE0nCTR9dhaqnFM9a0SVkC5HDc5PFev8Ah62F3o5CgMoAxXhek3YVkwSNvr{+}levfD/WAwMLNyw4yOtePJNM9WMkT{+}M/hHofxM8ByaBq1okscmSp7qckg1{+}dvxb/Ze8Q/DLxFdN4feV7eMlhGj4bGe3rX6ix3BtmVCcjpn2rzv4r6JHeQyXYjBm8vCsPzx/n0rrw2LnQ0WxrSoUa8{+}WsvR9Ufn18IvDl74x8aNpFy0tvceUzRoWKMXUjIP4E/lX0D4S/Zo8VXviK7jeWZLONV8oq5B5Hf1q/oX9gzeKYNQuLIJq1hkw3SriRWPB{+}Ydcg9DxxXvHhT4p6ho{+}pT3FxHHqllNGhihi2xtEQuCQ2Duz74r0ZVY1Hrpc9PFYLH0E5YZ8y/E8u8MfsmeJL/Xr621K{+}uBagq0Lq3JBHOSevNdRrX7Hur2H2f{+}zr{+}4kEjKkimUDCdzwP517bo3xs02PUJmvLOeGIopiEX71z/eyABjt0zWpd/H3QDPbR29rqk0jTKjH7C6qqnq2T1Apv2aWsjw3iM3jJe6zwyX9j7Vbe1ldNXuVcglenHHTOPr2rh9T/ZW1{+}DwlJeXF8V1FY3ZTC7AKQTjI98CvsHUvjDotvakqZ7liCBHFFgn/vrAryvXPinqGq6bcWz6WllDIhUS/aCzcn02jHHfJ5qZezSupG1CWb4iSTTXrofBHxi{+}G2o{+}HNG0qC0jnn1rUbkRrFESzN8pJ47due1aHwd/ZXH9rWup{+}Jyt5eSuD9kzuWPn{+}Inqa9/8Q65Z2UhjiHnsqbTI/wAzD/gXXNafw3imvbtJGXO05HBHHasp4ucafLDQ{+}gngYUIe1xT5pLbsj3HwtpVpomlw2VnCsMEahQq8U3U4FRmccMeckVb05hHAqg4OMZNZniK7RUc7h8oxgda8WWruzwVL3uZnmXjq9JQpu3HvtrwXxfItu7ux2so/DNes{+}NtTRpHJIxz3r5z{+}KnimGwtLieRyAgLdf0ruw1NylZHFiKnLFtnlfjjxrZWWt2sFxH59s8o{+}0w92i6N{+}Pp9K8p8SaF/YmsXNuhL227fBKejxnlSPwIqnq2qSavqk11L1kbgDsOwrc0XxHI9mNOvbdL22Qjy/M{+}/GAegPp1/Pgivr405UUnHXufGTqRxDalp2ZHp9kTBayqpdmUgADJ3ZIx/L866/xRdvHoun6KGLQ2Cktzkea{+}C5H0AA99oNQ6fci2s/NtLaOFoywCld238yfT9KrXsRn0uaYksTlmOerdP61LvKSb2RpGKhFqPU5VTm9hZugdWIPpkV0WiXO6bdg72IXJ6ZHP8AICubgIN6jAYAYEk9Bium0yGSKytnZSGM/nYPoTj/ANlrWpa2pnBan0z4Slt9X0G0SVmb7RYqDzzujYofz2R1S1WwmjS0uVmPn2TRkODkt8i9vTgjPrJWN8Ob57W{+}0xTkqGZUB{+}7g85/8eH510PiC5EV6rQMxjI{+}zSh1wQF27WH4j9DXz{+}sJNHvq04pkp0tD4P1F2Znk2GdXUZDgvwP5/hXiursJvDgdlCtHdsNuOQrAEH8xXvqMYvB1w4ciOCF5HZD/AhH6nJ/lXg9xaGOLVrdt2I1V8b{+}mGAIPr3rbDyM8QrNI4zUQsVq/mE7SoTA643f4H9KwtTHkX8UuPvIPpkcf0rpb9CpRdqncv8YHQgVg6uit5Q6hd6r7detetTPKqxseofA5vsnxM0fUYwv8Ao93FNGCCUXFwgIP/AAEk/iK/ajWrWO70KXcqNiISNuHAwCf5D9a/Gj4CWBubq5m3IZVjkdQ2SVxtGQP94L{+}dftFdRBNMnhdS2{+}AsQeSfl2nP{+}e9F7tnPUVuVnD/AEGy{+}H9ppxOWtXkiTv8glcL{+}QH516cfugD8q4T4VWD2nhnTGlIaV43ZyDn5jLuP8A6HXeY5NOOxhLdkbDgDpSdsVIwpp4FUSMIwPekp/50088dKBjc80mKcR1FJ0H{+}NAhMGilHFFA7FsdKOlFLxnFAgABpQMUtFABSAYpaKADtzVLWrma10y4ktoxLchCIkboW7Z9vX2zV2oriMSKAyB1zyD6YI/rQB{+}YnxvnvoPixb6t4lnn1O4a5InLxgRIUYYSJSMEAYAHU4c9q{+}6fh15dtpL2{+}nol3YtEhi8pFQNGRuXoMEhGX5h1GOvbl/2hf2brD4m6PeS2MwtLyYfMxUEFgDtJ4yMdj1HPUEg{+}O/ALxlr/AMNtfPg/xjNFZ6lYruWSZsLc2QPEqPnJ8skkrzgE9hxzpckvU6m{+}eGnQ{+}Of2yPBMnh/x34gkbyxFa67OqqEKssVwq3MWe2Nzz9OhDeor54t8FiOd4dsDH0xX6Y/t9fC{+}DU7R/Eaxqi6lZ/Z5XAyr3MO54GJ5wTHJMqnuQgr8zoiyXfkyodySMpzwcHiumDvp2M5LRPudJaQNbgRAYaXYynHOCxHX3P8AKtjR7IE6mMjelqpYf8AYH{+}orKS4UXlkwcBdhjYPwMgKwH{+}fSug0NYm1fVowA8htXXae/HB/nXPO6Z1R1QuiOhsoJGcebDPF254J/{+}vXsE0ZuLJkHJZIpj3J2fKB/46eK8a0F2OhXkqruMNyh68gAt{+}nNew6H5k2noUO5ZLQtwMk5Abr6fOa4sStTtw0ty9fLG1nbuozPHdRunOCyspBPH0Wtb4mJJc6j4enCuIrpV7feJQg5/wC{+}VrHuSo062m2b4/LidnPfY{+}Dj/gJNdT4wVbjQ/D8zB91tsjHHdXCH88fzrzZe7JM9Fe9Fnz74mtDe6deQAAOjFQCcdG4/nXlU13cWbkWlzNAoz80TlR29Pxr23VbPfPrUIGJBM4AxyCQSP6flXkF7YF7xvJH{+}tO8jGQF9P519DQ1jqfP4hWldHN3tzPcsDNLJK3XMjk1Wq9qEIWUjBA6ZA4NUipzxg/Q13JaHmSu3dl/Q9Xk0a{+}SdDxn5l9RX0j8N/GUc6RYfdGwyGzn8PwxXy{+}UbP3T{+}VdL4J8Wv4dv0WRz9lY5yedh9RXFiqHtY3W56GDxPspKL2PvbQtRWVI8E5HXJr0Xw3q32S8hkVhtU4yfSvm/4feMftEcQ8wP6YPBGP1617Not{+}GAHGDz06V8lUp2dj7SnNSV0fQg1IXNpHIrZzz171U1qAalYOpO4sPfnviuN8Ka6GtzC7llAOOef8811FtcExHac5FcdrHSmfOHjTQZdA1a4baY9xPlTqOD/ALJrI0LxhNp160UxdYuP3TE5NfRvibwjb6/aMrxgs4IO8ZGa8C8X/DK{+}0uRvLSSSDOVYfeX2B7/jXbSqq3LI9vC5jOj7k9UereGfF2mXl1b/AGuWWKH7zGEBiOnFdTZa/oFtI8l9Bd3shKGNIpVhTH8SsSpP0Ir5j086lpFwCzudmOZBj/61d9ZfELTprX7PfWKrJwfOinY4OecA9umfxrf3U9LHryxeHqK/M9ezsep{+}JfFGjzLDcaTaSWEQHzQy3Al5x/CcA{+}vBrzrxN43nMUsMMeWAI2dc{+}9M8X/E228RtBbafpUNja2/yoYYx5rjaB8zAY7Z{+}p61QstIutUUM8PkI/wB4kct{+}NQ3FSuyY5lRpUuVLX1v{+}Jg6bb3GtyRgpvc8AA4B{+}vtX0B4A8NjS7BC4O8gZbpmsbwR4JitVR1gIAOenJ59cV6fDAttBs29O1cdSpznzuJxE8Q7y2FM3lJnJyfeuJ8Xa2ILdhuxkZwa3dU1NYY3O4AAV5D4514bpPm4HXmsUrs4WcF478RGFJedo67vSvjv4x{+}LW1q8a0ifMUbZkI6Mwr1v4x{+}OjbQy28EhMzjaOenvXztfWzzwyuQWYDJz3r6fAUuX32fOZhUbi4ROfPIFa{+}kp5lyvBO7jisoIOR6VsaQjLNFtHXow4welfQStY{+}XhdM9GSFYNFluFVlEiA{+}vzjhv1wfoaztUAg0LpsM4646Drx{+}NVrK/ku5hafNCu4Boye/Hf0xzXSeNNOgEFvbRnfthDAqMduSPyUVwSfLJJnqxXNFtHmdlEXuEOOCwB9etdrLAtppo3EPswo6AksAf03E1l2GkmO7gwuCr/d7DkDJ/Oug8RBFKWy5CGRkIJ5zkAn/AMd/KlUlzSSQQjaLZ6H4euDajTZQfnYF4l68hFxn67K29Qu/7R8Sala5aOSMkhOckldyNj2JYVz{+}iEw6L4cu85Kz{+}Xu56qRn8Oa0b9XtfFdlesdqXcCE4PBwDz/SvJqK0mevT2SOi8Q6r9h8CXexyq7TEyN2LMPz5rxy9uXTWpVb55LmHkdMEruP/j2f{+}{+}a9A8Zv5XhmaFX3BzHK3qMuT{+}OeK4oaVILu2vGG5kZIm3H1LAn{+}f51phl7rJxL99I5bWQpktiQDhCCR2{+}Y8fkcfhXLXsnzhOQcNyf8AersNXjWS3YYEUqhz7dTxmuIvVMcojIIkXGR9ea9Kg7nl11Y{+}o/2afCc974c1y7WIb3gt4Y5D94hrlVOPxAFfrHqNwyl0DK6qoZySQdozx7c4/X0r84f2adKvrTwj4Ijht/LTW9UELk/MXhR0nwOccbTz/tiv0avo47Tw7eXkzh5hEzyFT0IXt/n8qIu/Mc9VW5bkPga3ktdB0yJgocxySNg9S0mev6fhXV4x161kaDCUsLZcf6uBFHGCDjmtcfNg4xWyOViYx{+}NJgU9qYOgpiGsOfakHP8qccAgfpSAUANoPSlbnFNIoAAM0UpHviigos8Z6U4UmOaXvQSHenACkPBpRQAnHekxg0/pSY5oATHrS4FLQRigCKa3VsYAHI4xxXk3x0{+}Bdp8U/D2LWU6br9oTcaffIC3kTev8Aun7rDGCCcjpj189RQyBuv51MkmrMpNxd0fCVjrOoeOvhp4l{+}FPjWJbLxTpik2fmKNhdBui2A/wAJ28H0bpxX50fEXRY9M1Z7oRGFmkaOSMAqUcDkEHv1/I1{+}y37QnwUXxtax6/pLtZa/YrxcQj5nQcj6kHt3yR3r8zf2gvBEl3q2plIEXVLg77i3jU/LOucvGD1DgEjknn1BxyqTpVFc7larB2PBdQPmSQyw5Ec6LJg8bX5GR/Kuw8IyLc{+}NIJpMBbi238jCsAQW/wDQTXnolxp8ayHm3lMTKp6q2T09iD{+}JFdj4dudt/pjkqjRyPA2D2YHp7Hd{+}hrqqRujKnLU0tHKpZ{+}ILWNhvhjRsHrgMoz{+}WRXqfw31BZrXTQ/3EXYy56gO6Eflg149MTp/j26ib5FvYZImB98gfrXSeANfW1iMTYXZcEOWPRCuTz{+}Z/CuStG8TroytKx65Dm2Z7ST51tJ3tsD{+}JGAKfnlv1rr9QtVu/CFwnyvNbEyLjncGUNx/30fyrgIL9by5kJKySvbpIx3D/AFkTFS3/AHy5JrvtEv1dEjZg6TxbQzDksoKj81YflXj14tWaPWpS3ieI{+}MsQ6/PLwFujG69wWYEDj65ryi{+}TbflTyHbZgg4UZPX26/lXtPxK0SS1jgYknyHG4jJIG/K/h98V5u2lI2pyzSDbuR5PXkZ4/Hj869jDT9y55OJi{+}fQ8y1bc95J1IDEAY9KzmXsRXQXVpIruxU7UGckZz3/UmsaS3dSoYYPfPFepF6HkSWpVIwaSpZUCE4O4{+}3Sou9aGDO8{+}GvjqXQL2O3uHJt8/I2fun/CvrHwV4xiv7SE7wHwAcHmvhaA4lHavSvAHjyfw/corsWhOBkjO3/61ePjMMp{+}9E{+}hwGKaShI{+}5tI1l7e4WRXXB7A16d4b1xbpVJYH1Ar5j8L{+}M4dTiXZIDuUc5zgV6j4T8RhJ1yxGe/bNfM1KbR9RCS6H0RpqJdoqn5j2qS88IQ6ggLRhwepIx1rD8HaysyruIPIHPevTdPVXQYKtn0rnUbmspcqueV3fwdsL1yWgBz09DVe3{+}BFkrCRYTnHQr1r3K2sUJyE59QKspbpHhcZHUdq6FT8zleIa2PEovhFZWjllhUkegHFadp4BgiH{+}ryewHcV6zPaoWJCjB64FQGyQk4x14qJU/MuNds4q10VLNVEY574rP1qVbVG6KwHeuuv0SFWOcGvN/F2pqiPk4GetYNWdjeMnLVnF{+}JtbEKMpYnP4fhivnb4o{+}No9Ptp3Zzxk8HqecCuy{+}JvjVNNinczBVQZJJ6e2K{+}RfGniu58T6k5LN5IY7Iwf1Pua78NQc3foc1eqoaGDrF7ceItTeaXJZm/IelXovDrNbEsp7ZNb/hDwk94wkYct7V3V74X{+}yae5KYOK9SddQajHoecqDneTPl/VbFtP1GaBgRsbjPcdq1NMjJtm8tvnSMyDHop5/Tmu1{+}IvhTzdPOpRJzCF347qTg/rXFaGVke3R2GNzqfQ5HSvZhVVWnzI{+}eqUXSqOLPQvh1p0Os65Z3RTzJFJVo8ZDsflH{+}P4V6FqHhNtRv55RGZY1PlRAHAJHC4/IfkfSuC{+}Dm2311Le4DiJGeaTbwQgHP4nPH0r27QtdhuiwkjEczs8qqOiqq7UUfgCPwrhqO03c9GhFOmjy{+}z8IS210XdQVEpVmPQN6D865XUrY3{+}pHazODkBvYnJP8AL8698{+}I9tZ2Wm2tjZJuwgBZjy7k/O7Hvzn8/auS0fwXLDELiWDiMFjkAr5oGFU{+}y4DH1O0dqxjUv7xrKH2UUnBsvCNkrYCxyTSqM9FYquR6nv{+}FT67ctKLGdSrCCNPL9MbmH8wKPE8UMOn2diG3YRMAEZ5wR{+}ZJzWdqlzHPBYxIGQSxmI{+}gIOQf/AB4VhLWzOiGjsaHiq4Mum2oi2yyvNFEiNzuJZgOPQGqGr6FNppuxh1jikR/l5B27MnjtuyM1r6M9tqPiHRVuIT5cMvmsZGzkpuKD8wPWug8R6fJceHZb9Q7NeSvk56J8xH6oT7ZFVR0gFfWVzxHxFatb2Rcjd5obHfA3YwB25zXn92rXGriHl3Zwij3PAFexeMNMW2s4GyMqNuTxwPm3fm2K4L4caGfEHxU0CxeNSJtSiMit0K7wSPpgEV6NBrlZ5mITTR{+}jXg3wkvhLwv8AD1rYyxXemWVzrsnlAkfZkZYgAuTjehJ/AelfYeqyDUNFsoA5ZLwRcKeqnG79M1wXw{+}8NWc{+}u{+}IPMhVrOx0u00GKNM9AjPLj2zKv4g56V0ngGHfptjYmQu{+}jb7OXf13RZiH4lcN9CPWqgrfM4qjv8juLSDbAhI2sQCTj2qbBAFPC7Vorc5hhGaQrxSn71FADMfnSYxS9yDRt4zmgBp9aaVBPSn4zRjnPpQAw49D{+}dFSYz2FFAE2eaUDIpeKMYoAWjHeilyPwoASjFLxRkAUAJSgY680lOoAQEE0tFFACMoZSCAQeCD3r5a/ar/Z7g8T{+}G7/V9MtlNzaRtceUF{+}/jkgEcgkZ9ecHFfU1VdVtEvtPuYHGRJEy8jPUVnOCmrMuE3B3R{+}CHxA8Py2GqXEpt3NveI6K7R7HV1OQj44LZGM9fwxnmNNlcWyyo3K7CV9Sv8Aj0{+}tfUn7QHh2Ly7q8SGOHZcGK68sNsjkB{+}WYcdDtwR2wfx{+}X7q1fSdXvLWRVQN86Ec4HXj6dfw96mjPmjbsddWNpcy6mx4znafUdJ1bzMrcRB{+}P4SMZX65FWNJuFttYlAOY7iMSD25yf54qtJC134XnjzuazkErJ3weGK/jg/Q1lafeGG7hUtkpznpkHhv6UNXjbsCdpc3c9b0bVxGbadsD5lVyOOSDEw{+}hJzXqPhLVUk06B2/5YzLOFHVUZiHP/AAFSfy9q{+}f7bUPJS8jYDAIYfiM/zGfxr0Hwd4mWyWOeQlom3RyRg5JG3PH5P/wB9V51WHNE9KlOzPTviT4cFz9oeJnBlUCRD0HQhvfnI/EeteD6vp4toNQXBDqFQA/5/Cvoy11p9b0bT2klSVdpt5JT0kZT/ACbCsP8AeFeVfE7R309HZFyZnKhs8Buw/HGPyrDDz5XyM0rx5o8x4XcxE2M0z4ALAsvr1Jx7dK5p1DASEcbjyOK73xRbrFLYaYnySCPzJFb1Izj9B{+}Vc9qdnFa28cQG1sknAzxn6/SvfhLQ8Oce5gzacGdQpPIJJx0AFZjDknINdleQ2tppMjbj5shCrheduM/5{+}lckQjO{+}FPfHP/wBatou5z1I2ZFH98VsWLbX5rJGwDPzZz7Vp2uCAc/N7Upq5pR0eh3XhPxVc{+}H51CszW/wDEgPI{+}le7eEfH6XpieOQHpgk1802MhYAZI46iug0jU7jSblZoGGCfmTs1ePXoKWqPfoVnHQ{+}//AAB4vSZYjv5OM4Ne{+}eGfEMbIoMnP1r86vAHxU8i5VWYoxPzRuefw9RX0v4O{+}JsU8EcizdcZUnpXz9WjKDue9SqKcbM{+}urHU0cJkgMB09aui9jIz{+}ArxPQviDC0a5mGcZznP511Vv4thuItwdcH3rONRpEyoXZ6Abpc8Aj2qvcXaouM4xXGjxXAFz5oB9CaxtW8f2ltEd1wCR79KHO440bG34h1kQq5yPYd68D{+}KHjWHTbaZnkAIG7k8Cn{+}PfjBZ2VvLNLcpEi5G5mxn{+}pNfIHxR{+}Kk/i64khti0dluIyTzJ9fQe1VRpSqu5dWpGlHzML4leOpfE9{+}6o7NbqxIwfvn1P9Kx/CPhObWr2M7SRnJOKraRo8ur3qxxoSTjmvpX4X/Dg21vE/lZfjnFepWqRoQtE8{+}lCVaXNIr{+}FvA4hgiVYiGxzxV7xPoDRWLgLjA5BFe0af4U{+}z2{+}9UwQO39K5nxfoq/ZGG0nIIrxfatyPXcEo2R8/HwsNX8O3luVGXjZcEfl/OvnvxDoqaLPZ3dumYJMMy9lfoR{+}eK{+}ydC01IrC93rhUJ/LFeK6NZw3/irUrBLOO{+}kT7TsgkA2iNjndkkAEZ45HOAOtexhK7jJ9jxsbQU4rucd4NnS2sry8nBClVjJHVh1wD6ksv613PgyM{+}ItXLNmPeQo2dEwdxPsMDr6cd68w1zWTFcNCCqjzWeTACqrE84A4xyAMdhXWeFfEI0223szRsUDN1yoPReO5/oK7asW05dzz6Ukmodj3eeysyqRLMtxdw8JKyggkcKx{+}nXHcke9O1aGOC2azgCPBAg3gt944z{+}POT{+}Ga8us/G76dGLl18{+}VhsjiA5Y8ADFWL3xiNP8AtVzeOZdsRlmdD8u4rwq9uuB{+}NciUnY6pOK1M3VT/AGj5k7ghEXzPlGM4l3Lj2wlV/E8KW1nahCMiXbnPzdAfw6VEt016IQwCzF1Z4gMbQQ{+}VHtzj8aPHtxttbMRklwVL8c5IAP8AI1q1siIvqV2vpFvrDynVQCpBXg8sxP8AP9BXsV75DeDNHUSItxPbH5DkkFl4A/E/pXg2r6nFZRac7MAAoUtjBwCefyIr13wZq1r4m1bRrKJg5lJd3P8AAuTkY6fdTP8A{+}utYK1PmZE5XnY5z4waIbdobSND9oeMBkA5Ygbf5lj{+}VYv7MelRQftAaNLdQqyQvLMExuBKxNt/Ns/0r0nx/MNX1HU9TiDTL5Lm33gDGOEbjoCScfQVh/s3w2ug/GfQ766cPBBfWtsxxgszDj8M5rajK1M5cQuaVz9VfhvoD6D4SZJwGv7m4lnmPJ3SMxJx7U7w3YLoPjTU7YbimqxfbSzEkNMpCyH2yDFx7V02nWYgijyAscCBVI6ZwAf5frWN4hiNnHb6yMq2nXPnSAd4GG2QH2Cnd9UFdtrJWPI5uZvzOmXOMenFKwwKXOGPOQeeKVulWZkZBApKfQelAEZIHakp3eigBnagn2NPoI4oAjPPrRTgCaKAJ8UZ5xS0CgA70vFHPrml6mgAxmm45px9qQDmgAApR0oo6UAKBmg8UL9Kd25oAavNJJwjH2pwGD7UMMqw9RigD4A/aB{+}HH2rXfEVlbtFE0im6RmGCFLnjk85ZnHTovvivzz8YaXc2slxFNC8N5p820qB0jJJXHt1/MCv1g/ac0K40u9i8QwIrk2nlMHB2owkVovxLlh/wMV8LftNeEo7DxLc6jZoC0qs7ryBLDjPfuCP0rhX7urbuekv3lP0PCPCV3HdLJG74LKI5M9wTt/kf0rHvrd9MvZA{+}SYZPmHqp4JFOH/EtvEu4CXgk4bI4Knp/n2re1q0XVNNhu0H73JXfnOR2B/Hj9K7L2lfozBaq3YmSFnnt3TEomj2deG6EHNaHh2{+}kt5Z9Ob/WA{+}ZASe6np/M/gKzPClwt3o7WrKRdWr7kXvgHkfhmrmrRtaTJewDLI{+}8Y9PT{+}Vcb0fIzsTdlNHtPw11iLUtMv9Ec4uY1FzalmwGAGMfoqn/ZU102rWZ8T6CY2QG6RQHHRm9D6g/Lj2Irw/TNcOmahZ6pafJJA4lwvAKHhhjvwa9qS7ia7hvopB9kmUP8h6Z5H4f/XrzqkWpXR3wkpRsfO/iSwli8TXEs3LygjaRgZyent/IjFYOsrGLSzZyyyPu4xzweeOte4/FvwQ1wP7UtmLx45jCgnJH88/5614lJbTIxe6VlZU2hGAwu48f/rr16NRTSZ5dWHJKxl6lZ7bQYLNgKTkdsVzsNszPcsOiISSfriut1Zh9lKxqRlF6/rWDYoyaVqUrHG4rED6/Nk/yrqjLQ4qi95GK4JKgdav2J5Ge1WNA0dtZ1G1tEwr3EgjEjHhMnr{+}FOFr9nuWVTuUMVBIxnB4NaSfQKSe5q2sDKoIAINakKkIFP6UmkwieMDrxVya0aLDc4HrXnyld2PVhF2uJAxDL2PUY6iuz8OePdU0Yqok89ByAx5H41xMYKj5iMdsVZhfyWBJODxj1rnlFS3OqM3F3R7xoXxue2wZfORuhK4YV18Hx/jMajzpRnsBj{+}tfNVvdgA4HStK3vCq57iuKWHg9bHXDEz2bPf7348SuhET3Lj04Fclrnxl1W6R1gXyif43Ysce1ecm6LJhW/E1DLJhcnvURoxXQ0eIm9mO1fV73WJ2lu7p5mPdyT{+}VZdtYm5nVdvB71aYksq9eOmK6zwX4Zk1W9QOmQWBIrZtU0Zq82dt8IPh6bydZvJ{+}XPcV9YeFvCkVjbR4QAgcmua{+}GPhJLOyjVEz8vU8e9ez6ZpqpbggEnHp/KvCrVOdnqQioI5{+}609YogFGDnJ215/4tsfMbYFBAz098V7BqlsqQj5QMdzXn2t2jXEzbVPHQVybM7ItSR5CukMLO/jHBbOABntxXyOmp3vhfxBqtwr/vZo5Y3D8gjdnlc4OCBjPQgHqBj7wk0pldgVwTkjivkH48fDxtAnuNSjuEMdwzg2/R0Oclh6jBOfT3r1cFUXPyvqeZjabcOZdDwKS9ee9d5fn2sWwe7VvWWoPFC7M{+}f42A7n/ADiuWc7bkq3XOD9fXP61dvJ5fssSQjBbCt6j/P9K{+}slG9kfJxnZtm9a6tPcXQkJy{+}SoyflQeg9z39q19d18zNZafDKjKnzTuSfmcduPwH5VxyXn2S1hVfv4OCOvXk1LoarLJOrNkycl{+}{+}7HQe5/wrL2a3L9o37p6T4auGubyWTI3ME5YZCDIJ/ln8Kv{+}MSsuhpKCwWMxr6tjJBJPr0/Ouc0WdbexlfO2SdSijnPHv8AQGuk8RsLvwnPtYM0Z7nJPfHtiuCS9656cH7tjgvF2qxxyRRSDMaohCKcDPIP5gGu{+}{+}GF5cLbxcqHkjYKka/MF4BI59f5V5J4lCG4RzJ/CrEHk857/j{+}ldv8AC/WUgFxJIuXEIh467Sx4{+}gAP510zVqByRlzV9T2rWdYtU057WSARwySxxM8ZDEogcgEZ4z1P1{+}lRfChI9Rj17WDASdI1Oy1t0SPjy1EgK9{+}AdhyeBn1ryq01i51fXls4ywJmAUI3I6gn3{+}6v4Z9a9t/Z30dNcg{+}LWgQyws13oE6R5bl3jCN8vXoFck47cUU4WikxVpps/V6zmF1ZQtGwZJF3Djse9OlgS4NxDIgeF0Csp6EHII/KuW{+}DOs/8JN8KfCeqO6SyXWmW7O6dC4QK{+}P8AgQNdcExLKcfwqOfqa7VqjxtmYvha6kl0hIZ233VlI9nKx6sUJUMf95drf8CrarGtrc2PiTUAABBeRpcD/rogCN{+}a7PyNbZHFCB7kZ4PHSkPQ0pPNIeRTEIBRilOR0o3ZXrQA2ilpKAG9CaKU80UAWMDFAHNKOlFABQaKDQAUUuKMUAGM0KMmgZpTmgBcYoo70oHNACHg04AUmMnmnUAeNftD{+}G/7Z8E3CpGGaNSqgjjIKMhPsGUGvi39ofw4mv8AhXRdQH79HjaBpIkA2DYrKTjoMuw5OOmK/RrxZo41jSLi1OcSo6k9cZBwfwP8q{+}P/AIhaJGngYoYYrV3SWCSROiyQNswV903Y9SorhxEXujvw0raM/MHZFbTXenXTKkIbajHqp/z/ADNJbajN4fvJrK9/49JhskX{+}6SMB1P4fpWx8T9HFhq9ywAI80rlTnJ5z{+}v8AMVlLCms6RGJQZbi2jKEZwTGD/wCy8V1wanFMmScZNIiilk0DV1kyHGcB1PyyDp19x39fpXaOEubdDwyPgI69DnlSfTIyPYgCvNYJ2ERtrjdJGPusOSvuP8K6jwlq6lBp1ywZXB8l88HPO38{+}R71nVg7XKozXws1bJvsTyR8si8bc4ypOAf1wfr7V3fg3V5I7F7HeTNCDJbAn/WJ3j{+}vp{+}PrXIyr5dyXlVmdQA4/vqRgj8f6U2OU2kuFkcSRYZWBxuHUMMd8fyrilHmR2xlyux7d4d8TWuq2ktpIFlfy/MQ5wSpA/MdAfTg9mrx7x1oHk3LXduW{+}zmXayMDuixxg8dvpWzJqfnG11WybyZgcyiNej55YAeo6jvW1qYXxFYPd2aItyYws9srfLICOq{+}nsfw6isabdKVzWcVUjY8WYxz2KTGZG{+}dk2jjAznjPSs2OKOaxhtIWO6admI4PHTt{+}fToa6m98PvEZRHC7xFid4HzA{+}hHHI/WuevY7nS5nmg/wBH2qY/LU/PtZSp469OM160JKWx5dSLSuzoLC4vPhxPYa1piwXXLQtOo3pE/B2gAgg7Wzk4zzjOOOZ{+}wOJhkHk967D4M6JJrB8QzXdstzZSwrbO0{+}f3JJMgkyCDwIiDgMcPkKcYplzZiSYkIq4bO0c8dqmU{+}Vl0oXjcpaBBibZwB3IFdc2iLLEDtyf/AK1YFrD9luY3xxnBGK9K0OJb2LGBjqD7151Wet0enQhdWPOrrSZrUklQMdsZBrLcMG2kYJ7CvYdQ8OmWIkoNuCK4nVPDrQuzKvTnkfyqYVk9zaVJo563Vu44rRt1{+}bcenoKatjIh27cgdCRxWla6Wznpz1yOKqUkRGLuQ/aN2GGRjikUvKduCC1bUWgs5UAE98AVu6X4MlndB5fB74rJzSRuqcmzG0XQJb{+}4iCqRnnOK{+}hvhj4GW28p5Ihnrms7wJ4C{+}zOrNHjAH3u9e{+}eE/DyWyINm0D0FeXXr30R6VKly6s6jw1pQt4o1CgLwea7q3ASBeOnQmsbTIdqbcbSOnFbDMAgUcjJ6cV59zeWpl6qPNDLnOTXLXVnukY4BJPpXUXqliAefXPeqUtnuUkjGOc4rN7mkdDjrqz2SM2D069cV85/tJaB9p0F7mOEubYvNIsfJKAc4454P5Zr6ruLQFwpHBzXkPxp8OR3mnTJiXHkStmBXLnCE/wc{+}vBIB6ZGa2oT5KkWwrLmptH5o6pZSWdy{+}QGVWyG7Yz1q2I0uVjmiYHcNpTOPwz25rS8QWJtJPLaTzWCBtvPykjJB47HINc5b35tHLRRYBPKg5H/wBavuYSc4po{+}InFQlZ7F26jDTKGcCNAQjOuT75//XWpaRQiPGxkEa5Z93ByPx9c1mLr9tuDTQMrZ646Vt6S9pdgG3cDncy52k/h3{+}maJtpbExim7pmtpE5urxNhKxoPusPmJ9sZ5J6/WtLUtYeHRktDHhpJG8xienHP86i0y2VnEcexSCGYgbSR37de2BUvikwXkBkij8uJHxk8Hpz{+}tcTs5WO6Dajqed{+}IQVmYkEs0Q25OSNp/{+}vW54RuzpmiXszYIclRkcggdM/U1ma{+}VubC3mTlhlW{+}h4/mBTYX{+}zeHY03FVdznHQZ//AFV125qaj5nHflquXkdb8Ktajh8RXOp3JAFvGSinnnBwP5V7z{+}yn4li0r4l6REqKJL{+}0uYpWY7VkDDbtbPXIGO3X618p6VNJZxXlv0bYxfI9FbB/PFelfDnxL/wj{+}r{+}H9WQ7WsLkJI2eAj4IP4EEVclbYhSurM/W79knVhJ8MbzQmY{+}f4c1i90xlY5IQSmSPP/AJAPwr2fAZ5mB54X8h/wDXr41/Z/8AGM3g/wCOHi2yLifTvFGn2{+}uQbT8u9f3b47ZIZfyr6x0rxFb3VmcOWfnIPXP0qo6o5JwfMyTXJDb3ulTgjH2nyWz/AHXRgP8Ax4LWtuyK5zxPexyWET527Lm2YE8/8tlrbju0ZRk4OKfUhrQmPNFJvX1pcg0yQpMD0paKAEwKQjBp3ekOO9ADfxNFLRQBMOAKWk20uKADBNLntSfhzS4IoAUUtAooAKKKOpoABT6aOpp1ABTCxPTp60{+}kxQAghDLsJ45zXzp8QvBk80Xi3TbdHa4huhqdtuIwDIADz6ZWVfbcPWvo9DzXnnxEs20vxZpOsrkWtzDJpt2du5QGIaJj/uyKAP8AfrKoro1pyaZ{+}Q/7Svg9tA8W30K8QTAXEQVcBlKgjGe4Ax{+}FeH2F3Lpd4ksBYSxt5iEjr2IPsRxivtz9tvwVZaJ4oEEEE8pWT5HySuHLHacddpLAd8Yr4q1SNZbjYi{+}VJjpnpx3/IVnQ0XKdU/eakTara295uu7MCHeSUj67G7rx{+}ntj8ci9nWxkhkjXdb3Ch2jB5jfOGx6cg1csL4wbWZdySDbIh6HB4I961L/Q7fWNOMttJGL5HANschnTH3hxg4wB17jr26k{+}jMpRfxRNnQtWTX7aMSSh7pVChuhcDgEj16fX8qluyXBAIV05Qjt7fQ15qstxot3G8ZMbqMFT35PBFdpBrsPiKNQWFtfYyrfwy/X0Of1/OueVJxd1sb06ymrS3NPS9Z{+}wXbnloScSQnnYT/EB6GuttbzymW6tH{+}XqQp4XPf3U{+}v9RXll/eOsg81RFcJlQ{+}OGx2Iq1pHiaWwnRt{+}wZywJ/M/T1rGdLmV0dEKvK7M9XldL5ZJ4UzME/exHuPw{+}8P9ociuW1HSI7k/J/owTJ8oMWYnBwc85Htn86u2Wqq6LPbzLGx{+}b5DwP8A61bca2utYkYpDeEfeH3X/wA{+}2a4lJ02djgqhj6HffZnhspbOKeOO1RYlC7VneN95E2BnBDSgkHcQV5443dc0sahq0tygtQJJGKm0i8qLGeNq/wAK4IwPpUkLnT3RZosNghZOMNkY4I6{+}n9a2by7W4KPbRvcSeWu9pWAO7ADYxwBxx3wBUSrNoqFGzOGn0hlY5yWz{+}tdV4TUtdRoRtKqSc8DNOl0u4SD7RPHsLDIXPGO1aWgad9nX7S/zhiBjPU{+}1c86l0dNOHLI7Ky00SxbNm4{+}4xg1mat4ZSdTlACBx9a6HTf8AXRI4BYjGAeV{+}o7V1o0hbi3Bdd2eTjt71x89md1kzw6fwgyycJx64q9YeEhwNvI6{+}1es3HhoblULkHrkdantNBVc/uyQeBgYqvbEqkjjNG8IRyEELyOORXc6L4SRCpEY{+}XjI9a2tM0kRhTt244JrrNNsFX5cgtxnIrCVVm8YJEmhaEkW07MkcHFd7pdmI4/u4I/L{+}VZGmWojcHcPpXRWxYBSBgdCf/rVySd2dCjZGjbOseMY554FWXm4AIGOwz1qquQoBBxnrTJJwHUdugFIixM3zy/e3fTvU5g/dbsDnoAKrQR7nzzz1ArTRCIwD1xjFVEiTsYd1CBhug6ZrwT9pLUfs3hzUIIEE80kQjMaAhhESDKVODk7QR8oyNw56ivovVIkt7SWVwFRULFjxgDk14B49nvBJfnylaxtbIz30ruquvmg4iATIO5Gwck4wM9SKdKL9onYVSa9mfDPinwFqdx/aV/DbyQW2nObZ0YHfu6tngdD9K80ubW2a5kZZGgYkgpJ0P/1q{+}6PGWhHStSubm{+}trbw3p{+}riFmkkMk1pAqx9Aq7mdnx3AbJzk8187fEX4N3GmFdatrNp7W5AlhC/dKlQcn86{+}spV0rXeh87VoOXTVHic1s0YAK7kOcEdKLdZbMq5Yknn5O1dNDp5t4LgXEL7TCzRn0YAj07YpllBE13CmxSqgHnv0P{+}Feh7TSxwOm0zqPC9q9/BtEpSSXBkJJ4HHUf56movHlrd6Sn2doWht2XcOPzz{+}X616jo3wd1a90a38Q6EjXtuFKyiM/KpH16dB16VzfxZh1vVrCxtdU0ldKKBk80QtE746A5{+}VsdMj{+}lc8YpvmNJyaXKzxbSr5Zd9vJgJKSAx7dMVsLYBmhgIOyM{+}oyT2xWfa6H5d6iEqGGSUJ44xnn{+}VbSqthOJ52ZWKhtipk4I6//AF66mktjli21qZMDROLl/KLBR1P1yF/HH61p6cJtO0{+}SCQ8TAbXByBIpDAH8f5mpQwjtGeC3eC2JMhZiQ7dhz{+}tbHgTTYfEel3ujsxjlQia2nznYwyBn1XJ59s1Ll7txpe9ZH2F8F/GtvfeC/hp47mYodEvh4e1U9hBMPLRm/wBlSUPPvX37baKix5T0GPWvys/ZU1mJfEuufDbxCxh0rxXbNYNEWKtFdqSUIPQHlwPXI9K/Rz9mjxrc{+}LPh1Hp2rXKy{+}JvDkp0fVfVpYxhZcekibXz7n0pRCZ1mv2zeRbwNyZJ42wP9lgQfzxWiFlHGST1qzPYtPcq8gBCnILdQP/14/Kri25d/T61aWpk7aFBJZonHJPtVmK9ZTlj{+}Bqc2pCg8mhrJWA4GcUyeVCrqSnggg1KLtSm4EfSq5scDpj1xUbW7IuyPIGejDiglx7FxLnOM4zVjcGXNZKLKD2wO1SpdGEjIGB1wKCLGhj2J{+}lFZgv1X7zc{+}5ooCzNkcUZoAoyPxoJFB5p1NHWnUAFLjPNNJxTloACueaXGKWigBAMUtFFABRRRQAqnBqrrukQa7pFxZzKGDr8vbDDkHPbkVZp{+}7AoGj5y/aK{+}DcPxO8B3ojjC61p9sSJNpV96glNp7fMATjqARX5IfEHw4z{+}RqEMUkFy7SLcRFQqxyhjvQcnOOK/fO806G6WcMABJGUbjqDX5S/tmfBS68AfETWrm3hH2a/kOoWxjHyybsBwFzwwYkkeh{+}lY25GmjqhLmTTPjXTYvtCzwlcSoC6KxwDjqv17j6GrcZJADExshOxs4Kn/wDXxS3U6yXEd5G/lyjKtFx{+}v19aW{+}gW9VrmEMsjclQvHv79f51ruWmWJ5LbXiI76Jo7ngecv8XHHTr9awr3RrjSWJB8yAnhgMEe47ZqZ51ugrcxv2LcDPtVyzvriWMxH96ucFB6e47003FeQnFPUyTqkksflXREsfTe3X61E0nksNrdOQrVs3tjC3zLF5LdMYK59fasdrXHCZ442sOD7UJxewveTsy1Y6u1q{+}YgyZI3ID8prstK8TuGDxzFcjDcZOfcd/rXCxRQGRcxqG/uu20H2z0rTbWI7SI20FiFlPd{+}RnHauerSUtkdNKu4bnp1n43shBtluFkyQGVCF59SCcV2mlRQ6np9rcWl/bM0zMoiV8uoHcjoAc8c54NfMMkz3N5hE813wgwOWJ9MV7z4P8Ea/wCFPFY0e8mjkis3CO8UZKqcsCu713Iy88kggdK4a2GhTjvqdtDFTnK1tDu5tEvAFjklyNvU8jrW5oHhy7vwA0wZFHy{+}Uo/yK62x8NtLZC4MexSPuv19jWp4Z0hrW5KKBktx6V4bkrWPdtYxLGwbT3VVhKrkAny{+}frnvXoWgWYljKhcZHTbgZqe70Xem4x89xjHNaPhmzZZgG9SMHiuaUuxvFXHT6OyrlgMY6YquLAr80Yzu6gGvSRoIliDMmEwOlZOoaGsSkIBx{+}JqbjS1OZtLYIVzk{+}vtW5aKgO0DAbp61WNm8bYI6c59KmswxfDHPpUN3NIo6KzXAVg3PtWrFIEQE4PrzWRakYDbeo{+}lakBViAV28Z4rOxqXo7h3Y7c/405iZGUkc9OT0qa2tvMUFRkAVYS0yACe9Ul0MXoOs4c4Ax6nNasDDoR2/KoLW1CMAAx78VpxWzMSFTBPc8mtlE5JyRy/jmR20hLSEhZ7txGCWx8v3m7HsD2/LrXzb43s7IanbQXV3cLby3E0TzTZaF5cRPG7kbss4WROCpA5Oe/uvxUvJ9TjOlaZIpmgVp7qQKGO1cnajg/KcqAc44YeleG6rYqPDV9eW8smqRy2a/wBoJcDb5flzBXjQkdkV{+}f8Aa962/htJ/MmKdRG3Nq1/rs8M9xq9pBqunXhsV0DRis/nxzArGzy4ZU42/MrZGWxzjHo{+}rfCa0l8LWenzwRSGOEIy4AUNjnFYHgLwFYaHB4fSKxSEu0sz/ZbtYIyPtKsoYY3SBc8D6V7hqcIWPIUg9ga1qzuouOiMorlk0z4O{+}LHwDt4r4S2tv5EEKbTGkXyP1ySeoPXmvHvA3w1eP4iWlhJb7kcnyy3Ibjp7V{+}gHjjQVvi7hFckEAnqK4XwV8Ho7bXkuLmBBLu3oyD7jfjXqUZSmkzkrKMGz3P4RfBfQPDfg{+}0SKwjEnlp5gZeQ4wGIHbpXmn7THwl0Txn4T1fUZIcRaejLb3EK5kZlBLNkn5gTtXPP3eK{+}ltDid9Jjj{+}ZCECyBTx06564rivj3bwxfDvUYtgQSRiBEXGSzEKgA9ctmvW5bRujw1Juep{+}RHiTwxovh/VUkk1BHWMCSWIoUkB7qFPfpyOOtcRPFc6/wCJnuNhET/NGg5AXtz65r7Q{+}P8A8NdJlWBmtIvtKJ5fmAcjHv3NeEaT4Li0{+}4RLeMkn5ck8n/OaFNJeZbg5PyMvSNKsryLyb2CQI0ZQk84Pr0qCx8Ea14W1xJLG2e5hkLLEyf8ALRO657HFe{+}{+}FfhC13b/NE3mMMgkdPWuw0TwPqGmzC2yDuIMbFenPbNQr7F2W589/ELTLvT5NH8daXA9vPaTxRX8RUq0Vwpykhx0OQFP4etfYHgj402Wl6zpXxf0B2v8AQtUhi07xTp1t8z27D5lmZM53xFmGf4kBx2rpo/hBZ{+}L/AAzqFrf2qTR3sXlTrjG/jg/UdQexFfGj{+}AvFnwP{+}KtxoelanLpc84KxXLFVgu4zkqX3/ACYI{+}U7uA2a0Whk9T9edOvLfX7K31CymjubW6RZYZoTuSRCMqQe4IrUS1Veq5OOtfnf8B/2mfG/wPZ/Dvirwbfy{+}HjM0kT2sTMtopPPljP3M5bapOM8Y4FfZnwy{+}PWlfFKUxaNFJMQB5kjo0Sr68OAeOe1aqxzShJbHozWgbiovsu3P8Xbp0rW2YfJOSo5oEQfA4GfSquYqTMz7GccgGmtbgZzwc1rCPjGM0ksIOPl6D1707j5mYE1viUKMfnVSa0bkg/L04rZuoWwpxye{+}KZLbMdqj8aVhprqczNZEv6fWituayLyE7CfoKKLD{+}ZZU0Z4NFKOlIxG455p9FFABQOtGaU880AOoxSA5paAAjFKBmkpcfrQAmMGgHFIWC9SPxqnearBaDLMPegC6TUU93HbJl3Cj61xus/ECG1jYQqWOOMCvGPHvxL164jmjt38iPHBUfMKTZooNnuHib4jaR4ftJZbq8jjVB3YZ96{+}Hf2uPinpHxLsha2MbSXlsfNt5zxtYZ4B9xj8cVyPj7xFqs5keeaa5fJBZnOBXimu63Pbv5sknfg7uTyDWDcmzeCUTwPxFB5OpTr1DncCBgHj07Gm2F68W1Ap3JwDnqDXQeLbNLu5aeJfmDHO0c1lGzBjjLrgOuN4HQ9s10XutSk{+}pWlZJUeWIFg5IeE4BDeoqjHZSMcsSMcq3Un61pWtnPDMS8fmGQYdc/eHqPetqXTBFZj5Vki6/MOVHoT/WpcuUpR5jlJtTvI12{+}YzJnAbqagu7{+}TykRkQPyfM28t9f8au3NmzSnyxhACSx5FZZKuQDjA4q1Z9DOV1pcrTTvdYDct6gVu{+}E7gWOoGa6i{+}0xiCVIoic/vWQqjY77SQ31UVnrCLgLFAP3jdTnoKtQ6lf6M6waa5s5FbbJdR8OzHqN/UL7Dr3zRKzXKhRTj7zOv8P8Ah1YNbsJP7NvLWfTWTUbl508shAQYguf4pHKqPqCOK{+}hfhZqV1rXge{+}mvUF14i015FW5Kl33u7SqxyCu7llBbkADb3ryT4f8AgnWPGUluLXU7m2EOHl1OcFzcSjO1VDdUTJxnqSTxwK9n{+}HWmXPg/xTcwajFEbme0K3E0W3y3xIuyXafuqd{+}CRkhgAOCK8LE1dbJ6o9zD0Wld9T2bQbu213wzDdQgDKfMmDgMByOQM4ORnGPlq14Yt4/PfeOjbeO1cZDOvg3xa9scR6Xq/wA0LsQqxzdcE{+}{+}MDJ712GlXaCcOjghueO9eHNfaWzPcgvsvoejR6dHdQDb8zHjkVlvpjaZfhlXCE9K1NBdpgOcD/wCtXQT6YJELHnisHqdEXys3dCiW5sVP3uOQeAahvNKRgw2YPrWl4HjG427A5PAyetdDdabGZdu0/XFUldHPOpyTszyy{+}0FuqD2yDWfFpGwknKHdwMda9Q1DREUFVTCnv71gXmkAAsSMHr7UnFpam0Kqlsc1DAUOQMLV8IxEe32zge1WUsRFIM5Unpmp7ez3FMHDKeMDis7G/NbU39FtP9EX5OT6Vfi01m2tjg45FT6THHBaL5hyCM4NSy6iUYEYUKcdK2SUVdnBJym3YSLTm8wkZOfU9KzvGmtjwnoTzRoJLqTKRKjDcT0wBnrycfSrUerASHlQMZ5OM{+}1eW{+}KvEC{+}J7m61SQtLpVuTBZRjBExyd0v8PY7V/PnNbxnGMXL{+}v6Rj7OUp8r26leB7jTNGuri7uDDdarIbRJ7gkCIeW0kkuD06L2xnPrWYJFufhb4j167KanG8D28STKEDKWw7HZjqxY8Y4x0p/jpodP0Lw2rXSwG1dppEeEurhyu8YxgcbupHAbHIALLC1tz8LNTsXGYLmOQorNtynPPtXLzaq52xhuxfBHitNQ0zwtLOtmky/abXYLdn4jnCrjrt4GdxP9417BqN/EY8segx1r5/8HzTweGtCMsC7bYyLcGG8EarMHiBJCg{+}YS3OOmeSepr0m71mfUAVhQgZ5kLbR{+}FbtvlivL9SOROcmjbgxqt{+}beBSXI/1in7v416X4f8AA0FtptuGjLyKCS7gZJrzHwtZ3CyxgzSSHeHcq{+}R9MV9GaNYNHZQ7sqdoyK{+}lwSThZnzOPk4SSRgmSPSIvuStjkJDCzc/lXH{+}KtPn1uWO71JRaabbN5sdsSC0rDozDtg4Prx6Zz6fqlkJMgjH06VwfizSfMtpF3MVI6E16NjzIyvqfFnxptTqerzCPIhBOOM4FedeBfB32/xFF5se{+}PfgqB/WvoX4heFiZ5AFG3JJJqh8JvBL/wBsPPJGwGMKT0xWDjqdaloep/Dv4eQiBcpwR8oPFeif8KksJ9jvAC6D1PXj/CtLwbpyW0cYKlSa7rA2EL0x1roS01OOU2nocfovguHT4DGEzzwTXmvxh{+}D{+}g{+}LfJOs6bFdRQPvV2QErxgr67TnkD2NfQkEQ8o5XrzzWZrGix6lbyI/8qLIhVGndnmHgbw7pej6FZ6bas11BbJiI3LGRlX{+}6SeSOOM16D4as7O3hVUtkifOdygCsWz8Ff2dJIYs7T0UHit2zsZbadF6qo60IbldG6xBVhzycdaeybAwHTGM4otoSAu4DOM1OE3jk9TTOcjWEhyRjjkj0odWKHoC3TFSRRkZw2QTiknYq4IHCjtQBUmj3cMvTuKiW3O3J4PvVwnzCqdSeTn0omjAODgYPQCgDOeXyzjaCOo70VZaMZyV69BjoKKdx3Zj08dKbtpehpCFooowKAFBxQRk8UmKcDmgAzjrQKZI6xgliMe9YOqeKIrYMsTbm9qTdhpN7G7NdxwruchRWLfeKIYDtjG41ytzqVxqJ5Y4zzUsFhI3LZI96V2zRRS1LNzrV7eONuVXNVmtZLvPmSFiK07axKoTjnpj0qzbaeVVgeWPUU0u5aOVutGEq8gKK43xH4XWWGQLHtjbqQOTXrc{+}mmVNvQnoKztT0PfCyeuaLIZ8Z/EPwOXhmIUnJPHY18zePvCN3byvtj2queMdM1{+}lms{+}BItQtpxtOV7EV4V41{+}FaS5LQjHbC/WiyKTPz8udLmhdiYwSeoPeqLWIVGBjA5xgHpX1vqPwCa{+}unZFKc9QM1Un/ZjkVA8cTs4/X/PFMdj5e0bSnuLlFaM7dwwT0Fdd4n8MCxsoWgCB2TdIXGQnp{+}desf8ACidT0m8VzaMoBz0Nc98QvDM6XluhDQAIAynAz74rmnK2rOulFvRHzf4iTbJ5UaDys/O6LjcenbpXKyW7K{+}1fmHUMO9fTVj4Hg1I{+}UIhIG4LqvU8/nXQWnwFtLyIE2qtxwwXFc7x0YaNHSsvnPW58m2W/TbyCd1OwN82BnjvXungT4a3PiaRNS0{+}W11O3cAlomBIOOjKOQfYiu8n/AGcocFfJwuM4IzWVcfs{+}rosMt7ZS3enyxoW32kzRHge1cdfFRrfA{+}VnZRwc6PxK6O60AwaFazuitJLaMYZoMFAsg/h5Xk/TI5HNYuja/f6x4k168bUog9lp8jm1IMZTJUCNMkHdlQwXnlN5{+}5VPw9od74ii0/Sb6{+}1f7Olu05VNRkAulZyVaUKRubkjJycYGcAAdzoPg2z0u/Flb2KQE2U4hjhT52PylvbHGWY88ADlq86nZTs3dnfUUpU02tjWvSvjzwXDG7FJnjDK6lgyOOhBPPB4ye4rU{+}H2qS6spivmMWp27eXcRjHzMAPnALEkNnPQDnA6VT8L6LNpGoyaQdqRyMZYF{+}6TuGT7nG36AY9a7K18OXHh7VIdaijd4lAjvo0H34e5x3K9R{+}NcqkneHc6nCSSkuh6R4fjZNo6YHcV3NlArpjBJbvWJpGnxQwoy5fKhlYqQGBGQefUYNdFZRBXTGVXqQf881zP3XZnRo1eJoaPD9lvI8KSOufT8a6d5AJCXXOT155rDtoxJKMHPv04{+}lX7qVtoJUEj9a0T5Vc4Jx55G9awJdR7T8wAxzWDrekxRqzKBt54pbfUiFPBX2HaqWo37KjDacHOea2lVTiZ06U4yumc9cKke9gc4J57AVShmDTgnhV5Ip2pXBYbvlQYxt7VhyX0jS4HIJ9c8VwuSPWUOh2Y1RrjaFPy47U9/MWMnqPas7w6pnCk8Y6mtzViLS3VEhFxNK2yOIhtkhBGVLL9w7ckH2Na01KqzCrKNLocxrTXF5G2m2ryRvMMXEsUjKyRnho2XHVuxB6E{+}ormzbRXWpPZwRAW1oFXCBCAw5K9CRxs7jj1r0tdIh0yxmmuJC3loXlmmfkgDuT6AY57AVxOkSm30ae/uN{+}zDzuW6oOWI6Ae3TsOtVWWqitlsKjNNNv1Z5D4wF3r/im8SC8Ki3lSFrZBlsEFWJ9Fzke{+}a6/wAXaZdXHhDU7WIsZ2s5I1UDuUOKteFWGu{+}M9RhW1mtbUn7SgKGIXC4XY74ciQg{+}ZgkDA2jsa9D1XTIbTRrm5kUiKONiQBntWfI27I1VWMY3Z4Z8KvDlxp8ccEsdzDDdyu8cYCPsVli3b2/hG5SMLyckZGCK9403wvHtDbS2APcVxXglIpfEui209rPbXCWvAlZd8Rd2LIwAGSdgGRwNp4yePZJwliqZIK9zjoa9H2d1G/Y82dflcrdSvpdlaRSRKYsOWAynNey2spS2jU4wFABrynQojf61ZKgG3eCSox/WvWEiZAATx24r6LBq0D5nGyvNIilG4EnkemKwdYtRcKVAGB6jrXTiMAFTzUEtmsgOBj3Nd55y3PGde8CR6tKzMm7nOcfzqx4b8DLYKxEQQA4GFxXqTaYibtwAI/Wli035QFGR1oNfaWVilpGniKNc88VtLAFUjOBiljtDDjgYqdUy3T/69Bk3cVEGOO3amTAyZGeM1MyjacDv1pgG05PHvQSyuYjhvkOKkS33OcqO2PepmO9MZOM05FwOhJ9M9aA2EYhcdsCl25VcdO9CjfuJGRjvT2YAYVeMfnQK5EqlHXjIJ7nvTJQXHI5JzgVIzgLu/i7H2pCdqnJBwOn8qAsRFjGu/G7jAxUbIeOSQepPWpXdeByT1xTC2yItxycAj07UANkBdyc/Wil83A6qP60UCMIjNA/WlHSmkc0AOpRxTVORSjrQAvTOaq3uoR2MZZ2AxRf3qWUDSOQMe9eaa74jkvZyindk8AUm7FpXNnV/Es145jiO1PY1QstNluW3OWwTnB70mlacWG5/v9ea6S3gEajGQcYzRbW5o9NiC10lQQpxnqRWlFarlQVzk9PaiNG80BRkhfyq/ANzDjJUHFMW5XWMlsKuAanhUAqgBz16VMsPzKAcYGc1Oi7NpHBAz0oBuxUNvtZeCRg9aqTWnz/dzmtpE2ur5PTv0pHVJI1kHIzj3oFzGEulJJbS/Lhj3ri/EfheO4jZfKHK8MB9a9RijWWNicE9qyrvTxLGwIGRx0oLTPFYfBxNz8kecYJBrsdO8GQ7CPKVxwdpHArpk0EKyuBkn2ratbHylwqgKy8YNBXMcNL8P7G6f5oFA57ZxXzv{+}0Z8Coje2F/bw4QqwbLBV6569utfZkVh5WXIGMZwe9edftD{+}HTq/gpfIUho5Q{+}VAJHHvWFaN4M2w9RxqJHw7ofg0aRdi3liwpPD4yPwr13QPC9rLGjgKB228ZNc3d2cuI1QSbkUB{+}ACp9x6c1taHrn2NfKkbBBxtr4{+}rK02j7amrwTOmfwbbSkEqAD0FNn{+}GFvqVtJEyAoQQ2R2xWvpOridFYkKVroNP1IbsklhjJwKzXKDckfNPwi8AWd7r{+}oz/AGeCOdVjt0EKbMxKCEyuwAtgDcwJ9D6nv/EnguHRdUsHkjA80Oi71UpnIPIxnOM429D8xztrU{+}Hs6p40trBNKubewsobpmvZhgOzT8ADGCmQ{+}Dk8gggYrtPi5crFpel3Nu7pLHdhSI3xlSrZz6gjjH{+}1W1L{+}JdmE52hyr{+}tTzzXPCqadZWOtx7g1tJtkCqxDRtwQQvLYO0hc4yBniu20jSEubfmPORk7xnP1rq4tCttQ0OfTpvm86MoQvA{+}Ydj{+}Nc14BMsWmyWU{+}1Z9PkazcKoCnbjBAA6Yx{+}Vcji002dftFJNCaHYjRmbSpGysWZLUY5MW75l4XqpYckng{+}3PRWsBmkGQFTI9ttO1Ww8{+}CKWPC3EEnmRPnoe4/EZ{+}nB7Va065S9ihmhVYxLyUzyh7r1OcHIPuOuQa2nHmXN16nPGfLePQ1VtViiyH4POcdfwqKa5IJ5LE9AB0{+}ta8ESyQYJ5x2HBrNv7UAEAAjsQOnrxUTi0tCISUnY524uZhJlMqc/wj0qCWZ2VmYc1dmtsEsRge1VEtv3mT8ykdx1rkaZ6ScUtDLv7aS5jBj{+}UD{+}HHWsS2sne5wdwGepHeu2uIf3HTBxjpiseOECdsnH49afINVNTodCijsrCSRsEINxJ4YnsB7mptCQ3yz6pOAXuci3DAqRFnKEqTw2CBx2Ue9YUCnXLxdPRithCN1yrLw/IKgHOQSR/3zn1FdjbzR2yTSSTQA24WSeN3AZYyeXGeuAGOBXox/dx5ep5NaWrm2YPj64a38PizWSWGa/nS2RlJ3c8sRjk/KDVPWNEli8Ialb2iK87Wzxjedu7K{+}pOBn1J4rf1XVfCetzI51WdJNMjkuzDLbtGsqIpLOu5RvwAB8pIyR9a{+}X/H/jzxL4k1WJmS5SykkZbPSLa3NykoX7{+}YhxIV4yzfKDxx0Oc5LmTIoSc4OMV63PbfhlCdU26pcafJZzxWaWR{+}UiFtrHbsJGW4JJPTkYz1rpfGVvPf{+}HriztAhlmwig9M7hxXyDo3xj8b{+}Evi9p8l/cXUtvdyxQXFrJb/Z0khJCfNEOA4HII5z3xkV9YfEDxRN4b8Pi9sIjLes6iACPfltpIO3vyBWkNUrjmmpcrRgeEpjd{+}OZbwRSw7UWENLglyAWbKjhclt49nxngivSrtzKxXHHoK86{+}GeL/Wbu/C4S5J2spJDCMLGQRuIBVgw6AkdiMGvTJYspg/LjqR711q99Tkk0nobHw9tCNVkIA2Khz2Oe1emHkD2rhvh4ARPKA2CduSfSu7jAA68CvpKCtTR83iHzVGIi5OTQIjuPGB2p4cfSgyAdDW5zMjnt1IHAzSxphh7CnsyuQM4p6YHQcUAyKYjIwe9BXC596eyqxGfyqJ{+}w5I96ADJC9TyetBjyRUYLHA4FWI48EEjI/lQDDOAvueooBz370EBVGcdeQO9KCHz35oJDrwOhPNAULHwe/pQY8A9yTmjBRTx1oArSnjApYwGQ7xnPGfSpAVJ5HbrTPu4AwQRigaIZAcFlUnPTdxRLIqqMHgdvSrDgeWHI{+}6OprOnkymSRk88UAWFkDDJAP1FFQROSnJB/GigRlHpxS9KKM4oAKjmmESknjFOcnBrnPE{+}rCyt2BYBcdKBpXOc8a{+}JVkYWytjJx161kaTpcrp5r{+}oGTWZGr6pdSzv8Adz0I{+}7Xa6dZ{+}dEfL6DB9s0G3kaNjaDY79hj8KvlGcnaDsIzU1tb7zsHygpnA74q9HCHkXHKle1AhtnAEkJ4yVqzHbjnHp{+}dLGnKt0yMA9RT1yWj4{+}ooFcjVNpQsMDHB/pUyuDt/vY5FSNEHVQMHnFCQBio2leO9AmJIFZE75H51HLGwtQOuG5qxtUMv{+}yOlTXEYeMKF69{+}lAitbRmPhj1HSkmhQsFx97g4qaJVaQAnpxUiqEfJHPrQHUqG1AVcYODU0aAAexxTpE4JHHOadwR8pHrQN6itEdvB56Vg{+}M7dbzQ7uJh5hEZIjx1xXQRsGJO4cGmTxRSo5bGfpSkrpoqGjTPjq70e3W9uAiN5ZbcBu4{+}lYGpaQ8D71TaPQHmvUPE{+}nrY6xdRoqRhZG5XvWFc2UchGQfUAd6{+}LxCtUaPusPK9NHIWF1LbuV64Az7V0llqbpGeN64zUTaGI5GmAwpGcCpLC12lgo3Ejn0rjOo4XRNWu7f4htf{+}Uk1gyywBnOHgnMmSV/2TGgGD6{+}5r0jxwwv/AA60YQzEOj{+}WrmNmwQcB/wCH615xf6XqV34o/sr54bOO6FxbsmxA00qlVDMWGU{+}UkjnGOOpz6B4ntN3hmYkK{+}xFb5n2hgMHg8Y{+}taq6lFmKs{+}aJ2GiahI9lA0vySCNVYMwODjkZ7/XvXPzzrovxFWeSTZb6xCIiCOsqkAc9uCB{+}JqTwlczXWiRmTyhJz8kEnmInJGFbPIHTNVPHlpIdAkuYWZbqzImjYZyMZBIxz0JPrwMUVPiaCCukzt9244xlc5PNY9/J/YWowXv8Ay7Tfu39VY9Mnsp56fxAf3jUnhTXF1vRLS7VQrOvzgHgN0Iz3GQefatC9to9Rs57W4RZIpUKMp96cZcrTIkubQ63TblZLJZEOFZdw4qvfZlYke{+}RntXBeBNfuNK1G78Pak7STod8M7YAlRj8pyWJZm{+}cnAAXb24ruJ5Sp55Hfmt5WtpsYRTjLUoXWCjBe3JzUEUXlYf04xirM8nlfMeSRiqryEMSRgA8DvXM4nWm0h10fkwO/rXF{+}INTFhMqRfNLM6xxptzvc5ITPUZAPPQck9K62SZCrM7YHJ2lwuT6ZPeuR8NaefEGuTeIDh7dsxWRZdjGLJIdh0zzj6D3q4xSXMxX1sdl4W0VdI02KEfvJG{+}eRj/Ef8P8ACppfGGm6Zc6vFrctvNLawpJDELXbMke4FQZD94M5wB6ge9aFofKQcc5HT/PvXlP7Sss1rZ{+}G2tiITcTuZJlXL4j27R7jLk49QK1S0uedXfPa5yfjn9oyOe336wsNlYNNugikgEgRcHDgdWI5yelcVL8ZbSHULuS31y6tLi0iwjm0KRhCwJxs3ZBbJwR/EKr/AAg/Z21T40Wd/wCIvGt08Gm3hYWtnCoR/J6JluvAA4BHSuT8cfAH4neA7mfTNE0XT/E{+}k5P2e8kZ0lCk5AcKy57flQ6bb52zphThrFS1N27{+}IC{+}PfiDo/ibW4IzZ6QggtoNnlte3eWaIFfUttJA6KhPevqLXbbHh/Tp7uOaU2qfaZI4TgybIywXjGMkAH2Jr460T9nHx/Dpx8U{+}LY1a{+}06SGWytbZziCN5FjdVjHGcvGQepw2T0r6k8Qa9e31joCxSy{+}a0DPcLC{+}FcKnz7vVcZ/HFPSFRRve5DV3ZdCx8Hgwa8depzJKMEAO7E8cnsBn39K9MklVwEX5ievt9a82{+}EsMun6NcmYjzgVRwG3DcBkkHJ4yT3P612lvqsbzhG3Ek8LjIz/n{+}VddNc07M56isrnp/hBEstNATkO26urhfcOOlZWiaasdnAAuBtB5rYVVQYA/Kvp4q0Uj5OT5pNi4Ap2A3Q/Wjdg7SM{+}9CkE8d6okjDFGIp6ytt4HFNI4J9{+}lKCdoHbrigB7rgZpincB/OpjjFQg7myePYUCF2gEccd81JnGBnkDNREnd17U4DdgjpQFh23eAMdf0qPlWCgYHXHpUw4x60gByGyOn50CGoxAHY5/KnF1XqceoNGeQOOe/r71G6k9RkjgH3oBDC4AJB6etPGCASOfWoHYE7fUVIu4KOcgdvagZFdBvIYKc5NY7z5kCZ6HFaN/PvYKTjHOKwrd994Ox3d6Ckrm0iAqDkZPPNFXFcIMBOKKDMwOe3NLg0Y49DQeB1oAhmJ2kA15d8S77bIkJYqGznB5xXpN5IUwf5V4r441aC/190JB8sbQP/r0FxVy7oExaM7uEZBgGu50p1jEewYUjBrgNLkeRFVTgL/D7V22k58kE7gM5AoLa7nTWAEOJM5bOCB1Aq8mVUENj5sZHWqFqWUscfMeQTV{+}NGVT/ABdOaAslqWlKlBzznqKeEG7B4APr0pIkBRuODggmph0Y7ffFBA3cAcZ79fX2qRcnjIBBPJFMeH/Z9/TNPJwQMYFArjfLLFSeCO3tUysGGMHI7moyfmGDz0pzZKcc59KAEJAkJI4zUsmF5U5Wq6gsAT27GrAAManFAxDtlC4z05pgVVAB65waQRlFLcjP5U1X3N3I65oAcLcnA6AnOM0NFgHnIIqcBjgAg96jlzGp/iPagDxL4iaV9n12QouWkAZhnr{+}H51yJtAC3GPQH1r1v4oaYZLNLtIczA7N2ePxrzBSxOw5Hsa{+}Vx8OWpfufX4Cpz0kZrwFN{+}eVPGP8ACkt7LyZNu0nNajxgfgen9aQgTYCrtIOeh5rzLHq3PJ/ilY3OoX26zuWsJtN8q9E8aFmIBdT095FH416Ett9u8GlZXG8wFXccYIz/ACrnPiRE2nt/a6zrarFbSLkRLJlgpYAhvlx8pzkHjtW/4NMV34LiijeN41R4tyPvXAJ79{+}vWrstDJPXQqfDK4iu9DxBJBIIXKu0AwpJAb8{+}fU565Oa6m9hWWB1PzBlwVPpXFfCGIxLqtqzAlJAVQQ7GAJI5P8XQDPtjmu/ubb5SBgH6VVVe{+}7CpytE4T4cyy6RqeraO6YFuwdGAzkHv{+}Ix9SGPXNd8WVj94njpXnnjOA{+}H9Z07xBEmPKdYbgKoYlDx378kdR1HpXodo8V3axTRNvikQMjr0IIyD{+}VZJGjetzm/HenzGyXVbA/wDEx0/LoF6yRkYePn{+}8uRn1rc8I{+}KovFGiW98pDiVATgEAHGGHIHQ5H4VZbGzHBzkc{+}leVTWkvw58bG8hJXQ9RlHmKBkQyHqRyAM1pBpaMHDmWm/wDWh6/PIxbjgDsOarNuIJGRS2dzHdWqzRSK6t0KEMPfkcHnI4rJ8a66NE0lIYIhdaleMEtrckrv5wSCDkYyD6YHvSSu7MzT1VjL1i{+}PiTUV0C2bMLoJNQdCGXysgrGQRlXLKf8AgJ967jTrKOCOJY0AjUbQMcAdqwPCnh/{+}ybI{+}fIbq9uGM11cSHLSSn7xJ9PSusjGyMjk/Tn8qLtlTtH3YliLBUYGDnNeK/HrWo/EN9H4fgV01XSF{+}2q20jfBIMMV9drJHn/f{+}texK4ijd3fAUE5z0FfLHxZ8K{+}I7y0f4iaE0n9qxXb3EcW0MHgxtKEdSrJgFT71blpZGMaSqOzfp6n0h8Mte0zUfA2my6aUEKx{+}XJEnBikHDIR2wfzGDW5JMsxDMNx4wM8mviDwB8S7bXru3/AOEZ8Qr4R16eRY5tLvmGwy8A8P8ALIpPQE5Fdd4q/aP8Y6PC1qp06CWFVjnufsjIyygYkXazEDa25fwz3p6tWRk4Om2prU99{+}IPiuPS1sNJjlRbrUZlyrchERlfJxyBlV/WuYuBNo{+}jNNC9v5xU2IlmIASJo23v6qpKpyvODjvivNfgxp2ueO9VuPE2qTXGyRPL{+}1NkBxkFvL9c4AyOAM8k8D0jxUhWe4hQx2lu9tDZRyOhIeSSRm8sfKwJKwkc8VlSi5Vky7KMH5m/4QsVtvDNsI2BHJ6bcDOAMdsY6dvwrtvCWmS6hrMWV/chgTzn8K5nRd{+}m6TbQzAOQgBwck/UjrXpHwztnur55mUiJBgH1J9q9rBwUqnMcOMqONPQ9QtU2wgDqMVL9xCO9KpCoNvApjHPfk19CfKoY2QM9c0{+}InJJ/CnLgLjvSp93FAC4DEA9qbsIJIFKxwSBT8gLjGBmgTGBuR6Un3TtHHv6UrEA8cYpqY5Lcn{+}dAWHIArEEdqcxXJJPIHbvTQSW9AaSRh2wMCgRIhVgTzk9P8KDtAHp/nilQAA9uOlNK5I9KBDZe2OMf5xSs3yjJ596exxgkVWuGBQ9aBkIcGRiMFelBmCMVbIUjt6UkMQbd6noTTHDebtGDg/eNAFK/laLzSCNm3HPU8isrTRmdSzA4Oc1Jrk37uZcndlRx9ai0yEs4L8oD0oNFojpon3oMLkDgEminW6lohhtvtiigyMMHNIRyaXocdqbIcAkGgDL1yUW1hNK33VUk5r5uub5NV1CeYuRGHIPUYr3jx5dvB4fvGBwojYmvnHSYGiWaaL51c/Mr9D7ig1gdno{+}peQUDNzjZz39K9L8Mv9rhiywLbSDk{+}leDT6nNbQSYBLRkED2r134eagNSt7eRXXBfDAHpxVWLPSoINwUk7iRgVpWql02HgY/GqFgsikDPAOOelayIY41PcHGKkhtWJY14BB4x/kU7BIAzye1MUbe/U4pCwUKV/KgzJskDa2DkdaaMHpyTzQrbwex7j1pqAq{+}OueOlAIZcDZt4ozsfkHHr7Us0m{+}QbjTZPnGQe350FEjsNhAJIHPSnQs0seBwq1TZnV9vY9xVi3k2nb2I6igCdX3jaR7fhmgREHcRlW4pq/KTg08uyYUgkYzxQA6AY6njpT88/hUY2nIycdaGHlnjJGetAHOeO7RdQ0GeMnA4YkdsV43PbeUThsg9{+}v4V7rqsH2uzkjZSVIIrw6/j8i7kDIFCsQMPkn8O1ePmELpM93LZfZIeCm0c4PPekUDOAAfpTDIWBHA54wO1PB3McZzivmGz6dI4r4vRBNGtZvKS4jExWSCRsI4ZWUgn6Metafwxs4LXw/BZIqB4IwZGgkWSFmbLZjKk5XBAGTk4/AafidrddGklvIo54oWWQh1BBIYYBBrkfg1catDd61Y6vEYbiKdjE{+}4FWjZ32YA4XhTwMD2rRHPN9jR8FEad4r1iwElwVlzIqMuYlIP8J7HB5GPTk8Ad7lmwMZwMcda4K3jaz{+}Jz/u5mDwtukik{+}QA4wHQ59PvYGOmTnj0TysDB7elbTW3oQnuZOtaOuqafPbMVXep2yEfdPY/gefwrE{+}HurfaNNuNOlHl3Fi5jKsV{+}7k46dshh0Fday7VQE9Bn61wur2h8LeNLbVQzLZX37mdEVdu84AYk4I7dPSsZdzaNtUdi4DIeMc{+}maxvEWiW2u6VcWd0A8UoIJP8Pv8AyrcVPmyzZGM81g{+}Jvs5hgt7u4a3tJnIlkQqp8sKS3LegBPGTx0qoQdSaiuonVVOLm{+}hwfw3{+}J9np0OuafqGqWOqWmhtvkvI9QR0jUvtO9txC8kZXrn6113g63m8RTr4ovsP9qjDWiqVKbCBmQY4ywUcj2rT8F/CX4ZaXo3iCx0zQtCubOUeRqMbvHcqWJUKjZ4DlkBHGTuY9Rirmn{+}H7HwTLaaFpSQ2{+}kwxeXb26A/uVyWVVzyQFIHPp04r1sXg/Z01KLvbc8zC4znqSi1a{+}xsRMFZRjBPerakgVFDbJwNwOPSrCqQDk8/WvHR6LOW{+}Imom00NbSJcz6hItogORndw347c9eOlatvpMFtosFgMNHGgRsqPm9cgVh30n9uePoLfyy0Omxb3crlfNfBC88AgbTxnr2rsIoX3L0PbrQ1d3KbUUkeBfGD4LeGLTSb3xDp/h8XGvKVhtI7ZDuaZ84YYIwQEc57cHqBW94f/Zr8B6Lrl3qH9kpfagJ2aWW9czBZs/PgMf72etb/wASvGcvh3V9Jt7cB/JkF68QiZ2lXBi2kgEAbpE565YYB5x2WhxXCaNDLeqq39xmaZMdHclmH4EmhXbsTzySu2Ri1S2Ty40CBRgBRwK8x8RO114p09baQtCk0yypt484BGK4PT93JGwcHpuHcivV7v8AdRuCwfIP515/qM0kmqw2a3qu0121zJCigiKN9giJbrnbEx2/7YJHFdFJJO5jNtpHQzxStHkj7oAVxzx0AAr1n4TW7/2M05G4s3BORivMpZoki3KxA{+}6oHpXt/gu3Wz8O2iKQVKBj9TzXs4BXbZ5GYyaikdAmSvNIUIAYdenFLv6elIz5PBxXtHgAgKHHX1IpzqTj09qWNPlPORTRlcjoaABHGMHtTpX{+}Wq/O/qKeWJAGPpQFhgLMxGMmoNSivZdKvl0{+}dLe/eB1tpZF3KkpU7CR3AODirkJAJJ9akPJJA{+}ooEfPmhfGbxX41v9Nt7BItNi8UFW0aWS33tbC2jY3/AJoPpKqov{+}/S{+}FvjL4o8Xano1vHBDYReJZIo9Mka33GA2qBtTEgPU7w8ae65r6B2qzcfezwaYyBM7jnnH06cUxHgfhP40eKfEGraLaSQQxf21NBp9vtg4jubUxNqoJJ9GuY0HZrVjzmrHxA{+}IUWl{+}MtV1vR/FVjbIPC0V1YRtsmj1KZZ7kpCmT824jbtj{+}c8YIwc{+}6RMFJJ6/wAqNm6QYQYHI46UCseUXnxB1WwsPGms6pqL2ljp{+}pRaTa2tvbwgwNIttiSSSRguVeYgsxVAuSQcCuZtPivf6rofh62k8XWmla1d6peWyyzNaiKezguNpmckYZtmxF8ogM8gIG0Er9ASKBGVZAVYYIIyCKhcLIoGOBQOx86eKPjLf{+}FPBuoSWd0LTWLW61{+}7iidYRBdLb6hcRxxHzW3uTtGViG7nOV{+}UHvvH/iiN/EHhGKw8ULpdqdck0/UPJaEqZfsskiQSFwcMTswvBO8d8V6PKqM6Lt5ByM9qhmQBU2orDdkkjrg//WoA8c8D{+}NdX8X6leQXwji/smNNP1ALHtEupKzCbb3CBVjdfUTj0r1XTIlDAgcngVi3wDXhJ5{+}cc5rcs32xgKef1pGj2NNYMjnOf97FFEcZKD5tvrkUUGRiNUbd6KKAOB{+}KhK{+}F73BIyvUdeteDeGZ3ktpGY5zKVI9R70UU0aRL2t2yRWtwq5Aj6fpXT/CG9kiKwqFEYYMBj6f40UUyz3aykYtMemGHArYVice/NFFSZMeTiMnvkCnNGBx{+}NFFBIkKBj1PrTj1/GiigY2UBsEgHJxUVud5wenIoooKJ5IVwG7jiowdrpj1xRRQBalXDZHWoyxIXJz8tFFAEoG1Cw4OKCxIUnqRRRQBFccpj1HXvXhHi2If21dnJwrkBT0FFFefjf4R6uX/xDIjOCOO1TRudiEcZ54oor5GW59gtjN8ZRpL4W1JWUEGBs{+}/y5rhPh9qlzaeJ5LIyGdLnT4tSlknJZ2lYuMZ7KOcAAdTnPGCiqRjPodLq9uJvHWmXxZxNHAJFCtwMtsIx6YY16AHLQjsRgAj8DRRW89l6GEeoA4A7njk1h{+}ONOh1Lw1dQzbtq4dWB5Vs9RRRU9Gax3LOlSvLpFgztvkaFdzkAFjjqccV598bdOi1LR4rWUssbxTZMZ2kHZnOaKKcHZpi3k16niUPwT8P8AgX9m7xnYabNf7ryS0uZLt5lE26PeV5VVBHzN1B619Ix2g0XxVpelQSSNYpZrPHHKdxjyCmwHrtAUYBJx2oor0swlJ06d33PPy{+}MVUqadjurcfdxxk44{+}hq/CoPB5BOOfpRRXlI72cR8PMTXHiC5dQ1w{+}oSq0hzkhQNoP0zgV20JJRCepJooq6exFX4zzzxJKL744{+}GrGaNHt7LQby{+}QY5aUysvzHuAAMD1Ar0C6uHaPqASwGQOeRRRRH7RjDoY{+}pMWDR5IVsA4NeYG3Fr8RtTjSSQot40ADOSAqRgDjoDjAz7D3yUVcdn6G/2kej2EC3OqW0T5KF14HpnpXvViAkSqowqrwB0FFFe9gPgZ89mPxo0YjkA{+}opQN0mD0oor1TxyeMc47Diq0zHIOetFFAkJGNwJJzinkZH1oooBj84K{+}9PxndyaKKBMSNByMdASKcyjdjFFFAIiKhCQPXFCkgj6ZoooBiyOSSM8VCrHgduaKKBjFUNICeuP61Wu2OdoJAJzxRRTQjlrljJfoCeAe30rfsQCVGOMUUUi5bGrCA0YJA/CiiigzP/2Q==')");



        //    driver.Quit();


        //}

        //////Tirar foto
        ////driver.SwitchTo();
        ////    driver.FindElement(By.Id("btnCamera")).Click();
        ////    Thread.Sleep(200);
        ////    driver.FindElement(By.CssSelector("#pnlCapturePreview > div > div.section-content.buttons > div > div.imagebutton.accept.dark")).Click();
        ////    Thread.Sleep(20000);

        ////    //Bot�o Avan�ar
        ////    driver.FindElement(By.Id("avancarBiometriaFacial")).Click();

        ////    //Cenario 7 -- Formaliza��o 

        ////    //Validar P�gina Outros

        ////    wait(By.Id("avancarBiometriaFacial"));


        ////    Assert.IsTrue(IsElementPresent(By.Id("panelAssinaturaEletronica")));

        ////    driver.Quit();


        ////}

        //[Test]
        //public void CT005_Cadastro_de_Clientes_Funcionarios()
        //{



        //    //M�todo de Print
        //    ScreenShot print = new ScreenShot("VendaCartoes", "CT001_Cadastro_de_Clientes");

        //    //M�todo simulador de teclado
        //    InputSimulator simulator = new InputSimulator();

        //    //Execu��o

        //    //Acessar P�gina do CCM
        //    driver.Navigate().GoToUrl(baseURL);


        //    wait(By.XPath("//button[@type='submit']"));
        //    wait(By.CssSelector("img.png_bg"));
        //    print.PrintScreen();

        //    try
        //    {
        //        Assert.AreEqual("Login", driver.FindElement(By.CssSelector("strong")).Text);
        //    }
        //    catch (AssertionException e)
        //    {
        //        verificationErrors.Append(e.Message);
        //    }
        //    Assert.IsTrue(IsElementPresent(By.Id("login_username")));
        //    Assert.IsTrue(IsElementPresent(By.Id("login_password")));
        //    Assert.IsTrue(IsElementPresent(By.XPath("//button[@type='submit']")));



        //    //Validar Tela de Login
        //    Assert.AreEqual("PSF Security", driver.Title);
        //    Assert.IsTrue(IsElementPresent(By.CssSelector("img.png_bg")));
        //    try
        //    {
        //        Assert.AreEqual("Portal Lojas", driver.FindElement(By.Id("tituloPagina")).Text);
        //    }
        //    catch (AssertionException e)
        //    {
        //        verificationErrors.Append(e.Message);
        //    }

        //    print.PrintScreen();

        //    //Realizar login
        //    Login();


        //    wait(By.Id("btnSearch"));
        //    print.PrintScreen();


        //    ////Alterar Filial
        //    Thread.Sleep(2000);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.ESCAPE);
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("labelFilial")).Click();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("changeFilial")).SendKeys("54 - L 054 JD / SP");
        //    Thread.Sleep(2000);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);
        //    Thread.Sleep(2000);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);
        //    Thread.Sleep(5000);


        //    //Validar P�gina Inicial
        //    Assert.AreEqual("PSF Security", driver.Title);

        //    try
        //    {
        //        Assert.AreEqual("Buscar por:\r\nCPF\r\nCART�O", driver.FindElement(By.Id("tipoSearch")).Text);
        //    }
        //    catch (AssertionException e)
        //    {
        //        verificationErrors.Append(e.Message);
        //    }
        //    Assert.IsTrue(IsElementPresent(By.Id("cpfSearch")));
        //    Assert.IsTrue(IsElementPresent(By.Id("btnSearch")));
        //    try
        //    {
        //        Assert.AreEqual("Concessao do Cartao", driver.FindElement(By.Id("menuCollapse3162")).Text);
        //    }
        //    catch (AssertionException e)
        //    {
        //        verificationErrors.Append(e.Message);
        //    }
        //    try
        //    {
        //        Assert.AreEqual("Atendimento", driver.FindElement(By.Id("menuCollapse3164")).Text);
        //    }
        //    catch (AssertionException e)
        //    {
        //        verificationErrors.Append(e.Message);
        //    }


        //    //Cenario 1 - Cadastro de Clientes

        //    //Consultar Cliente
        //    driver.FindElement(By.Id("cpfSearch")).Clear();
        //    driver.FindElement(By.Id("cpfSearch")).SendKeys("5339718379");
        //    Thread.Sleep(2000);
        //    print.PrintScreen();
        //    driver.FindElement(By.Id("btnSearch")).Click();
        //    Thread.Sleep(2000);



        //    wait(By.Id("iniciarCadastro"));


        //    //Validar P�gina Cadastro de Clientes


        //    Assert.IsTrue(IsElementPresent(By.Id("formPreCadastro")));

        //    driver.FindElement(By.Id("dataNascCad")).Click();
        //    driver.FindElement(By.Id("dataNascCad")).SendKeys("10/06/1992");
        //    driver.FindElement(By.Id("iniciarCadastro")).Click();


        //    //Validar Mensagem de Retorno CPF

        //    Thread.Sleep(8000);


        //    //if (this.IsElementPresent(By.Id("dialog-message")))

        //    //{
        //    //    wait(By.CssSelector("#modalContent > div.modal-body.row"));
        //    //    string Retorno = driver.FindElement(By.CssSelector("#modalContent > div.modal-body.row")).Text;
        //    //    // var setMensagemRetorno =
        //    //    Marisa.TestDataAccess.ExcelDataAccess.SetMesagemRetorno(Retorno, Cadastro.CPF);
        //    //    Thread.Sleep(2000);
        //    //    print.PrintScreen();

        //    //    //refazer o contador
        //    //    //int counter = 0;
        //    //    //int NumeroCPF = Marisa.TestDataAccess.ExcelDataAccess.GetSemMensagem().total;

        //    //    //while (counter < NumeroCPF) ;

        //    //}


        //    wait(By.Id("avancarDadosPessoais"));

        //    //Cenario 2  -- Dados pessoais


        //    //Validar P�gina Cadastro de Clientes


        //    Assert.IsTrue(IsElementPresent(By.Id("formDadosPessoais")));


        //    //Nome da M�e
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("nomeMaeDadosPessoais")).Click();
        //    driver.FindElement(By.Id("nomeMaeDadosPessoais")).Clear();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("nomeMaeDadosPessoais")).SendKeys("Maria Aparecida");


        //    //Estado Civil
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("estadoCivil")).Click();
        //    Thread.Sleep(2000);
        //    //new SelectElement(driver.FindElement(By.Id("estadoCivil"))).SelectByValue(Cadastro.EstadoCivil);
        //    driver.FindElement(By.Id("estadoCivil")).SendKeys("CASADO");// PEGAR DO BANCO
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);



        //    //Sexo       
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("sexoDadosPessoais")).Click();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("sexoDadosPessoais")).SendKeys("MASCULINO"); // PEGAR DO BANCO
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);


        //    //Deseja Receber a Fatura por email?         
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("faturaEmailOutros")).Click();
        //    driver.FindElement(By.Id("faturaEmailOutros")).SendKeys("N�o"); // PEGAR DO BANCO //


        //    //Email
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("cliEmail")).Click();
        //    driver.FindElement(By.Id("cliEmail")).Clear();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("cliEmail")).SendKeys("starline@starline.com.br");//PEGAR NO BANCO

        //    Thread.Sleep(2000);

        //    print.PrintScreen();

        //    //Bot�o avan�ar
        //    driver.FindElement(By.Id("avancarDadosPessoais")).Click();

        //    //Cenario 3 -- Dados Residenciais 

        //    //Validar P�gina Dados Residenciais

        //    wait(By.Id("avancarDadosResidenciais"));


        //    Assert.IsTrue(IsElementPresent(By.Id("formDadosResidenciais")));

        //    //Cep
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("cepDadosResidenciais")).Click();
        //    driver.FindElement(By.Id("cepDadosResidenciais")).Clear();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("cepDadosResidenciais")).SendKeys("06033050");
        //    Thread.Sleep(2000);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);



        //    //Numero
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("numeroDadosResidenciais")).Click();
        //    driver.FindElement(By.Id("numeroDadosResidenciais")).Clear();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("numeroDadosResidenciais")).SendKeys("1010");

        //    //Tipo Residencia
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("tipoResidencia")).Click();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("tipoResidencia")).SendKeys("ALUGADA"); // PEGAR DO BANCO
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);


        //    //Telefone Residencia
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("telefoneResidencialDadosResidenciais")).Click();
        //    driver.FindElement(By.Id("telefoneResidencialDadosResidenciais")).Clear();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("telefoneResidencialDadosResidenciais")).SendKeys("1135926538");
        //    Thread.Sleep(2000);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);



        //    print.PrintScreen();

        //    //Bot�o avan�ar
        //    //Thread.Sleep(2000);
        //    //driver.FindElement(By.Id("avancarDadosResidenciais"));

        //    //Cenario 4 -- Dados Comerciais 

        //    //Validar P�gina Dados Comerciais

        //    wait(By.Id("avancarDadosComerciais"));


        //    Assert.IsTrue(IsElementPresent(By.Id("formDadosComerciais")));


        //    //Classe Profissisional
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("classeProfDadosComerciais")).Click();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("classeProfDadosComerciais")).SendKeys("Funcionarios"); // PEGAR DO BANCO
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);

        //    //Atividade
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("atividadeDadosComerciais")).Click();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("atividadeDadosComerciais")).SendKeys("Funcionarios"); // PEGAR DO BANCO
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);


        //    //Renda Mensal
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("rendaMensalDadosComerciais")).Click();
        //    driver.FindElement(By.Id("rendaMensalDadosComerciais")).Clear();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("rendaMensalDadosComerciais")).SendKeys("10000");


        //    ////Tempo de Empresa (ANO)
        //    //Thread.Sleep(2000);
        //    //driver.FindElement(By.Id("tempoEmpresaAnoDadosComerciais")).Click();
        //    //driver.FindElement(By.Id("tempoEmpresaAnoDadosComerciais")).Clear();
        //    //Thread.Sleep(2000);
        //    //driver.FindElement(By.Id("tempoEmpresaAnoDadosComerciais")).SendKeys("10");



        //    ////Tempo de Empresa (MES)
        //    //Thread.Sleep(2000);
        //    //driver.FindElement(By.Id("tempoEmpresaMes")).Click();
        //    //driver.FindElement(By.Id("tempoEmpresaMes")).Clear();
        //    //Thread.Sleep(2000);
        //    //driver.FindElement(By.Id("tempoEmpresaMes")).SendKeys("10");


        //    //CEP
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("cepDadosComerciais")).Click();
        //    driver.FindElement(By.Id("cepDadosComerciais")).Clear();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("cepDadosComerciais")).SendKeys("06033050");

        //    //Numero
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("numeroDadosComerciais")).Click();
        //    driver.FindElement(By.Id("numeroDadosComerciais")).Clear();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("numeroDadosComerciais")).SendKeys("365");

        //    //Telefone Comercial
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("telefoneComercialDadosComercial")).Click();
        //    driver.FindElement(By.Id("telefoneComercialDadosComercial")).Clear();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("telefoneComercialDadosComercial")).SendKeys("1136811452");
        //    Thread.Sleep(2000);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);

        //    //Bot�o avan�ar
        //    //Thread.Sleep(2000);
        //    //driver.FindElement(By.Id("avancarDadosComerciais"));

        //    //Cenario 5 -- Outros 

        //    //Validar P�gina Outros

        //    wait(By.Id("avancarOutros"));


        //    Assert.IsTrue(IsElementPresent(By.Id("formOutros")));

        //    //Vencimento Melhor Dia
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("vencimentoOutros")).Click();
        //    driver.FindElement(By.Id("vencimentoOutros")).SendKeys("10");
        //    Thread.Sleep(2000);

        //    //Confirma��o de Vencimento
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.ESCAPE);
        //    Thread.Sleep(2000);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.ESCAPE);
        //    Thread.Sleep(2000);
        //    //Cobrar tarifa de overlimit


        //    wait(By.CssSelector("#tarifaOverlimitOutros"));

        //    driver.FindElement(By.CssSelector("#tarifaOverlimitOutros")).Click();
        //    driver.FindElement(By.CssSelector("#tarifaOverlimitOutros")).SendKeys("N�o");

        //    //Conta B�nus?
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("tipoCartaoOutros")).Click();
        //    driver.FindElement(By.Id("tipoCartaoOutros")).SendKeys("N�o");


        //    //Campanha
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("campanhaOutros")).Click();
        //    driver.FindElement(By.Id("campanhaOutros")).SendKeys("931 - Desc Primeira Compra Marisa Itaucard 10%");
        //    Thread.Sleep(2000);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);


        //    //Bot�o avan�ar
        //    //Thread.Sleep(2000);
        //    driver.FindElement(By.Id("avancarOutros")).Click();

        //    //Cenario 6 -- Camera 

        //    //Validar P�gina Outros

        //    wait(By.Id("avancarBiometriaFacial"));






        //    //   Assert.IsTrue(IsElementPresent(By.Id("fieldSetBiometria")));

        //    //Ligar Camera
        //    Thread.Sleep(2000);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);


        //    Thread.Sleep(2000);

        //    driver.Quit();


        //}

        //////Tirar foto
        ////driver.SwitchTo();
        ////    driver.FindElement(By.Id("btnCamera")).Click();
        ////    Thread.Sleep(200);
        ////    driver.FindElement(By.CssSelector("#pnlCapturePreview > div > div.section-content.buttons > div > div.imagebutton.accept.dark")).Click();
        ////    Thread.Sleep(20000);

        ////    //Bot�o Avan�ar
        ////    driver.FindElement(By.Id("avancarBiometriaFacial")).Click();

        ////    //Cenario 7 -- Formaliza��o 

        ////    //Validar P�gina Outros

        ////    wait(By.Id("avancarBiometriaFacial"));


        ////    Assert.IsTrue(IsElementPresent(By.Id("panelAssinaturaEletronica")));

        ////    driver.Quit();


        ////}

        //[Test]
        //public void CT006_Cadastro_de_Clientes_Liberais()
        //{

        //    //M�todo de Print
        //    ScreenShot print = new ScreenShot("VendaCartoes", "CT001_Cadastro_de_Clientes");

        //    //M�todo simulador de teclado
        //    InputSimulator simulator = new InputSimulator();

        //    //Execu��o

        //    //Acessar P�gina do CCM
        //    driver.Navigate().GoToUrl(baseURL);


        //    wait(By.XPath("//button[@type='submit']"));
        //    wait(By.CssSelector("img.png_bg"));
        //    print.PrintScreen();

        //    try
        //    {
        //        Assert.AreEqual("Login", driver.FindElement(By.CssSelector("strong")).Text);
        //    }
        //    catch (AssertionException e)
        //    {
        //        verificationErrors.Append(e.Message);
        //    }
        //    Assert.IsTrue(IsElementPresent(By.Id("login_username")));
        //    Assert.IsTrue(IsElementPresent(By.Id("login_password")));
        //    Assert.IsTrue(IsElementPresent(By.XPath("//button[@type='submit']")));



        //    //Validar Tela de Login
        //    Assert.AreEqual("PSF Security", driver.Title);
        //    Assert.IsTrue(IsElementPresent(By.CssSelector("img.png_bg")));
        //    try
        //    {
        //        Assert.AreEqual("Portal Lojas", driver.FindElement(By.Id("tituloPagina")).Text);
        //    }
        //    catch (AssertionException e)
        //    {
        //        verificationErrors.Append(e.Message);
        //    }

        //    print.PrintScreen();

        //    //Realizar login
        //    Login();


        //    wait(By.Id("btnSearch"));
        //    print.PrintScreen();


        //    ////Alterar Filial
        //    Thread.Sleep(2000);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.ESCAPE);
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("labelFilial")).Click();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("changeFilial")).SendKeys("54 - L 054 JD / SP");
        //    Thread.Sleep(2000);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);
        //    Thread.Sleep(2000);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);
        //    Thread.Sleep(5000);


        //    //Validar P�gina Inicial
        //    Assert.AreEqual("PSF Security", driver.Title);

        //    try
        //    {
        //        Assert.AreEqual("Buscar por:\r\nCPF\r\nCART�O", driver.FindElement(By.Id("tipoSearch")).Text);
        //    }
        //    catch (AssertionException e)
        //    {
        //        verificationErrors.Append(e.Message);
        //    }
        //    Assert.IsTrue(IsElementPresent(By.Id("cpfSearch")));
        //    Assert.IsTrue(IsElementPresent(By.Id("btnSearch")));
        //    try
        //    {
        //        Assert.AreEqual("Concessao do Cartao", driver.FindElement(By.Id("menuCollapse3162")).Text);
        //    }
        //    catch (AssertionException e)
        //    {
        //        verificationErrors.Append(e.Message);
        //    }
        //    try
        //    {
        //        Assert.AreEqual("Atendimento", driver.FindElement(By.Id("menuCollapse3164")).Text);
        //    }
        //    catch (AssertionException e)
        //    {
        //        verificationErrors.Append(e.Message);
        //    }


        //    //Cenario 1 - Cadastro de Clientes

        //    //Consultar Cliente
        //    driver.FindElement(By.Id("cpfSearch")).Clear();
        //    driver.FindElement(By.Id("cpfSearch")).SendKeys("4491536902");
        //    Thread.Sleep(2000);
        //    print.PrintScreen();
        //    driver.FindElement(By.Id("btnSearch")).Click();
        //    Thread.Sleep(2000);



        //    wait(By.Id("iniciarCadastro"));


        //    //Validar P�gina Cadastro de Clientes


        //    Assert.IsTrue(IsElementPresent(By.Id("formPreCadastro")));

        //    driver.FindElement(By.Id("dataNascCad")).Click();
        //    driver.FindElement(By.Id("dataNascCad")).SendKeys("11/02/1967");
        //    driver.FindElement(By.Id("iniciarCadastro")).Click();


        //    //Validar Mensagem de Retorno CPF

        //    Thread.Sleep(8000);


        //    //if (this.IsElementPresent(By.Id("dialog-message")))

        //    //{
        //    //    wait(By.CssSelector("#modalContent > div.modal-body.row"));
        //    //    string Retorno = driver.FindElement(By.CssSelector("#modalContent > div.modal-body.row")).Text;
        //    //    // var setMensagemRetorno =
        //    //    Marisa.TestDataAccess.ExcelDataAccess.SetMesagemRetorno(Retorno, Cadastro.CPF);
        //    //    Thread.Sleep(2000);
        //    //    print.PrintScreen();

        //    //    //refazer o contador
        //    //    //int counter = 0;
        //    //    //int NumeroCPF = Marisa.TestDataAccess.ExcelDataAccess.GetSemMensagem().total;

        //    //    //while (counter < NumeroCPF) ;

        //    //}


        //    wait(By.Id("avancarDadosPessoais"));

        //    //Cenario 2  -- Dados pessoais


        //    //Validar P�gina Cadastro de Clientes


        //    Assert.IsTrue(IsElementPresent(By.Id("formDadosPessoais")));


        //    //Nome da M�e
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("nomeMaeDadosPessoais")).Click();
        //    driver.FindElement(By.Id("nomeMaeDadosPessoais")).Clear();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("nomeMaeDadosPessoais")).SendKeys("Maria Aparecida");


        //    //Estado Civil
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("estadoCivil")).Click();
        //    Thread.Sleep(2000);
        //    //new SelectElement(driver.FindElement(By.Id("estadoCivil"))).SelectByValue(Cadastro.EstadoCivil);
        //    driver.FindElement(By.Id("estadoCivil")).SendKeys("CASADO");// PEGAR DO BANCO
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);



        //    //Sexo       
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("sexoDadosPessoais")).Click();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("sexoDadosPessoais")).SendKeys("MASCULINO"); // PEGAR DO BANCO
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);


        //    //Deseja Receber a Fatura por email?         
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("faturaEmailOutros")).Click();
        //    driver.FindElement(By.Id("faturaEmailOutros")).SendKeys("N�o"); // PEGAR DO BANCO //


        //    //Email
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("cliEmail")).Click();
        //    driver.FindElement(By.Id("cliEmail")).Clear();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("cliEmail")).SendKeys("starline@starline.com.br");//PEGAR NO BANCO

        //    Thread.Sleep(2000);

        //    print.PrintScreen();

        //    //Bot�o avan�ar
        //    driver.FindElement(By.Id("avancarDadosPessoais")).Click();

        //    //Cenario 3 -- Dados Residenciais 

        //    //Validar P�gina Dados Residenciais

        //    wait(By.Id("avancarDadosResidenciais"));


        //    Assert.IsTrue(IsElementPresent(By.Id("formDadosResidenciais")));

        //    //Cep
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("cepDadosResidenciais")).Click();
        //    driver.FindElement(By.Id("cepDadosResidenciais")).Clear();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("cepDadosResidenciais")).SendKeys("06033050");
        //    Thread.Sleep(2000);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);



        //    //Numero
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("numeroDadosResidenciais")).Click();
        //    driver.FindElement(By.Id("numeroDadosResidenciais")).Clear();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("numeroDadosResidenciais")).SendKeys("1010");

        //    //Tipo Residencia
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("tipoResidencia")).Click();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("tipoResidencia")).SendKeys("ALUGADA"); // PEGAR DO BANCO
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);


        //    //Telefone Residencia
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("telefoneResidencialDadosResidenciais")).Click();
        //    driver.FindElement(By.Id("telefoneResidencialDadosResidenciais")).Clear();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("telefoneResidencialDadosResidenciais")).SendKeys("1135926538");
        //    Thread.Sleep(2000);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);



        //    print.PrintScreen();

        //    //Bot�o avan�ar
        //    //Thread.Sleep(2000);
        //    //driver.FindElement(By.Id("avancarDadosResidenciais"));

        //    //Cenario 4 -- Dados Comerciais 

        //    //Validar P�gina Dados Comerciais

        //    wait(By.Id("avancarDadosComerciais"));


        //    Assert.IsTrue(IsElementPresent(By.Id("formDadosComerciais")));


        //    //Classe Profissisional
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("classeProfDadosComerciais")).Click();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("classeProfDadosComerciais")).SendKeys("Liberais"); // PEGAR DO BANCO
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);

        //    //Atividade
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("atividadeDadosComerciais")).Click();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("atividadeDadosComerciais")).SendKeys("Liberais"); // PEGAR DO BANCO
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);


        //    //Renda Mensal
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("rendaMensalDadosComerciais")).Click();
        //    driver.FindElement(By.Id("rendaMensalDadosComerciais")).Clear();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("rendaMensalDadosComerciais")).SendKeys("10000");


        //    ////Tempo de Empresa (ANO)
        //    //Thread.Sleep(2000);
        //    //driver.FindElement(By.Id("tempoEmpresaAnoDadosComerciais")).Click();
        //    //driver.FindElement(By.Id("tempoEmpresaAnoDadosComerciais")).Clear();
        //    //Thread.Sleep(2000);
        //    //driver.FindElement(By.Id("tempoEmpresaAnoDadosComerciais")).SendKeys("10");



        //    ////Tempo de Empresa (MES)
        //    //Thread.Sleep(2000);
        //    //driver.FindElement(By.Id("tempoEmpresaMes")).Click();
        //    //driver.FindElement(By.Id("tempoEmpresaMes")).Clear();
        //    //Thread.Sleep(2000);
        //    //driver.FindElement(By.Id("tempoEmpresaMes")).SendKeys("10");


        //    //CEP
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("cepDadosComerciais")).Click();
        //    driver.FindElement(By.Id("cepDadosComerciais")).Clear();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("cepDadosComerciais")).SendKeys("06033050");

        //    //Numero
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("numeroDadosComerciais")).Click();
        //    driver.FindElement(By.Id("numeroDadosComerciais")).Clear();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("numeroDadosComerciais")).SendKeys("365");

        //    //Telefone Comercial
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("telefoneComercialDadosComercial")).Click();
        //    driver.FindElement(By.Id("telefoneComercialDadosComercial")).Clear();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("telefoneComercialDadosComercial")).SendKeys("1136811452");
        //    Thread.Sleep(2000);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);

        //    //Bot�o avan�ar
        //    //Thread.Sleep(2000);
        //    //driver.FindElement(By.Id("avancarDadosComerciais"));

        //    //Cenario 5 -- Outros 

        //    //Validar P�gina Outros

        //    wait(By.Id("avancarOutros"));


        //    Assert.IsTrue(IsElementPresent(By.Id("formOutros")));

        //    //Vencimento Melhor Dia
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("vencimentoOutros")).Click();
        //    driver.FindElement(By.Id("vencimentoOutros")).SendKeys("10");
        //    Thread.Sleep(2000);

        //    //Confirma��o de Vencimento
        //    SendKeys.SendWait("{ESC}");
        //    Thread.Sleep(2000);
        //    SendKeys.SendWait("{ESC}");
        //    Thread.Sleep(2000);
        //    //Cobrar tarifa de overlimit


        //    wait(By.CssSelector("#tarifaOverlimitOutros"));

        //    driver.FindElement(By.CssSelector("#tarifaOverlimitOutros")).Click();
        //    driver.FindElement(By.CssSelector("#tarifaOverlimitOutros")).SendKeys("N�o");

        //    //Conta B�nus?
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("tipoCartaoOutros")).Click();
        //    driver.FindElement(By.Id("tipoCartaoOutros")).SendKeys("N�o");


        //    //Campanha
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("campanhaOutros")).Click();
        //    driver.FindElement(By.Id("campanhaOutros")).SendKeys("931 - Desc Primeira Compra Marisa Itaucard 10%");
        //    Thread.Sleep(2000);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);


        //    //Bot�o avan�ar
        //    //Thread.Sleep(2000);
        //    driver.FindElement(By.Id("avancarOutros")).Click();

        //    //Cenario 6 -- Camera 

        //    //Validar P�gina Outros

        //    wait(By.Id("avancarBiometriaFacial"));






        //    //   Assert.IsTrue(IsElementPresent(By.Id("fieldSetBiometria")));

        //    //Ligar Camera
        //    Thread.Sleep(2000);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);


        //    Thread.Sleep(2000);

        //    driver.Quit();

        //    SendKeys.SendWait("^+(J)");
        //    Thread.Sleep(2000);
        //    SendKeys.SendWait("acessoFlash.faceDetect('data:image/jpeg;base64,/9j/4AAQSkZJRgABAQEAYABgAAD//gA7Q1JFQVRPUjogZ2QtanBlZyB2MS4wICh1c2luZyBJSkcgSlBFRyB2ODApLCBxdWFsaXR5ID0gOTAK/9sAQwADAgIDAgIDAwMDBAMDBAUIBQUEBAUKBwcGCAwKDAwLCgsLDQ4SEA0OEQ4LCxAWEBETFBUVFQwPFxgWFBgSFBUU/9sAQwEDBAQFBAUJBQUJFA0LDRQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQU/8AAEQgB9AH0AwEiAAIRAQMRAf/EAB8AAAEFAQEBAQEBAAAAAAAAAAABAgMEBQYHCAkKC//EALUQAAIBAwMCBAMFBQQEAAABfQECAwAEEQUSITFBBhNRYQcicRQygZGhCCNCscEVUtHwJDNicoIJChYXGBkaJSYnKCkqNDU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6g4SFhoeIiYqSk5SVlpeYmZqio6Slpqeoqaqys7S1tre4ubrCw8TFxsfIycrS09TV1tfY2drh4uPk5ebn6Onq8fLz9PX29/j5{+}v/EAB8BAAMBAQEBAQEBAQEAAAAAAAABAgMEBQYHCAkKC//EALURAAIBAgQEAwQHBQQEAAECdwABAgMRBAUhMQYSQVEHYXETIjKBCBRCkaGxwQkjM1LwFWJy0QoWJDThJfEXGBkaJicoKSo1Njc4OTpDREVGR0hJSlNUVVZXWFlaY2RlZmdoaWpzdHV2d3h5eoKDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uLj5OXm5{+}jp6vLz9PX29/j5{+}v/aAAwDAQACEQMRAD8A{+}ysEEd6UKMUoFKB6VzmoEZ7UpBHNKBjp1pcHrigBQgxSr{+}tHPrR3680FJAQc04DJx3pcZwe1KR0x{+}dADSMYP508LkDOBSgH604Ad{+}nvQITGf8aAuOOxp2D74pwHAFAhqjAxjinbenrTgABRj0oATZxmlwKUDINKBngUANIPbpTgvJ60/Hy0DNADQm0U4Lx0NL1Jo6CgBAvB7Um3jjpTyKMZ7UAIeeKMY5pQRz60uc9eBQA3qKbg9qeRjGO9ZuteIrDQLfzbucKTwkYxvc9gB3NAzR6UY/CvJ/FP7Rvh3w/Ym4jhu53UkPFJF5RXB{+}bO7HTjpnrXjJ/bKtvF3iY2dhetoIiRntpn5hmJwCkwI454BHfrjurlKLPr48LRXzp4U/a10yW/GmeIJYLe5kcxx3ETDZ5mf9U2DwcBiM44BHUZbgF/bwXSfG82k6jbxXNqjGDfapklwcbgSRlTlT2PWlzIfKz7KB9aQ4r5z0b9q2114XF3bqrWMR42LkhDvUFs/dYtG3GemPcjJv/2r/wC15mfS41NqjYJdShAAB388YJBwMcj60udDUGfUQUEZByKTHH614t4P{+}N/9rwLHO8bzkKyrbKcMCOSR1B//AF9K7rT/AB9a3yZgZZl3EBklXY56YDHr{+}HFNSTJcWjrsACgjOQRWfZ63bzIGdhDnqHPH51qDBGTgVRJERwM5FGB16mpCucZ5z0pNo7/SgRGPel28cUpGRSbSQBQA3bgY60DnFPH1oxnkUAMxgCmhRg9vepMU3nsKAGbeKTbgVJ25pMc{+}tAEOM4oPHSpdtNx69KAI/wAqRhkipD1pmMEmgBhXHHQUmNpx6mpaaVLexoAjOCO/4UzGD609htoI5oAjOV7UHJPtTyMmmkcUAMYYNNIAFPAyPfPrTTzmgCNuKYRgZxU2OOcUbaCrkA4xzRjt0z3qRlweKYcfWgLjCoB5NFPwDyOaKBloAk88e1L0ODxS529qXGcZoJsIBgDjFP8A50HhvUUu0EmgBOWpwUKM4z70u3sOtOCY/GgBAM9KcBjHenAY96XHGKBAB6UY/P3py/r2pcZoAAKVVJJPSnKo7Ud8UABUD3oHSnAc{+}1OAoAaAeeAaXHPvS9/0oIA79KAEJ74pRyRijPalFAwxyaNtA4OKO/vQAE88UvTmgAk8Vn6zrlpoVsZruVYx2DMB{+}poC1y3LIkSlnIUVz{+}t{+}NtN0clbi6WJ8Z2Zyx/AfhXknjn45RQG4eKXMcakpCmVY{+}{+}OCM9iR{+}FfL/jz44re3chsSLuWSJlaPYWCH/gWR19qydRLY2jSbPbfip{+}1XcaJqSw6JDcXVt9151Awpx6EjPPvXz547/ah8SeLJlsTd/wBnbyY5JI23YyMfdPAz1z7da4K98WatNJJJPOgWUMrRsVcbDn5cdB17CuJ1Ywx3LyyLuUkEN0z06UlJyNuRJHUf8Jt4zsNGu2GrHUI3fy5bXUHEsTcfKCjcq2A2GBOeQMV5t/bk{+}m6g15ChjLffQEyQyAjBBPcHPIPr69Ld14n0u0EkSQqrMuxnExYFfQjGKxtQ8Q20loI1BRMYXfGOc9cYGa0WpDXYt3V9c6hNJe2kzrI{+}BNCzlnBB4YE9ccZ79euTUGurN50N4gKO3zFASdvHPPX8Kx4vEAiKsrpJtPGVK9a2x4ogeUSbg25Nx45zzkGrsTddTQ0nxLqmk20wW4kRLtQ7KhPY7hkfifzq5c{+}L9Wg0uKBJJ5PtDqzxY{+}V8Y2rjvyP51Ui1rS7pYXnB{+}U79sePmAPAOemeK1LbWZtSjDpbxRpyPMdsAd8etQ0aRV9mdR4a{+}KHiPSdCmtrm{+}nneQEN821FyMFSe4AxhRgdc9Oe2{+}GvxNh026dtVv5ruJcqENx5SIeehPzH8Oeeteax3kKwGKS/2x9dsKbQc9eWIHr2q0jaPIuBdT8AgbQTn8RkYPvWZSjY{+}xvhv8bfA99G9leaxZ285YgK7HbKuONpbkntz6dK9r0bW45reFtIujFnBCRzAwP2GCOV/AAeor8yG0PS7gs/2yNhySkpKN79scVs{+}GPEvijwiBJ4W8TyafEcbYSziFvqCCPrxT5kiXC5{+}p{+}j{+}LJZ3MVzbypKpw6ylQyfXB5Hoe/wBeK6WG4WZcjHX1zzXwB4O/bF1XTHt4vGenxIQAi6lY4ZR/tfK2Meo/IV9U{+}BviQnjGwt7zRruzvkkTdEIZBukHQjacbh9MfnVKSZhODR60cY6UY6VkaV4ijv51tp4JLK9xkwyc7gOpU9xWuDz61ZmJtFJ0PtTu1FAhCAB0zSFSBQcjgUoOetADCuaaUINS{+}1Z3iLVRoHh/U9UaIzLY2styYwdu8IhbGe2cUAXCtJgEHpXmf7O/xwtf2gPhZp3jWLTG0KG9nmgWzmuBKQY5Cn3sLnO3OMV6cyYHWgCM8D1pCvX60u5c7cgsOoB6U1ZElBKMrgHB2nOD6UAN9ecYo9K5vXfAFnrOptqVveXuk6myhWurCYoXA4Acchqoz6V4p0SymlbxXaTwRIXMl/p4GxQMkllYfniueVScW7w08mv{+}Ackq9Sm3zU3bumvxvY6yW4himhieVElmJEaM2C{+}Bk4HfAp2Ac56V8/T3nivU9B1jxzc2uo6/HplrJ9h0jSN0E{+}o4IzHFtUsobAyQMnGO3HsPw88RXfivwRousX2hXXhm7vLZZZNHvcia0P8AzzfIU5H0FKhWdePPy2XQjB4mWKi6jhyxe192u5vDPQDgUh5FKkiSLvRw6k43KcivG/2l/g/42{+}L{+}iaNaeCfiDe/D{+}6s7h5bi5spZozcIVwEPlOpIB55zXSd57Fjk008E{+}1fmR{+}xb4c{+}MHx81h/Ecnxm8R2{+}meG9YtftmmXupXc63sYYSMh/e4wyqVIII5r2P9n/x34q{+}Gn7Z/j/4TeLfEmra7pepxNe6A2r30tyY0XM0aRmRjj9y8gbHUwe1VYm59osDzjFJjOQcYr5N/wCCiPxf1/wP4H8L{+}EvBuo3uneLfFOqLFBJpk7Q3AhjIyqOpBUtI8K9eRuHrX0p8P/Dd54R8DaDo2o6nc6zqNjZRQXWo3kzTS3MwUb5GdiSdzZPPrSsM3Svp0puOPwp/8RpCOwxSGRZC8Y/Kin7d3aigdy2OTz0o6ilC807HpQMQen61IBQq8A07oKCWJjBx39acBRt5zk08AnHFAhqipFX3oAHHrTgcnFACFQKMY/GnhBz0NLkD60AJ0ApwHXNIBxinY5oAQj0NKB0oI3Z7UYoGLSdM8ZzThzRt96BpCY5NAGR6Up9KcBxQDG4xRjLU7O2vOfiv8V7bwJp7QwNHJq0qM0aOSUhUdXfHQDsOM0m7Ak3oi78R/itpXw/s281jdai4Jhs4cNI57cfXuePWvjP4vfHPVtR1JG1K6Tz2fdDpNoykRAcgyOfunPcc{+}nGK8{+}{+}JXxpe8ur2LTbqa{+}v5pMXOpzEHeeDhQOAuQcKMjivHrvV5J2bYfPvJGy1xISVDexHLH{+}tcrk5Ox2QpqKuzuNd12a/RzfX8cClyzRQAqpB67j1J46MfTgVz8{+}viRZE02Jyqr887ybcH1JPA/n6VUsLc3siwRQNq1ymczz/JDEfoD264H41sXljpOgobjxPq4mmABis4R37BYx29249B3qdIvzNldmFGWv4WEk4H{+}3EjON31JXP61y{+}s6ZZacrPdPlucBsKx{+}uc4{+}ldJrPjzV9VRrTQdNj063xta4kO6UL0GWPC/z{+}tefyaG19ds91cyagxOA6ElSe4BPJ/CtaalvLQym09I6mfJqGntI8iQ{+}bjj5cnd6CqV7ezzHesQjJHRu1dBdaZFpcCmZ4rdv7gOSv8An1Nc5e3lvG5EOWH/AD0f19hXWmuhzSutyuqTN/rnGPYVqabYQzXCwuxYAEsQcHODgVlWxN1couCQCM59M1t{+}HrS4u73bGoLP3HNEnoTFXEtNtrclJTjYSBnnHpWjBqwkPlpyx{+}7xnPr9D{+}FVNes/7PuXgkBLZ5Kj8v1rBeZoZw{+}4kAjuR/Kl8SLb5XY7IeLZokeAQwuoYksyq34/dzSDWLW6iCFDCwPXoT/WuQffOzTLCSMljsGRgfyxU0LQSHcztH6FcHB9x3/MUuRIFNna21leLkxs8iHLhW7{+}mfX8c1asfEc1g3lyQBOMNKh2t{+}XXP0rB8P67f6BcpPEPOt9wyUG4EjsQRwfpXf6a{+}heMYN7TRWN2/G9f9WTnow7ZrGUuV{+}8tDogufWL1LlhdnxDbF7O5tp2C4NvLIN2M88Hr{+}FXvDHi7WvhxqkV74d1KfS7tWJNpM{+}YCw7Ecj17VyGp{+}G73wzfEL{+}5lyGWTGVYeoPQj9PXFbdl4i03xIqWNwbfS9UPUSswikbHGRg7CfUcf7OOazdlrEvfSR93fs/wD7U{+}nfE22i0TxMyaZ4lgAXc52rKegKNnr6EHn8a{+}n7C6EkSB33sR8r9Q3/ANevxde5k03Ultr1WsLpMPDKrBSM85Rx1B/EGvrv9m79q680CSHw54uumvdOY4h1E5Z4TgZ3jqV557gc{+}hq1O25hOn1R95Y60mDnsKpaNq0Gr2kcsMqSoyh0dDkOvqCKvkYrY5Rh4IzSNk9KeRnGRSbT68UCGHp1rnviM3/FvPFA/wCoXdf{+}imroyvX6VleJ9IbxD4c1XS0lELXtpLbCUjIQuhXOO{+}M0AfnN{+}xz{+}xF4W{+}PnwDsPEvjTX9buvNmubfSrGxvBHDpiLKwZlUqw3tJuc9sEcZ5rH1Px94qs/2Rv2h/hvrWu3OtzfDzXrDTbHWXkbzXgbURHs3ZzhTAxAJJAfbnAAHsPgj9iD4zfBXwwNG{+}G/xtj0q2vt7apbXWngwiQsQJbfIcxsY9gOMHK53YwF7dP2EbPS/wBl/wATfC7TvEjSeIPEt3Bf6p4mvrcu086Txy58vdkLiMgAsTlmYkkmtLk2Pnzxx4R1n4VfB34OeAfDXi3VINY{+}NOo21zr{+}v3VxmRS8dshiRhhhH/pC5GdzCPGcMRVv9rT9jfQP2cfgJqniXwD4n8QabIj2trq1rc326LU42mQKWVVXDrJscY4wrcd6{+}p/jD{+}yZp3xf{+}B3hXwVeavJpuveGba2Gma/axndDPFEsZbZuB2NtyQGBBCnOVrxrxl{+}xH8ZfjP4RfRPiT8bU1W3stjaXbWungQmUMAZbggI0hEZcDOTls7uCGSYWPrX4WlpPhn4SdiWZtHsySTyT5Kc1yWt6hJ8VvEx8PadIy{+}HLFw{+}pXUZ4nYHiNT6ZH6E9hnM8T6tqWh{+}HdG8A6K7vPaWdvY32qFCkcahFjznnaGxyfwGT06nw1qnhnwBocGl2tw1xs5mmihZvMkJZSScesbADsAK8mrVjWqeyvaK38/L/ADPCr4iGJqvD81oR{+}J9/7q/X7jmf2tIxpn7L3xEWz/0UW{+}iSiLyjt8sADGMdMV8Z32reIPix4f8A2aPgfDr95oOheJNBGo6zeW0pWW7jXzT5ZY9flgkwDkFnUkHaK{+}5PjVosHxT{+}FvjPwfHdTaa{+}oWUtm189nJKkROBkKMF{+}TjA68{+}hrwPxn{+}ylpHib4UfDjRtO8ct4e{+}IHgOBYtM8SxWkkKtgglXUkFRuVSDuypzwckH0VUprS6PX9tSWnMjxP9rv8AZh0r9mjwr4R1DwL4i1u00TUvEFtaX2i3d6ZIpZgkjR3C4AwwVZFPX74xgZB/SwLjtXxL4v8A2IPip8Z7bSL/AOIvxki1vUtLuo5bCC201Vso4ersQnl7pGwnzY4APJyMfSem/D/xfafHTVfF9x40mufBlzpy2tv4UMbeXbzjy8zBt2MnY/b{+}OtHqjoifKf8AwSdGfBnxF/7CsH/otq1/{+}Cg{+}g3Xw68XfDP456PCz3nhvUorLURGOZbcsZIwT2U/voz/12Ar1r9kH9l28/Zh0LxLYXniCDxA2r3cdyskFsYBFtUrggs2etem/GX4Y2Xxj{+}FviPwbfOIYdWtGhScpv8mUHdFJjjO11RscZxRfW4W0Pj/wjd237VX7f0viS0lXUPBfw{+}0{+}I2ky8xSz4JjP186SRwe4tx9K{+}7T3B614h{+}yN{+}zFB{+}zD4I1TSX1OPW9V1K8{+}1XOoRwGEMiqFjjClicL8x69XavczyMmkwRFijocDHNKw9OKTAH1FIBrIc8GilNFAFvoen4U8D1oUZNO2jGTQO4m3JPangY{+}tA7U4DpQIQCnhBnpRgAdOaeCPWgBAuTnPHpTqTG7qKcORzxQAEcgUBR260fdpwoAbjnFOUZpSvB9aUcdaCrCY4FKDn2pQe9GMmgYntShcmlC4NL0oFcQ9MUD1pc84qlq2q2{+}i2E97cvshhUs2Bkn2A7k0CMLx54xg8K6XI5eNbgoXUOeFAIG5vbJH16V{+}cv7Svxxi1nVrvRdNuGkjLZvr/ILv2CL6ep9zjgDn0L9rP42NpSTWFrdJNrF7hrl1{+}ZYQMhY19AmSM/xOSeq18PzXD3E7PLI26QmTLnknu2O/f2GaxfvM6YRUdTXk1J7{+}4dUQQwgn9zv6DPc9f8{+}9b2kaPLeeXJLI0duQTiIhS44yPZRnk1k6PpQeNJbiJlhJyFb7x{+}uOnX1zzgZOTXoFlpn2ezWedTawyR7sv8u1QRh39v7qAelZyfLojpinLVmffaxdafa/2foCN9oC7GnP3VXPb2HqfwFZVl4btrVXmnlN3qLnzJb2V{+}hJ6DJ449OTxW3YW6a9dMI7aSG2Y56ZeUHpu75PPHU/y7p9L0fwbbw6prsSXEykSWlhwyRDpvcZwT7HjnvWTnyaLc0UObXocZD4JiksEurxzbaWBuNxLHgOuOSkfUnHQtXI{+}LvHdjBE9h4XsjbWwXbLfTnfNKPd{+}ij2GB6ZrV8XeI9X8aXE9zP50FmzkpbLyzjt9K4zU/DtxIxMwxHjiBGz{+}eP8{+}9awgr3m9TOUrK1NHK3Km7PmlzPLk5y2fzNZps3BLEhf0roLgJpcbxzRhHf5RGBk/U/p{+}ZrInLzsElAiQc4H3vxNdq8jia7jbOJVDlSSeQuD3rsfAOnvLq0K7Q4Y8kDkD1Brl4QoJRCu4cheR29K7L4eh01KL940bZGVCckVE37rLprVGv8UNKFpqUpSOREIG0tjO3HXHbv9eTXkl4PLcqecHoeRivob4gW9vql2GDxiSQBfndSwOOmMgDHHUivEdY0aS3u5V3R7Qfvhs5/Ks6MlaxpWjrcw4JGRgI3ZOc7c96mluWmdTIoUr8u9RtLH1PbNJ5VuG/eTkt6BDx{+}eKvQWVtOFY5mX{+}8eo/Kui6OZJ9CbSbmQvuV9roMgYBBPv7V0VhewTOjCOOzuAT{+}/jO0P7MOn4HrWDFZz21tK1ukM8C/eVic/lmrui62A6xvZwSbuNsiZYe3PUf5xUO0jSLa3PVtD8Tm8sxp2rwGeA52SLnKdOQDyPw9sg4qDxL4J0{+}{+}soLlvmhlbbDdR4zEf7rcH8ue9clpt5cTTJNZ7XAwTAiDeuD1QfxD26/Tt33h3xJbamxSZI1lCeXLbO21ZV6kg5{+}Vs4wT0PXjmuRxcNUdsZKSszgbuK/8PmO210NqegO{+}Eu4jueA84IP8JHXB4PtQLuTQHt2F695pRYG31KMfNC3bcBzx6fXHfPq2q2YsGgjhn{+}3abcf6ppEXEyn70cg6Bwexrz7U9At9KtpH0uXzoptxl0{+}dThQDz69sHrxj8QJ8wnG2x9RfsrftLXWi6nD4f1WU/ZjgYVsqQSNska9{+}uSo98eg/QHSdUg1W0jnhkVw6hvlOQQehB7ivxJ0m8m08oLZmgER82Fsk7O7LkdRjJ{+}nT3{+}5/2Yf2n5NRW30fWrgR3UbbC0pAHJA5J45zx2/PIuL5dGc9SnfVH22wGM03n86hsL6O/tkmjP3hkrnoas9TW5yEZHFIelSN1ppGRigRHgtk0oHHTBp2OKBjNAARkc9qaVB5p{+}MimkDmgDkV8AWlzqM11qDNchpXkjjV2A{+}YjlueSMDHYY{+}mOktrSCxj2QQpCg/gjUAfpVhl5oK4HFZxpwh8KMadGnSu4ojIzzUckayqVdQynggjINTAY96b0Oa0NrEe0KAoAAA49KYeAe9SkA5GOaaUwKBkZOMdD9KYQc56VKwwOlNJxQCGA8U0/ep{+}M9KaQc89aBjXGabtp/WkYZoEMxRQ3WigRe28YpQM9qAOtOHygDvQAbexpQo496cBk5xTjxQAgXb0p2MZoxmlHIoAAO{+}aXHWjGRSqM59KBoAM9qMd{+}1PIxSdRQOwZ4zSZzSjp1p2BjHagY0LT6Tbgj0p3eglsTHWjGKMc0cHjpQIaxA{+}ueleKfHn4hWfh7Tr6S6neKy0qMT3TRnDNI2BHEn{+}0xIGe3zH{+}HFeua3fR6Zp1xdSnaka54OCT2H4nA/GvzB/az{+}LMmoXg0iG7eZY5GubjaTiS4bJyf90HaB257msqj2SNqcb3Z4X8RfG02ua5d3tyPMvZ5CwgUZUc8KB{+}P6/nj6VZXDOuSrXDENNIfuxjsPqOcD{+}fAqnoto{+}ral5rfNs7ngKD/k/5Fehx2Nn4d003F31XkQMMNuI9D/Ef/HB78DOTUFZbnVFc7v0NDSNOsdDtGvNRCuTtNvCxJBPd29vQc5OfQk1rm{+}n1/Uz5m6O2iJU8End06d2x2/lzWMmqtqlx9ovc7m/1UEfAUY47ccY{+}mPrXQR3MWhW8cXk774LuEanCwAgfMf8AaP8Ah9Kzs4{+}pumpeiOutb9fCGnJHDAI9TdT5KkZaIFR8x9SQKzbPRn1i{+}e61S8kuHxubbllU4yAT3bHp09B3zNHjn1PUJprpmk3ktJOxIbOei/ljFdrd{+}K9H8J2Ky6mUs7dFDtAiZYg92P8AQcmsX7isviZovf1eiHPo1rPAPN/0OwQENcEDZGo/idgMsfYcDjpivJPG3j6CeWew0CARW2djXYULNIenBx8ox{+}PoKp{+}PPitq/wARZmhg3aZ4fiYiGBBteUDoSAcA4xx7CuafR106xguL9njjlz5MKHEkvvnsuep784ranR5bSqbmFSspe7DYoym03NlJJmjGN2/5U9cnHJOc/j{+}VKSK3vFyWBVeRyc49O1SN5{+}ouIY4iFHKwxj5VHqfb61BeWDQxqjOCXYk4HHH{+}TXcuxyO{+}6J7Z97jafKReAAQM/SvR/A5TSQ0wmieZsEhXVmHPrj{+}teX2dvIsqmGIs2eDivTfCVlNuhRYZEkYfflU7cd8g/wCevNZVGkjalFt3I9bmW/uHlLiWXaWHdl9sdvwrk5547rcrbCQeA/b/APVXqXifwzFY2yfuTJLEg3vEhI6jBU9wQea8q1SD{+}0LmUFwJM5HQbvyrOlJSRpUg4swNXs5I58jeVb7u7ncP5GqVte3GnMrhQ0YOCAuM/wCfpW2kVwsTQuBcQ9dpPTtkDsfccVXl0zed3LA9Mjk{+}xxXUpJo5HBp3R0Wh{+}ItN1CFYrvdaSLnbMqjK/wDAhz{+}DAj0Fb1/4ImnsBNAsF9GfmSa3OCw6{+}vX615xNozW/76GTcmOUc/Mv{+}I9DXV{+}AvFdxoV{+}qxXAi38NFIu5H5/iB/mMH3rGUWtYG0JJ{+}7NDrWCXTLxeJLOUDOJE4Y45z3z{+}FdjpgTU7dL2Nle8Ti4RByB2k68g9D0wR1rrW0zQPiPF5bkaZqgUeS0BJRz9DyOfT864jUtA1jw1cMmpQOyRsUS6T7{+}MYwSODwercEd6yVRTWujNXBw9DvNOuEvC9tcI3mEfvLZ{+}OBjGD2IPQ9jjsRkv8ARRPIVMqsCgeKfH3mz7dDg4I/oa57RxJdQE2VxiaEBvKbBJHTABzke3XqPSuwsNSju7cw3AWMmQJvA4jkGcHnv1PuPxrF3TubxtLQ8j8QaLd{+}FJf7StAWs1f/AEi3UYMR/vD/AGD1q9pOpi01O3ngdRHKge3bPyvGeSjZ5BXsfY9sY9J1zQheWi3Fsoa/VWj8liBHOMElQGx1GSBnOQR3ryjVdM/sK1EY3tZb/MiQrseIscMuD6NjH1I7mtVaaM5JwZ{+}nv7LnxQk8WeF7eG5u1uLiJdpAwGAHQ4yfp6Zr6DRwyggg{+}/rX5GfBL4qah8M/EtlewTmS14aSJmOyaPPzKfoOfpzX6i/Dn4gaX4{+}0SO/0yYSRMivtzkqD2P45FVSl9l7nHWp8vvLqdf1xSH2pxGOnIo74roOUY3IAoxTgM9uKUgA0DI1JHHSlK96UqCeKMYxQFhpzjpTSvNSU0jn2oERlSTTSO9SEdfpTSOo7UAR4OfejnGMU8DnNNblsigCNl9utNPAxUjcnHT3ppHNADMZPpTHHOR1FS7SAcDmmFcUDI{+}2Mc0gJXOKkcdwKZjI96BiADFFJkDg8UUEmgq8dKWlAJPtSgY{+}lAC{+}2aOM80AZJpw6CgAPANAFLtz7U480FWGnI7flSgZzSgZNOA4oAQ9PWg9Bx{+}dLtyKXFACKnJpelL2{+}lGM0CEx9KMg5FFLjI4oATtQ5wpINOA46VS1C9W0jlkY/JCmTjrnt{+}NAjwP9sL4qJ4A{+}H81ukgF1dBkCj124A/DIP{+}RX5OazqNx4h1iUMzSSM5LEc45/z{+}Jr6J/a7{+}Ltz8QfHdzZROPsenF4FCnIkmB{+}dh68kDPoorxTwtpo0wJdS4E0mHiJHUDIyPx6HufpXPzWbkd0YaKJteGLaLwlG0jsiyw4ZpGAO1vTB4yP5gDsc5N7q//CR373c5KwrxBE5Lbz3Y{+}pPU/Sq2qXzalcm3XcYi37wjox9Pp2q9penLKzlhtjQDe/YgH7orK9vee5vb7KNOxjTT4F1CVDNcyDEEeBhRuxkjv/8AWqWzEt9MdrSNKRveQdFOQevfA/KotQujdFCGOIsIGUegwPpU8Nz9jtBLO5WNMFkX{+}X1P{+}NJNrXqytH6HSfbYtEtFSEtK6KWBC5APPOPXn6D615zqzz{+}LdXD3qulnE3y8ZcdSOO7Hnr0rfvZLmSeOaRAk0uG8sLkQRkZGR69D{+}IrV0HQjEH1C72QQRjKq43HnnJ9z19TkCrilT1e5Mr1NFsc/LYQ6XbxXD2qG4bi0tCpIAxne31/WsiDw9eazem5vpljThnneQbV77Rjr9BgfhXZLZG9E9/c3rWlmOATzPIPT0Qe3P9a4bxB4n3PsiYxwtwgbJZ{+}eDipU3J2juaeySV3sT6le6ZpVmba1ZpnYAskS7Qx9Wb/61ZVvEb/yFjtUUsDgbN7fePTP{+}FSaJ4e/taTK2iydSSWf169a9u{+}HfgaOyuI7i4s4IVSP5FfdknPufesalaNGL6s6KVCVaWqsjlvDnhBnMDXkCW0ZJyJI{+}uB2AHt6V3lpG9tsdYZU0{+}McM0GAxz0G/rz1Oe1ehWngmS9kSeaaS36kOUHTpwOv{+}RVxPCRuddhjtGlPzpAr3PJRc5JHoeK8aWK5nc96OFUdEeean4f1XxPHF5Ud7FaxjJlPl8tjpnA9OuPz7eSa98PrqNZJyrSQBvMLxjlARnI9ffFfdukeBIdT067t1LBZ8KHT7xQdTk9Mg4/GqHjb4ORQ2VwtrGArHCArkYAA/ofaqp4xxIng4z33PgAxNp8{+}25TzYj9yUH7x{+}v8Aj{+}nWtTTdJtr4vjLSMCD8mHQepXuPcV6j4u{+}GF3pO67Fj/oLtieFUBMZz99R6EYPHWuNuPA7Ws0bQMUVv3ilW4PoVPUcV3/WIz1TOJ4SUdJK5xmo{+}FLuwJnVRPbOCRIBnA75Hp7/n61zWpaHNCouo0zFnG7PzIce3avarGaRn23SBzxlZOrdRk46/X9TRqfg4m2e5tIDKkg{+}eDHUe3/1un1raGLcdJHJUwfMrxPKPDWr3NjNGBdZiU4wCSAfTjJH{+}cV7t4Z8XWfiTSo7PVJHuMkxghh5inHbuR7fQV4jq/hkWzS3VpFNFMv8ACvyt15DKeD{+}GM{+}tRaF4theRLa6TytpwtxGNrBuqn2wa7JRjWjzROGLlSdpI9U8Q{+}F38MznULNvMtS4ZZ4uOeD0z{+}nB4{+}la1vBLrFtJqEQErDas0SqAGXtx2I9faoLXxO95BGtwq3KygKSR{+}6uAQcBwejHnnv9aqafFLoc0V9pbPPbM2DFIfmjP8AccHuOx75rC72kb2W8TYsdWOqW8lncHy5ojtQk43Ln5Tnvzzn9a43xQxsrx0uo/tFhPlJ43XIXPBI/PPHsfWtvUr1kuIr1EQRTMG44aNuMjOMbSc9f6Vo3{+}nReINNunhaEswG2Juq9sc{+}vzDqeo9KV{+}R{+}oviWu55VpswtbybSZz9kuUfNtKcmNmAzj1AbB9eRX0Z{+}zR8XNY8AeK7GG3YTafdt5awPL{+}6yxGUDDgHptz6j3NfOnifRR5dvJHJ5cwO1Wb70cyfwk{+}42/XrUWh6vc28yS20wifIYeXk55z/iPwrWUb{+}9E5094S2P228K{+}IrbxTo1vf2rbo3HQjBB9DWsV5r5K/Yv{+}OK{+}KoDomoTp9t24BJwzN1yR3JwfxU56ivriumEuZXOCpBwk0N7Ubd31p4HOMUoUgetWZoh27ck0m3cKkbnr{+}tIRtPtQVcjbjtSbcdakI3AiowCaCRKaQKcaRgMcigRGVwP6U3BA{+}tS9RTWHFAEeOKaVNPoPPFAEbA5ppGaeR/kU0jHOfwoAjK854puCD1qRgDzTcbgc9KCkxuKKTBPUZooEaABH0pQCSaUdRSnGfegQYOelGM4pRyacBgUDsJjHTijB9aXg04CgYKMe9KBgUUo5oEJR0pSMUdCKBCDoKUdqOATS9hQO4ADFKB6UgpwwOaAEZgiEnNeBftZfFMfDb4d3YgcR31/mKLnqxByfwH8/avc9RuBDbu2Rk4VeO59u9fl3{+}3f8XD40{+}JsukWE26w0jNijB8h5gcyvxxwcL/wABNZzeljWmryuz58uLldY1RrmR5JFJKZHVnJycfiT{+}JHtV3VJPJcRsoE0h{+}crxsAGdqj0A5/Kq{+}nqllB5xjx5ahYUPQnB{+}Y/zP19qoTz3ErCeRmLvlY/U5PLfj/hXLu/I71ojS06JDOkUaY3nLyN/CPTA9{+}P8A9fFtj9sP2a1{+}QZLMw4yO5J9B29/pTLdVsLX92SJSPnkznYg6kn2/U0spj061WWTPny/MV6bEA{+}VPryPxNRd3ujRLRXFunt9Ot0AYKQMk7ulXvD9g19B9vu482kZAhtG/5bSH7uRnpnH6dsVi6DpMniTWPKlZEtYvnmLcDthfzPJ9M{+}lel6RppkYXM6ywQJkRqR07Zwe55x{+}HpW1uVXe5HxPTYqWWlzNcJbB40PzSTyZ{+}VR3OfbB//XWxLYLdwC4kkddItjuCn{+}Nv77e57D0x71fstHazt0nuV{+}RiRb2pXiQqDgn1Rc556nk9BnM8W6{+}2l6etqXS4kdv3cbP{+}7LH{+}JmHbP4n{+}XHObk{+}WJ2U6aj7zPN/GOsTahMqxRh{+}R5Nqik5HZmHpnp6kdhUHhb4ez6/qAmvJN8rnllw2T3xjv61uaPocurXf2iWR5VkHySOu0NkfMUTsOwLdu3GT718NPh8JUjZoSQcfe{+}8w6e3H5VjVxHsY8sTtoYb20uaRneBPhV5EUaWlsUZWyZiu7p2z3r0vR/hu8{+}rwy3SO6RKykAdTkHIJHH8{+}DXp3h/QGtY02xohUAcdh6V1VjoxkVX55JyPxrwJVZTd7n0EacYI5Cx8H2yKAIQqgYJcEkj6mtaDwLvRBbw{+}Rl95ZV2gnGBz{+}P45ru9O0/orREjoM{+}ldRZ6cuc7Dx0yOlQotkTqqJy2keHUsLe3SPBKf3gAMnrnP4Vp6n4dS5tDE3zYHOBx9a6qHTV3Kyqpbp09KtzWCNHtYDJ5rpjS0PPliLs8N8UfDy31eCRGjRSylc7eo9CK{+}aPHnwwn8IyPG8HmaS5yjlSfJYnrn{+}6fXt16V93ahpwJOBjjOVrkNe8J2{+}r2c0M8SOrDoRn6VKbgzrjUU1qfnxq3h8KFWaNnTdlWP3kb1B6g06xN9pdwuzN/C6BSH4PfHTv7{+}2CK9v8AiD8MB4VOxoy2muxMc2M{+}Tnorf7Poe35GvLL7Sm0yeSGX7n8LDtXQp8xbpqxRu/DcGpmW5sURJcAmKRPlOc5DA{+}/9K8o{+}IvwqKCXUbCzkVkOZoI{+}SV4yVPHQ{+}uPwr3XT8XEyOjIlywKuSMrKMd/U{+}1adzpTTBvlQGIbXCjLJx1AHUf0rSniJUZXTOGvho1Vqj5I8OalJpaMt1IskKNslU9BnhXwex9exxnrXpOk6o04LLlkcAOrfd9sjn25HQ89DUvxM{+}FDoJNZ0aNWOGM9vFyrryWKex7r2J9q4Twvqr2ToGkcxMAEdjhl9Vb3/z617kZxxEOaJ4DhOhLlZ6fLCLcSLcFJbKY{+}W8bgAj1B9COuenGe3Oe{+}nzeE9XNlta9spvmgdshivI/qR39agtPEFtbqY70GWBhscY5A5IYeo54PYjHIxW5dWxvrNbaOdJXVPOsLndww6lQfQjP0IrG7jpI1snqjkNb0yOOKW2WQyCdRJEzkkFwMr/ADKn/eHpXnNxJEpW7i3wNGxWdOu0k/MPzAP1Jr07U5IdY08ysxSUqQ4x/q5AecY6A/ex/vV574mt/s1zHdyriO4zDd7P4WHDN{+}ob8RXXRd1ZnFWWt0d38IfH83gTxdaahby4fzElhkGQGCsCf5YwfUV{+}w3gfxVaeNvC2na1YMHt7uISL7eoP0r8NdNuMu9hKyxzQnzIX{+}mQ2P54{+}tfoB{+}wL8dW3t4J1abAkO{+}0d2{+}UHncv1zj86E/Zzs9mY1F7SF10PutV5FOwO1KBn2pQuDmus4SMgHoKZtBqfHHSmMuKBEJTmoyvNTMCaYRg4oHuRMB{+}NNIwcU5v5UNg/WgBnXtTSMnHangYpGzmgREy44pBjHvUrAGmbfUY{+}lAEec/SkK5HNSEcU3GOnSgCIg0EcH0p7LwabyBzQA0ADrRQfpRQO5exgf1pcD1yaKUgYoCwADrS8ijHFKq56mgoRRT6OlL0oJbDFAzQDjrS4x0oEDUgGaX60DOKAF245xRjPNHagelAAOaacucDge9KRztzRLKttA8jsFVBkk9KBnjf7U3xbi{+}EPww1HUYpVXVbhTa2QPXzGHL4/2RzX5EW5l13VpbuUmWMMxLE7t3djn9Pxr6k/b/wDim/iXxhb{+}HIZCqWzHehPIPQ8dsc/jn2r5lVlsdLVCuM5wq9QAf6n{+}QrklK7udsIW0Y3USbh/K3BTIwQAnJAyTj{+}Z/P1qVGWK4ZI8tIoChj6{+}ntgcn61k6WztI1w{+}FMZKgHpn1/T/x0eta9ko{+}0QwBd0svIcjnGev1z{+}tQ1y6HSnc1II44bcvKu/YQ7h{+}jkfcTHpnk/gK5i6vZta1bZE5ZASQxPUDO5z{+}R/Wr/AIhvhHvtV4jiDb3GSS3A6foPoKoaIqNdRjaEUphwpwCB2/kPwqqUftsU5X91Hp/hrSI7fTEWU{+}TC37{+}WQdflHG714PT1PNd5aF7mSzt0jCM/z7JQSsY5Blf1AHQd2PoK4vw3Zi48tZ8yHPnyI3ORnCp7kkDPsM9q9W/sxNE0jN4yte3eJJnfk8dIxj{+}FQQT7kDjnHLXqcqsjqpQTZia5K9zeDbIsdsibVct/AMZ5PHfJOeSfpXl8NxFrPieWRHjunt03eaykW9v/AA5A79uT78ei{+}MvFU{+}r3b2sUotNNjIR3UndKQfuqP4jnjjjJPetLwjo/9pvBbWtr9ntY2Hm4OTJJ15PfHT0rD{+}HDmludcI{+}1mox2O5{+}HHhh9Q1DzCWmVOBLL1b39AK{+}mfCeipFHFvUjHoOwrifh/4aTT4YwQB34xzXsei2mCmAOfWvArVeeWh9NRpqnCxtaZYZ2gDg{+}vauq0/TFBCf0qppkI{+}XAx3zXU2ap8vy59xSppM5q07C22mrHjIAbtkVp28A29CCeeadFEDgHhT3NXoIwPY/rXbCB5c6l0NitBgD1GPxp6W4Jx3X0qyqlTgHHHYf8A1qaSVbIYn/gNdNkc25lXtodvc4PUDrWTdWwAJIBz3xjmujmOd{+}TnJznpWfNHlCdoPsKwnFdDeEmcR4h8OQatYTQyxhgwxgjNfMHxN{+}FsvhyQyRZOnOfkYjIiJ42n0HPB/Cvr{+}86n5SMdK4vxdpseoW00bxhw4wVZeDXE5cp69Kbe58MOtxomobJUG3djaRkEeldVpepQz4eGVY5AOPM4Jx2J9vWtj4i{+}C10{+}SRYwRGW4BBO32{+}leewq9vdhQ/lygbtjnhse/Q962upq5clZnTXNibmbz4jh3fDJ0VyB3A6N7/wAxxXmvxF{+}FhntLnVtLikuEibdc2apiVBn5mX1x179OODgel6Jdx3QVCgYYx5bD/DuMnB{+}vrVu6aW3mbyn35IVZJGzvJ/hb0PbPQ9{+}lXSqTpSvFnFXpRqKzPluxhMskmnSTpdXMMe{+}2lXjzEPVWHX1/XuK6vwJqPm2MmnSTGONX3R7yN0bA9j25A/n0Jrc{+}K/ws/tCy/t7w/G8GoWcjma0yAxxyQPQ9ePf8K8vtNXF2kd/GzRurCO4{+}UKVYcK2Mcccfh7V7sZRxELx/4Y8CUZUZ8r/4c9C8TWBsrldT8s2tu58u52DKA9N3t1B{+}hPpXBatZx29zdWs0fnQTJ91TlUcdDkdiOM/4V6p4c1KHVtJmtLlRcQyIY51K8{+}XyA{+}PUZ56dx3zXmusaHLbzXWnXO5bq0AeGUjl4{+}2fXGcVVKTTs9yKsdLrqec3vm2UysVYXVueGYcnb2x9MH867PwF42uPDviGw1OxmMMiOJI8How6r{+}v41zV/5sllKUZhdWLfvQWJLpnhvyPP0rM0{+}4MMrwhsrIQ8LHqrfw5/kfpXfOPPHTc4U{+}SVmfuX8D/ihZ/FrwDp2tWzfvmjCzxk8q4yD{+}oNeg7D/AI1{+}bn/BPj40/wDCO{+}JZfDGoTbLLUGLorH/VzY5AHoQuePQe1fpMo6f0pU5XRzVYcktNmRtgUx04qYqfSmsM{+}1amJWYYqNgBz2qy4x2qJlA4oAgZMn1FMOM4zipSKYy8ZoGMpCMjilz8tJyBQIaRk4zzSbeOTTxTTyTQAwgmmnr0qTPHFM780ANPJxzTCO1OYfNxSkZIxQBDjPr{+}dFObg0UAXR3p3WjHFABz34oLFH6U4dPajHOaXrQSwAzSYxTl6UEc0CFIzQenFFA5NABjJxR6igjvnFOA5NAxuTjHpSjtijaO1OA44HPvQA1Rl2/AVynxJ8UW/hjw1qeoTsFhsoGlZh1L/wAIH4/0rpp3a3BZnVMjOdvT3618s/tx{+}OY/DHwvfTVm2X{+}qyeXGpfBVBn5uMenP19qznLlia043lY/O3xXqh8ZePNX1u4n3hpDtLryx4zjn0/rXLa3eF7gwLxghMKeP89KvXjrYwNKuCiqdnPOff3/xrE0ovNNLduNywkY6nc56D{+}v5VhFX959DtdlojV08pEuJvltLdC0sndiecfXoKvQ3r2di{+}ozp5M1xwuOPLjA4A/z6VmQI1xdQaefubt0vQ5x1J{+}nP61T8RaodRutkRKxkiOJR2UHrRy8zsNNJXKtxqPnyqWTKK/HvgevfGQPrn0rpvA9k19fidyVtk{+}dwe6AgKo{+}rY49jXGRvv3JGflYhEz6ZwPzPNe0{+}BNHisbaCViGtY9rFc4DPg{+}WGPbu59M1rVkqcLImknOR6T4H0hUeW4vSimIG7ubh/uI38IOOygZx3A9a4r4nfEue7ikaGQ2kdwfLhGzEhjGeSeuSSWOO7AZrZ8QajcSQ2{+}mK7Q2UuLm4QHAkHBUHvg9eey9q8Z8Q6o3iHxPdTxLmKF9sSgcEg/ljPP1{+}leVTXtJ80tj037kbLqXfDVlPq2qxR4LyggMx52DuB74/XNfTXwx8IrYRI5jy45JP{+}NeW/CPwi8I8yYZkYhiSPXqa{+}ovDWjCOGEBFUYAPoK8/GVrtpHt4OjyRUnudV4bslVUyAO3vmu409MbeelY2mWarGu304rpLCHcwGMfSvFPUeiOi0wYQNyMnp2rpdOBKgE9{+}lc/p8fK9OK3bEfvABzg9q6aZ51bU3Lc4yOoHpWhbkuAMD6ZrOhBrRtcg/Nge/pXowZ5MiztDHbnlenFRyoRxjFWEIPfb25FMc9eRnHGe1b2IuUJTtJUgY9CKozkkHaQBjk1emDOenI9sVWkVQp56c1zTuaRMa6UnIyD71z{+}qW6ujbs4HPFdPcgKT8oz2FYWpIHJHBAPB7V580elSep498QPDy39tONgzj8a{+}cvFGlNaXh3pvZeq9M{+}49DX1/rdmLmORWHzEZIFeGePfDIM0rhMEDOQP6VMJcr1PRa5onkmkzpbzmdHKxnCukg6fl6V0l7bGOL7QoM0UqAMQcq49R2BH68VzF5byWt/tXCy9m6A{+}x/x/A1qaZfm3YDZiFwVntie3qp7HrXZ5nI/MqvdPayBpJhLFgAXJU7Ux/BL3Hs3OPUjpwfjz4dRia517SrIlpFKajp0WBlT/y0XqMd8jofY5HW{+}Kr1tCn80fvYJV4mjO3cnqfQj/HrWT4d8app99FZ38i/eCw3cQwkiHGEdB3x/d4PUDNddPmh70Tzqyi/dkebeGtYk0S7khQN5nG0PjLD0xj0x{+}B{+}tdX4g0{+}HxBpFvq2nIfNg/wBYrEk5zhlb/H3GfbovGXw{+}s9TZbqz22V5GytCRwkyEdOB94Z7dR0HauQ0nULnQdSZbyIxwSkpcRn{+}Ejgke3t6Gu7nVRc8d{+}xwOHI{+}SR5Z4jgFhrcV4U2Wsq{+}TIPUHIyf1H/AfeuOu4JbKaRQpzA5Ug9QOeDXuPjTwzHdyywxphbgFADxtk7Z9iRXj{+}rRMrRzujbwnlTKeSxXjP124/HNelRqcyPMqwcWdD4I8TXXh/WNO1qxl8u4tpkkzz1ByG{+}nSv2l{+}A/wATLX4rfDTSdbt2xKYlinj/ALsgGCPzB/Kvw2sZTpl2Yg48mb5onHTOOPwNfb//AATx{+}Nn9g{+}Mh4WvJfLsdWby9hPyxzj7pA7A4x{+}J9KbThO/Qya9pC3VH6XD86jZRzTweuO9DAHg1ucJXfnOahYcVYYcc1E6jg0AQd6aRUpT2qJhjFAEZ44700g5J/nUuKZjGaBjeaaaec5xTWHFAhMH60w49KkPQUzoTzxQAzmmnkelPb0FMPANADD164op23Pc0UAXQNvvTvxpBgGlHIoGw4{+}tH4Y96XHFHJNAWFXpQTg07OBTQc0CFpTx070AUoGTzxQMbtpwOOtLx70UAA4Jp4GB/OmjNK54x{+}H1oApXWZDz9xefXLdvy5r83/ANv7xhHr3xG0rQopSy2sWWx6uwH/ALL{+}vvX6Ma7OLWwmYcYQkMPXBH8zX5AfH7xavir42eLdVMvmwW0rRwuoOMINi8j6fpXNV10Oql1PIfE168dxLArgrDkluuewH86dCV07SoYyMMkW92zgmR84H4L/AD9qoRBJp42uQ2xm86UZwSB0UfoPxqfc2ralHG4AijO99vTJPIH8hWlkrI0WrbL9pKdO0xZOl3dfNhOqxj/Ej/Oa5ydiWMhYD5SoVPQcH{+}ePxNaGs3pmI2/KWACqPuov{+}HU/lWVdsIo9qAlioA7kAHp9SeT9KqmurFOVnY1PDWmnWdegg{+}7HGMu2MbfX9P1xXtKyLAltbbwttbKTIQAMkcuxHucL9Aa4z4caT/ZmnSXk2PtUgJPmDIBz/QAH8BWrJqitCyAbUYfaH3HJKg4RT{+}Iz9Aa4sQ{+}d2R24ePKrsb4i1{+}VbC/u5HdZrk5LFumcjH5cD3D{+}tYPgfSvtl8jOCF35Ix3/z{+}uaxvFGrSXd3bWmclSJCAeenGfpyfqTXpfwu00nY/TGKxmvZ0vU6aT9rV9D3f4e6LHbpGCi4wOnP5ivctB08Ki/L6celeZ{+}CrTaIwmMkD/61exaRbiOMZPK8mvmKmsj6mm9LG7YW2FUBc{+}tdDp9rgZwM4xms7TojjcQM9jmugtVVSuORnrWSi7lyloXrW3BbhcnpW9Zw7MYJH1rOsTlgwH4VtWq85PU11U4nmVZdy/GmUzjPfHpV{+}2VVBByOwPvUEChTu4GOlTySgdRjjk9q9CKsjzm9SUkKuRz7U2QkKMrweetV/NyAc8fXFKsxdfr0HaquhDpkLR54OOKoTLsJAH5VqxspGGOSPeqtwiqDzyex61E1oEGYN0jEFdvOcfSsS8gYghT19K6WdFk65YfpWdeQbV6Dj9K82cT0KcrHHX1qz7gOevHtXEeJdHFwjAISTznGcGvTLy3XJIHQc81z{+}qxxMCvR8YAzWDielCdz5b8beGdk8sgXHPp0NcpCqTt9nmkENwDiOZiNrdtrH8sH/I9/8YaVBLDKoHznPLelfPvjNP7PuJCFKjJ7c12UrtWMqndGZ4gmkhhmsrhAksZKgSchTjkH1B/{+}uK8Z1i6bTYpFt0W6tg3{+}okHMZz8yA/qD9etep/21BrGn/YNRl8iTbsgvmGQnor46p78kfpXinjye/wDD2s3AmDQTxBRMo5DofuuMcMMHqOoxXs4aF3Y8PFzsrnsvw88e2V1ZC2vLnzrKSQKryjMsTntIOTg56/jWz4q8JfaZbeGeSNRKuy2vtwAfP3VfsT6H8Oh4{+}abbxFc6JfC/twsto6jfGTnC91LDkjoQR0H6/RXw58f6d4j0lNOuiJbKX5GSUhmQkZ4Pcdx69RzkUVaMqEvaQ2MaVWNePJLRnP6ppt3PaSW08Zt9YszskV1/1i44Pv0/T1ryzx3poJkvYo9pmkJljByPM9c9sjOD6nHavoLXLA6XdQWVzcvcLtBsNRcjdKnaJz6jsx9gexrzrx1pELKbqB0WKdtssSrkRuBggr6Hr7EVrRmk01szOtC8XfdHhBj8{+}F7ZfnaP54gfvAdcfnXSeEPEV94b1DTtZ06aSK8tpQ5ZTkq6nIOPfFYGoQ/Y710bK{+}UxKleSq9156gH9Kl0{+}4VLuRSA6SDehJ{+}Vu{+}PavVkuaOh5a92Wp{+}53wQ{+}Jlp8W/hjoXia2kV2u4F89V/gmHDr{+}efwIruzyK/N//AIJsfFx9D8Yal4Hvrgta6oguLJd{+}VEoXJGOxKn/x0/3a/SHrRF3RzVFaQxhnrUDjFWTzUUg61ZkVjk54IpjjIqVziozntzQBF6UEDmnNnpTSOOtAxjKM5/lSdqfjBprcnFAiPGTSMB0708jA4OaafWgBoXrTSBTj9aRuhOPxFADePeijBPQ0UAXdmT1pyrkA0DkU4crQAhXikA45OB604dKOhoGJtBAo28{+}1OpV4GaBCgYo28/Wj9KUjFAARim45zTqCMUDFAAFNKFjkt8vbHFPAwRS7cnHSgex518dNf/4Rr4c61qX3RBay49MhcgfiRivxl1e9ZLXUZrjc0kr7nOODznHvk/54r9X/ANszUJk{+}GaaTbz{+}VJfySeZwP9TGhkfr7AfnX5DeKrh7nUYrASZO7c2Octk/y5/Oud{+}9UsdULqnfuUELR6Y9zKczXDfL7KPu/hnn8BVyCBreykU9SvmynuMdB{+}oH1NQn/AEu8SH/l3txwfp/jVu/mSJUzt2AhtpPDYPGfx5P4Cm9XYqOiuYt2s2/c3EsnzFfQHGB{+}PH4D3p{+}k2bapqUNpCC2Xy7E4JUfyzj9arSztO0sxOR97J7nt{+}pJ/Cup8C6U0kMk7Hy2u2Mfmdo4EGZH{+}nb8K2k{+}WNyIrmlY71bqJLGC0TbGHjMrsvQR8KPzGSPYisi9dPLnaUFIYV{+}0TYHOMDav/AHzj/vqrFuol3zMNyyP58qkcLCmFjj/4EcLXM{+}NdTdbH7PkCS8ZpHI6lQe/1Yn8q86KcpWR6UnywuctpkrXusee5JaR9x74Hp{+}HT8K{+}h/ht5cEK546Ee9eBeF7IyXanA4PevZNK1tND00zsocxrjFPFrmSSKwTcPeZ9J{+}HPE1jo9m1zcTxQxR9Xc4x/9ar7/ALQehxOsNrcC7JydyMME{+}vp17V8z{+}DvB3ir4wXbAyy2ml5wXK5XHoBnmvb9I/ZXgt7ZAuqToMcqUAO73rwqtKlTdpPU{+}gpVJyV0tDv8ATv2krMZXyTtXjeTj{+}fb3rc0n9pvSpJf37BFH90jGPrXiOpfs66rayPHDrSyoW3bZY{+}3PA5PFcnrfwm1zRy8kcEUg2/ftm4/75NZxpUpbSNXVkt4n21pfx78NXKDGoRI56DeM/wA66vTfi7ok08Fs17Eks4zEAw5HtX5gatBq1kwRrZuBydmCPetHw/491LSVtUeV0KS7lMhyY/Xb7Edq1eGlFc0ZGMalObtONj9XrTxhZzFSlwrB{+}hBrZTUFdMqwI68dMV{+}eHh/4n397aRNHdSQ3FpKzkhuo2kqPzJ/Ovp/4X/EObW1YSS74oykfB6sUyawU5rSQTw8HrBnuougFbIx6YphnUcFvmPYVjwX4kAII2kdzTzccZyTWnOcfIai3hVD0z2I7VXvNRWHl2JwKyZ9SjhjYkkKBXi/xQ{+}LUml2GptbTqpi2pGP9rqx/AHNZzm7aG1OjzO56vf8AjfT9OSSWWZdqIWyWx0OOteaeIv2kfDumoUWUyztnaic8Y6mvlvxp8TNQ1XyUxJJCA5EeTh8scZ9hmvO7yHVrx/NKvI55d1B4P1ohTcviZ1uMKeyufVmsftS6aZvKjGwEbhLJwP5/zrkdX/aVimjLJKjY4JA6{+}uOa8EtPCerajKzPGuc5LMMD0x2rprD4Oz6haq897hepQPgY9MVo6NJatihOo9FE0PEf7RVxPGxjYnGQWx09s9D{+}IryPxD8Zbu6ldLyEupyQ6YGfqK9zsPgro0UBF3tkcKPufMSfxzWH4q{+}EHhZIJJEsgz4yWHBJ61tSqUYu3KZ1oVZK6keMReJbfVLRngkBweUJ5XHNcneeMbZ5F0rXkabTAxEF1Goaez3ddoP34z3jPHUgg9eg1TwkNAvLh7ffHbsc7G7c15f4qjLuT0y3AFe7h4QbvE{+}cxU5qOu5sXmh3mgXFvYOYpbe5UyabeRnfDcA9EDHqrcjB5B64OcaPhXxB/YNxE8B/cE{+}VPayEhocuflbIHGeh7Z9RWR4P8app2iz6LrMDan4blmBMAOJbZyP9dAx{+}64HUdGwAexG94j8MyALr{+}m3C6pazZzeRqdl4h4YOp5SUDlkPXkjPU9M0vhkcEJfaifQXg3xfp/irS49J1QsbKVfLjncnzLWbnhiO3r9c8jNZ2veDJNDvLnTLwq77Pklydl1FxjH{+}2vp6YxXinhzxNJ4d1E3lu73FnJhL21J5UDOGz7ckN1BB9efofw94g03xtpSabqFyJHKBrS8I547exxwR0IHHSvGq03RldbHs0pqsrPc{+}Z/HGkyWV65cYlgwrYGdyY4auQdntrcGNmSSF9yt7ckD{+}f5ivor4i{+}DZroTwXEZXULVSY5ccXMWeeR1I4P5nvXz7qeny2zMjqSCMBlPyt6EH8sj1Ferh6inFHm16bg2d78M/F194O1/SvFujyGO{+}0m6iuPLBIG0H5gfYjI/Gv24{+}H/jC08feDNI8QWTrJb39skwKnjJHIr8GfAGopb3yJL{+}9hBKvHjllPDAD6E1{+}m/wDwTs{+}IFx/Z/in4f6ldGSTRrnzrISHloWJBI9uFP/A6v4ZWOaouaPMuh9oA802ReOtPApGGe9anKVXT3/SoHBz6VbcYqF0GKAICOlN2g1Ifu4pjDbQAxh/k0hGehB{+}lPPrSDjNAEW3H9aQjAzUhGabigCNlzg03aR1wam6A96YTjjigCMjBop3HrRQBcC96djNIBxQBtNADsbRTR0p45FMHzGgBwHrS0uOBigryTQAEYpRnJpBkmgDPfGKBgMk9qcBnpR0pVHegLCjrUgH40wLT5WWCB5G6KpY0CPjr9t3xMtqlxIJSkdpZTwAZG0yPER{+}vmKf{+}Ae1fllDMTLeahJ8zFisYJzn3H4CvuH9v/wAQtLp0UUblrjUbpiCp4x2B99pXp0r4oCBruCyjVTHApyf7x68/p{+}lc0d5M7baJFi0AstPd3UAv8xOORx/nFYup3b3ThAv71xyqn7o6Bava1fZ2oW3KhGAP4j9P1/KqtgRb3AZ0zKTvO4cgjOB69iTW1NfafUJv7KFNksk0NnCOXbGD2xx/Pcfyr0rS9Mit7G4CtiIFLKMDoFX5pG9{+}ef8AgVcr4U08u82oSKGVPliz/EBx/wCPHA/4FXbTY03S47edVRo/lZT/ABsSGb/2VfoDXNiJ9EdVCF9WU47nLOrq0ayMJ3B6qACI1/4Cpz9WPpXnuu6gl7qN3IOVj2pEvoqtXbeI5DZaPeXDkm425ye7E9fzx{+}FedW6mS9ZMYG3kfgP60YZXvJixTatBHc{+}ENN82WNgBufjnPWvT7Xw0lykVvKNyD7w65Nc74K0tsW{+}AACufxNejW9lPFMGK5HHzY4FcWIqe9Y78NT91NnqPgrVrfw/osVvboqR7QRhsEGte7{+}Kj2seEZiADkjtzXm7LILZWYgNymOueK878e{+}I20iORQVkcpho2JBQ{+}p{+}leUqLqzsj23VjShzSPSfEvxzlsV2y3JAflAOp9sdc1iDxr4z8SRCWw0id4Tt2TTkRqwPAxvI6{+}1fNk3xGm026eaAC8u85W4mXlfYDpipLX4ua7BZG2k1No4925QTll{+}YHjuMHtXrLL5RWh5DzKnd3/AAPoi80vxpFF9pu/D01xEVGWgZJ8f8BUkkY9qTRIrDWFIktVWVThkI2sp9COoridI{+}OfjD4Ua7p8HiK1YfabCG6jhuwYzNDKm6OQZ6qy7SDXsPhnxVoPxqRbi1KaT4kjDHzGUKJFB4Vv7wx3rmrUalL4lodtDEUqy9x69mVbTw5bWUvmwoyKVwyq2MgHP9T{+}det/CjUV0y4GZMq0hkKg4{+}YjGf51gab4flv4DBLBsvI{+}CvY/T1FXNN0u40HUI90bBCeozmvLqao9KLWx9S6HqZmhQg7gRn1rZeYgEgn8RXC{+}CLrfbRAcZXr/APWr0Ge3H2ZWzg1jGTaMKqUZI5rX9RMFrJn7pHPavlr4izR6vqFzbhS4dzwnf8a98{+}IN{+}1raS7AT2ANeOaToE13dXF1LEGAJbcaUHd3Z0q0UeZa1YW{+}jWH2mdVhSNQFwvPsBXIRad4p1adHL23h3TirFJdTcl2UY5ES8jr3xXoHj3VbXw6v9t6ttKRNjTrAAl55Ox29/b8{+}wr50{+}OzfELSTZa1rljNpVtrCSTWqSMcqmV7dj04PavawtF1dTzMViVS3PUdRhm0fRby4ufHVjttyf3aWhDSEEjCjzM/wmuR0n45XtjYxI91Z3jEbUjDMsm7HU7hjA6de1eAaz4itNT8P6Ja2unXFrrNuZ/wC0b97wyR3hZy0ZWIr{+}7KqSp5O7rxXffDP4D6t8TPBWta/HM9uLCUQo3l5SQ7dxB9McfnXr/UqcV754izGrJ3gj2rSvjnFdtHFu8uZhgh/lJGTnHPPGOlbi{+}O4r8SLIxdgueTxmvi5b240W{+}ktmciW3crkMSFIPavXvh/rVzr0UcKK67Ttkkx1H8h/n61yV8DGC5obHXh8xlUfJNHX{+}MZBdRyGMncema8U8YWjLCNqgDd0x0r6ZufBEyaHNNcRqpI{+}Uda8J8baa0dkXQAlZME{+}h9qeDmk7GONhdXPLEcgqueA3{+}NdP8O/HNx4Q1KeJ4vt2lXYAubCRyqy46EEcq47MOQcVy7r{+}/dcYODgVJEuJ0ZRgMMg{+}{+}a9qSUotM8CLcZXR7J4i8JW0{+}mReKvDFw1/pU0ebiNT{+}9tHHDLIo6MAoPHB25HBIWHwN4nFni18zyFc7kCHgEc/L7jqB/jXIeA/HOo{+}C9du/sZR4J2AmtJhmGZc8q49CO4wRwQeK7DWdB0/ULCXXvDbMmnlwJ9Omb99p82fuNgcocHaw7Z6HiuKcbLllquh6NKet1ue9ad4gtfH{+}mQ2N4Fh1ZP9XKvG/0Knse/I788E48E{+}K/hC50DVJZola2DEjZj93v7gg9uvB6fiKuaJrjxpHKHdZY8bwpwR6MPQ9Qfp6EV69Z6npXxN8P/wBnaskct8YfKW4bP7wYxz3B/P8AGvOinh53Wx3ztWj5nyZYzompkmNrV26mE4AI56H/ABFfW37O/jSfwP8AFT4e{+}OVmH9l6oy6JqUrvjcx{+}Ukg8/daFs{+}qmvnnxz8NNU8EakIp4z5II8i5YZRx/CpPTPbNdf8ODLrvgrXtEEbNcWUy3tsyZDK6o{+}OnTK5H1A9K9OTUkpxPMUXG8Wft3BKJUByDkcEd6krzD9nP4hL8Tvg34Z14tm7e3W3u17pOg2uD{+}INeng5UGtU7q5wtWdiJ1zULrirJHNQyjOfSmIrOM800rtGKldeKiZc8mgBgxikJHPFLjHSkzz7UAIeOlNJ6045NJgYwKAGbhTR605hg0YyfSgBpGaKBRQBcAOef0pVGG9aXGTShe2aBiGkwO1OCYGRSjp70DsGMClHSjpzQBk0AHSk6dqdjn3oPFAhtPB6UADNO2jOfagdxyjJFUvEF19k0a8l6eXE8n/fIzWhGAa5j4nX0WmeB9ZnmdooxbOpdDhhuwuR{+}dJ6IErux{+}Vf7aXib{+}0PHlpaAh49Pt1IiJOPMdVA6/7Cxn6mvnSCM6fFI0jhZGy7leu3/6/QfhXonxo11/EvxK17UJ2BSOZs8lgNvygA{+}wGK8u1a5NzMIgQjMfMkJ7emfoP61zwXMrHe9DPuJ2lPm7CshJ8tf7vv8Ah/OrlhZPNN5MeQ4UDJ7Z6n8ADTbKEbHvXxsU7Ioz39P15rY8P2sjeWQMT3B3l84wucL{+}Z/RT610SkkjKKuzstBs1jt7eFFBLuqqoHPBwPbGTn6LTb26OpeIBCMyRROzMynO585c59ug{+}lWhcGxsrm7UbjBAVi9CWGxePYEt{+}NZ{+}nW40qyEsnySMfvN2z0/Xn8DXlP3merH3YmH4yvNzG2U5jjIDY5Hf/AAz9GrnNN/dnzGXLN8oHrzz/AEq9qEhurbzOS0sjSbcc44Cj8sflVO1njaSOLhjGSfzr06UeWFkeVVblO59B{+}ArNLhLUrwVUD3H5V63a{+}HhJDuYYLA5wK8t{+}F6gQRDIII3Zz/KvoPw5brcQRcfK2DkjIr5vEycZM{+}qwqUoI8y8T6XqmmaaDDGGfaSGYZB9OPX/Oe1fK3xM1vVL3UpFuwdn9xVOM1{+}mNt4Gg1jT3jOF3Z5A5NeV63{+}ybY/bJr1R9scncFmbPPsOmK1wmJhTd5rUwxmGqVlaMrHx5{+}z78MYPiF44s4NXYWmlKd8rE/MeeByeM88039oL4Mz/C/4g6nZwqZdImkM9jdp8yyRtyBn1HQivoi9{+}BENtqEkpiaxljPyxxjaMjr09au6d4d13Sba5M2ofb9ibY4byISDcxGMA9B2r0Xjru6OOOWw5bPc{+}VtG{+}FmpeKNc8O2VpctqsmpxRhGSOULB820ozOoHy45Iyo9a{+}8fi18BPhto3hqx/srUzpfiqytUijvtOlOJZFQKN6jIOeeeDzyavaRreqmyfTmjsEit0jMMlvZqjAk/OFx0BwDgY613{+}jaZea2TbWcIED55EYOM47446frXNWxspbI66OAp0le55f8AAXU9W1LU18P{+}JTFevGFEWsWYLRsuBjfkAg9unb2r2jxd4Mh02VogQwRRhwOCCM10{+}i{+}GLjTY1Q7YxxnYAPxJrP8AElv5aspcsRwSSTxXj1He8rWuepFpyST0Kvw{+}R4jEnUg4r2PyC2m9BkjvXk3gtBHcKvbPFexI/wDoW3GBjuKwpJNNkYvRxseL/EK0Z4WUsFG7qKr{+}DLC1GgXEcsAeRiQqg/fOelbnjVFeTGOlV/BjiCXY4G0t37VNO17HRO/KpHimqfBDxNdfElvFM0VhfSwbXsYbpi0dqMHICA4yT368CoPi1oV78W/DsPh3xn9mtoYJs2l5p9m/mK{+}DnbkkbTwDn8uK{+}rrvwpBd2oMJKSEfwniuE1v4UanJIJLO8STBzsfIAz2r1VOpBKzOGM6NR{+}{+}kfC{+}r/s5eDl0TTdLbxVcxQ21zKzQ/Y0MpmaRY8FxGGCnCEBieORjNa1l8Tb3wR4KTwL4R0WaHTBuZ7kplp2P3mY45J4H4V9H6/wDCnWt1wr6RFcM0iyblA5KkEHP4CsCb4QeIb1P3enxWqhyQTx15P86uWKk9JBHD046xSPkKb4C3firUFm1OGG3u5cN{+}4QIGVu{+}ABgjNerfDX9nvVPDN4kVxAjxxH5JlQDcOx{+}vFfT3gr4JxaUI7zUXWe6RcbQuQv0zXZX9jHaxbI41AXoQKwrY2o48t9Ap4WmqnP1Pn7xV4aaz0aZGbLBc7VH9a{+}QPiJaoINQiQEFZAy{+}5zX3j8RLcSafNxksp5xjtXwd8WZGsptQyTjrg96vL5OUzLMIcsDxS8/dXc0p5xweO9RbCohA{+}7nH4Hmn3z7tw5wTuPqeMUkZZ7Xy85aMh{+}B2P/AOqvrNj461mSW7O2pRiRmIZghwexGK6Xw/4yutM/fROokZQkiOMrKO6sO4OK567gaO4icDAYA/Qgkf0FRthJpF4TLMP54qJRU1qaqTi9D27Q9LsvENk2qeHZR9pQZutKmGZIfof44z0z1BAz6nFg1WXSLxJ7csqk8xk/MPUE{+}ua4rQtZntLuK7tHkguYujqcEjH{+}f0rp5PFtl4pUxaggstTQbWuoVxFMo4yyDo3TJXj25rilTab6o7oVE12Z7t4d8TaZ450BtP1tUuoFwFyPnx3Iz355B4P1wayNI{+}F7{+}FfFEj6dOb/Rr{+}IwLKjlXH8SxuOueCozk/NzkYNeTWV1c6PPHuPyOMrLG{+}6OQeoI4Nd/4b{+}JUun3cYmV5Ivly6H515GCR0ZfryPWuVc1N{+}7sbu01rufZH/BPTxulveeOPAlxP{+}8guRqNvGScgN8kuPYOoP419qo24Z4GRz7GvzS{+}COp2ngX41aT4qV/s4nzb3qqcK8cnBY59CQ34c{+}tfoxousW2tW3nWkyTxN8ysjZH{+}c5rupTUkeXXpuEr9GazCo5ORT92QT2xmmtjFbnMV26VExxUz8VGQCaAItvY{+}tNPpUjcGmnkUAM74zSYxTzxSMuaCugx{+}lNx0qQjAHFMIxQIjKk0U8UUAXOQad060Y3fhS47nrQA09KXGcUqrilNA7iEUDrSkZFAGTQICM0pHHTNGDnpQM4OKBAFxT1GaaoJ5JxUqDB{+}tAD0Tj0rxH9rnxoPBHwl1TUvMjX7Oo2RuM{+}bMzqI0x9ckj0r22SUQxM56KM1{+}dX/BR7x9c3es{+}HvDCTlbeCJ9QuVUnDOxKqD2OArY9mFZzdomtJXkfDmtXOBI7vvd28yRs53HsPfkk1zEcbXTkA5knbbnP8I5Zv5flV7WLtrq4MSDdjksD/n6fhVcH7NCFH/Hw3y4AxtX/E5pwXKjplq9C1HAl9qsFlbuVtVwC/8AdUfebP5muy0JISZJpEIVF7D7oIwB9dvH1NcroNs0FrJIB89yfLHHOwHnntkgCurs3/s6yDFQ7oPNck/ekb7i{+}/qfxrCq9LI3prW7LdqH1q8mtpBtgilDTkdAcEsPwAQfgaxvGmpedc{+}UiCN5JAEwOkYAwfyH6muk0{+}EaN4Sdjhbi6c5kfqAcFj{+}QP51xkLNqutC4fO2NS3TPc/0Fc9JXk5dEdFRtR5erKes/6K2xcnCCMgfQk/qSPwrB0{+}I/2jcnH3QVA9MHj{+}VdDLFLeMWADsTuIA5HI/mafBpzWd5cWdzbSQSXSCWF3UjcccgV6CkkrHn8jk7nrXwq1MS29uW4{+}RRwTX094Pl32cXqBzxXxr8LNQNtLHHwArFSv4//AF6{+}vPh5dq9vGzHtng/0r53HxtJ2PpcBLmikz3DwvP5UCq344ru7TSkvVQMoYNxg1514cuUOD1B9q9K0G72bOjY7GvGhLXU9mrD3bopar8OLS/P7yBWHXLCsNvgjpEzgtHIpLZwGOK9at51lXgZz6rTwcnlO3QivQSR5bqS2PPdK{+}E{+}iae6EWXm46eaxau2ttMhs4QkUSKFHG0cYq8WAPIA47VUu7ohDjA/Gm5qKM/em9TP1LbFCVx8x/KvOdfuPMkKjr7Guy1e5PkOK8/1IFmZs4brXnVKrbPUoUrK5q{+}EWLagpGQcjr6V66rZtiD6da8j8ILtvUbtmvVN5EJOcZH4VdHRMyxavJHnXi4hpSpOeT3rO0WRVZVONx9a1fFg/f5xuHNY{+}lMu/PcGsW7SOuKvBI9H0TUf3SwyAHb0Nb8YEwX5u9cPpkp3r0Jrr7CQbBxwe1dtOr0Z5Veiou6Jzp4lZhwT05702XR1UcoOnOec1dUgjj9R0oduDjgY6dq69DjvI5y/gjhjbC4PqK4zXMKGIOSOK7vVDtVgSB7CuB8RH5HPTArzKzPXwybPJfH16rWsynBwD{+}VfCPx4Gye7kLbg5HQdDmvtPx5diKO4JOMA18h{+}PNGHibxFFZld4dtxz0wP8iu/LXyzuzLM4c0LI8JudIm/s03QHyIwGB15FZ8an7Sq5IDDyiPfHFeqah4McM0Cq/wBmI3gDpuAPHv1JrzO5jEN7OD821xyOMEHivq4VFO9j5CvRcLM07u2MiWbEkZGTn9f1FZE{+}8yuc4bllz{+}orsJ7VJtLPy4IOQe21uf55rCvLQRgFVCqATuJ/A/y/WnGRnOJBbOYg7HKrImcemTj9MilnMtu5ZsxyZOG/ut/gahlBhgXGFcAqQepCn/ACr0UkV6j78kqu18enQH8OlN9yYvobXhzX5ba3YDY8IbL20g3Kjdx7exrpLObTtUEnlTGwlI4jkBKKcY69QPrke4rzdJJNOu84O5PlkA/iT1{+}orVFyWmDqSFIDLs4Izzx/ntiuadO{+}qOqE9LM{+}qLG21WXwxoWsXkqW32J7eKSUEMJImAUg9iBtGPr2r7Y8I{+}D9Y0DQ7XWPC2peVNPCjz2n3oQ{+}AWBTkgZzyv5V8IfBTxpDrnwt8X{+}H7x0kuba3S7tw6hmkwy5IB46HrX6L/ALUU8S/DLSb0skhyV86PgqTgnp23FhWMI80rPcVV8sdNjU8P/F5Y1W08SWk2kXI{+}X7Uyb7WQ9Mq4PA9j0rvbLUor23WWGaKaMj78TAg1hat4ZTUfMWUZcjByudw9GHRvyzXAHwpq3hOeaTQ7y5snI3/AGdF8y2bPcow{+}X3/AJ11Lmj5nE1GWx7AJlkfA59{+}tK3B4/GvJbT423ehv5PifRZFVeTfaYnmKQOpMWS/HtkV3/hrxpofjO0NzomqW2oxg4YQSAsh9GXqp9iKpSTIcWtzYYd/0NNxxShge/FB9qokacAdKbj3okcKORRkNgj0zQMD0ph46085prA0DQ00UHAooAvKDgnPNKeB70oGBjtRj05oEGDRg96cBQTQA0gnpSgYoHt0p2CKBCEY75oAJpccetOGFwKADbgDJp{+}KTGe9A{+}c47Dqf6UAZevXht7E7CBI5Cx56fU{+}w6n2Ffjd{+}1x47/wCEr{+}L3ie8hkb7Olz9mhDHHyRgRrwenCZPua/WD4x{+}Jj4Z8G{+}I9Y{+}UJpulyTLvOAZHBVOfoCce4r8OfF{+}pvq2p3kzsSXm3c8nnP/wBYVlL3pJHVS92LZT01YvmkkywHLE/y/H/Glt4jcy5kJXgzSPn7q/54/GqZcrGLdW{+}VW3N/tNV{+}5L28ENtGMXE{+}Cxzx6gfQcmtHozRbaG3pTGZycHyeU{+}XsMc4{+}gIX65rdjt/t1/HZgBlQiWQHpu6Bf5D8/eq2mw2{+}m2ALHCxjHI{+}8{+}OB/NquW7/wBj209xI2ZD8zf7319QD{+}Z9q86b5nZHbTXLFXGeMNT84tZRHGwqi4PYE5P4k5{+}gFHhOxH791Ut8oB9yTt/lmuYtZpr{+}9uro/OpAX8CP/r/pXqXh7RxYaHdXIxkSRjPXACEkfma0a9nFRRMf3knI5zwZopv9dsbEqCZJMkgds9P517X8U/hTbatolk9sTBqVvtkRQOmMHP0rjvgtpQvvH2nNIAcqPlX3wP5mvqbxR4Ul1ad5bWGWf7NanzDEhYKgHJbA4H1rzsRWcaia6Hq4WknS95bnwro9tJo2tXUTIVKzbwPqP6V9MfC/XBJaw5wDjpmvGPH{+}jtp{+}vi52lDKuGwONw/8A1103w41v7PJFG7gZJzRXaq0uYeGXsazgz668OXu2NCG7jvXqGgXQPl7mOK8J8G6qJ1jweAAf0r2PQLoEJz2r5qV4s{+}nS5o2PTrCfai47jGRWoi5yd2SfauW0y6wnBzj0610VpP5ihhx7V2053R4tWnytkk6jZ2Hqc1mXQBQgt9e1bL7ZUPI4PT1rLv7lIYn3EdM1cloRTvdHJ645gjYnhQOoNef6lfqznZjB44rX8aa{+}J7y3tI2ILtg47ism6topbiKNFztTLZ9a4XHU9iOkTrPBimQBgv516QsZe1wy544xXC{+}CIGTZwCenXrXqi6cTp3mZB4xgV2UYOSdjzsVPlkrnlniOPG7vweorjbXUPJu2jxjHWvRPEtsY5XBOfSvOL63jtdTRzht/H0rmnGzsd9N3jdHd{+}H5DMqlW3CuzsIyq4A6DnPGa8j0HXXstVFtkBW{+}ZR9K9S07VhOihm564FaU0upyYiLexvwNuPJ7enaklOMAH/wCvUazqIwUOB61XurgGIkHg98V3XsrHnKN2ZGpzHaSSRnvXCa9LiJskdCM5rsNUk3KSeQK4HxRMEgkweQCa82o9T3MPCyPC/iLcBVm9Rn6Gvl3WrmR/FTpASknKjafpX0J8TNWWIzHdtAJJJNfLtxqlqni61ur3zDaR3kbSLHnLJu{+}YDBBOQCOD37V6mAi7NnBmU17q8z2HXdAtdI{+}Ff9p3AjiESLKD6YBDfmCR9a{+}NRci6ubqdvuSOTjPPXIr0X4xfGS48QQXfhrToprHSYr2YmGUncE8xikfJOAowDyfrXmNhGJBICTg4UH0Pb{+}VfRYSi6cXKfU{+}Wx2JjWnGENl{+}Z6Ho5e90iJic5j2kAf3Sf8TVWezkHmYX5VOB754/{+}J/Or3gVy6ywHkQtv6c4xk/oWP4V0d14fNvNIuB/Eqg8g91H5hRVSfLJoyiueNzzK4sT9qEIyWZWwW6k7Tn9Mfkaz9NvHtpzOOUYlv97nDD9f1ruNR00W2qW0/wDyzDA9MY55H5Y/P2rjbK1DXN9bY5hYyL64HDAfhg/8Broi{+}ZHNKPLJFy{+}tj5aTxMWKD5Wb7zKPX3HSorWdEkVM/u2GVwfufT{+}f/wCqrUDhZjC5zG3Ksf4WzjP45AP4elQXFqLWRonQhJOVbvGfb{+}R9vpUrTQ0sdD4c1m40XVY7i0YJNDkMoOAyH7yn2IJ4/DtX6p/sG{+}K7PW/hVPaW8ikW9wAYgeY89m9/8K/JCBpHjBHEq/Jwev8Aj/8AXr6d/Y7{+}PM3wm8YwXl4z/wDCPai6W2pfN8sQJ{+}WQj2PX2z6CsWuWakU7zi4n62uD85A{+}ZAOvfrTJEinKEorqQfvDkd6i0{+}9jukjmikWWGZA8cikEOpGQQe/HenyjDoo67yD9Of8AGtzg2Me/8I6fqbFniKknIHBU/gQetcZq/wADtHuLxryyggs7zqLm2VoJwe2JFNeodDTSO/Sk0nuUpNHlmlab8QfChZUu7fxJagfLFfERyr7CRRzx3KmtWL4mNZgrrPh3VdKcZywi{+}0Rn3DJkkfgK7pRktnnnv9BSlAynPIPrQlbYfNfdHO2HjfRtSjVobh8MOFlt5IyfwZRWhZv5oLAEqSdoPOBnNXFtYlfIXB9uKkKimLQZjIpMc89qkwMUxhnpQIZtU9s0U4LRQUXMehpV{+}lL19qVQe9BAmCTQVFKcge9KVNACAcUEZ4pQPwpdvORQAijHUU7PrS4prkqOBmgA6nHJ96VuAqD{+}I4pyLtHr7{+}tRXUqwQXEzEDy0JyexxQB8i/8ABQH4hJoHwalsEbZLrt44IHXyocqB{+}JX{+}dfk0JjOtzIRlgQ2R{+}X8zn8a{+}3/8Agpd42km8T{+}G/DaHEdpp/mMAOrSZ/{+}Jz/AMDNfEFjEwJIGBgjd1yTUwV7yOpaJIktbdYoHnlIdA21Qf4mra0Czkvbx710JYnZH/U1nQWMuoXPlxoViViAFHP0H17mu5gtmsLOGHKpIy5B6bEH3m/PP61hVnbRHRTg27vZDoo96EKAViJVc/xMTkt{+}g/KuT8Ta4bu8isYSwjjyXx/E3{+}c/ma6XxDqK6HoysjbLqRMxR45TPQn3PX24rz3RVa51ZNxJJDEnqelKjDRzZVaeqgjvvDuisdOglKnY7gZ/Aj{+}q17BDo/2Lw/LFMjAO0koA4B2xjg/jXK{+}E7QzjSbYgBZXLALxuAZVH{+}favYfiFp66dokbbQNiTEeWvB/dZI/OuGrUbkkdtOHLBs8Ik8b3Xw21jS9UsYRcPvSPYxI3YzwPfgV9PL8dfEnhAT2sk{+}naHd61DaM9rfQx3iNaTRybp1kUlVCk7WB7n2r5R8ZaWb{+}x0e/d4ba2iu2dnnb5QFVWAx1P3vbrWj45i0vxtef8ACVWovNGiUpaRLoyMbdMhnAX7xyx3Ejd60p04T5ebS99SoVakU47rTQ7n4jatDrGmaVeoTHJcyMiBgFZyqA5IHADKyMCOCDx0rlNDvmsrqPL5IY89KZrXw91/S/Amj6he2OoTSLcW9xb6hf3Cp59oxcKsUbHccY5AyQCDgAjM01mUhhnSM7COx5/lWUYxjT5U7nQ6kpz5mrM{+}h/hnrqzrGpb/AD1xX0B4f1JPKQBsn1FfG3w91prWRMMwGc9f0r6T8H{+}IUeFCXyfr/KvCxFOzuj6XD1FOOp7bpl6AoG4qQAM10tnqgAx37{+}9eX6brIZVKnPrk10dnqu7aCeMdaxg7Dq01LU7aXVQkJC8Z/SuH8V{+}KVt0MaEszHA2nkml1LXBFbuQwAAz171z3h/SP{+}Eh1Rr26/wBTGcIhPBPcmtZSb0MIQjDVnDeKNZn0rXtOubsbIpiVBY8A9cf59K15/ElvvWeOUNlccVH{+}0h4IvtY8DySaLzqFqRLDgZOR2/EcV8O{+}IfiR8R7GP7P/AGVcWbR5DSJExJ7cZziuilQ9uvddhOryLmaP0P8ACHj2zSQR{+}epYH16/4V7Jb{+}LFudNUo48rGQQcV{+}QHhX46eL9HvV/tJJLlQwJ3x7HH4jGfxFfSHh/9pC7bR0kWf5duSJOCv1FbuhUw7tujJuniEm9LH1v4s8TQW1u7zyoo6jJ5NeTXfiyG{+}mEu9VTdkfSvlT4ofH/xR4iza6PuMmcB1TcR9B0/OuO8Nx/F/V5UFutxhjgvcEMPyNH1W65pySNo1FH3YK59t{+}FtYTxJ4z8uBxJ9mQyOf0r0htcGl3KxyybC33Q3pXkn7O3w71Pwfoc13rE/2nVrwhppW9B0UegFeoeI/D9tr{+}nNFOWjfqkiH5kbsRXBJpOyNZb6nYab4gMqKTJlenBzWg{+}oZQ4PHtXguheJ77w3q0mk6lICyf6uXBxIvY16Xaa6LmFSGBGOMU1J2IlSW6NjUrlWUgYznpXmvjPVFSKQcH8a6XV9UMUR55xzivGfiJ4iS1gndnAyCBz1rBpydjqptQVzw/4veI1ikmAfr0zXibTXOmeEde1qyFy19CiQxGBAyL5rFGMmeQNuSCOQ23p1roPG2rvrerzBSSgOOe9czqvibS/Di2umavC8mmX{+}BcBGYFQpxvGOSVDsR74r6bC03CFkrs{+}Vx1ZTm23ZHB/FC1Ex0XUNxa4uLRVnTYMIwA6uPvMWLk556daytI00yWtw6j5kdSD9Bk/zp2r{+}Jor3V91hbtFp/lxo9tLIXDsqAM3PTJBPtmuqj0QadpBMXMFwqXUUxP3kY7efcEEH6V7UW4QUZHzziqlSU4kXhi9TTNRikbAjkIQqeMkHGPxXIr2s6JBrOieeGybdwrY{+}8eMBvxwh/wCBV4beogWNQCPLYSk56qMH{+}te7eBb7{+}2dHkttw2Pb7HZQDuZAGXA9wAPqK5sRpaR2YfrE47xdoU0Vs8wiCRoUljZOcrjJH6kfhXkd6h03WZroANtYSYHRsjJH5E19LazHG{+}lwyFVEaLjGMbju6H6Ekf/WrwPxtpv2WKWNeAJdqhu646n8KdGd3YmvBJXXQqT2kcqL5LB18sHPXPyn9cc/WqyA30CxsQJMDaxPfHGfqP1zWlogNxoli4GZIt8fHcjBGfw4/CqcoWxu5G2sIgADGOu3g8e4Nap62M{+}lyrbYVWwQj4OQfUV3fwV8Qxaf4xitb{+}L7Tpd6ypPaE4EgY4K57HkkHscVxWrwqsquOVmAww7nsR/nvVe1kLR{+}dG5WWD5yR7Y4FaP3lYz2Z{+}u37Mnj5vDV9J8L9cuZXezhS70K7ujte5s2AO3P95MlSP9kjqDX0lvKNlv4m{+}8OnTivhXwWs/wAW/wBnvw94t8PFpfGXhlvtFk0bHcZVwZISe4fJwPfGea{+}wPhX8QbT4o{+}BNI8RWQVPtaAyxA/6twCGU{+}hDAj8KiL6GFRdTsTg80mPl7AUgjBGVyp9u/4UPvVSSoYD0rQyGJyT9T/h/Slb07UifKqgg5x6U5gce9AhvFNJFPIyvpSYwKAGDFNIwfapOnbNI3PtQBEMtRTiDn0ooGXjnj{+}VKB3pCOBTsYFAhCO5pRwT6UvajFACYJ6ilIxS4pSM0AJk4pHBK5AzjmnAYpPrQAqtxnqKxvFshi8PSIrhJLiSOBSTjl5FH8ia1m/dqT/A3H0rj/AIkaotk2iRFQ6m5e5cEZASKNmz{+}eKT0RUdz8lP23vEq{+}I/jv4klREnit5o7KJGDBkEaBSDj3/ka8GsbKa7fcpEUIwOOMD2H{+}etelfEqb/hMvGms6lcS{+}Q8moXRlYHOcynjP8RwMZ56Vmwiz0uAOY2lwQiA8mRh2Vfb3Nc0qvJHljuejGlzavYdpukw6Np63MygZOYomHzP7n2qNL9LWO51C5OeRiPs{+}DwP8AdHp3wPTmGW/lZmu72XIPJOenoormNW1N9Va4ABWGMqiL1Gc//WrKlScneRpUqKCsjP1C/uNf1mWWVy2ScA1N4Ti2{+}J7fK5T5sg/7uaTRrDzdUjVh97HGeevaul8HaUx8WWW1Mks3UdD5ecV2VJKMWl2OSEXOSbPZfCGjsdT8MwvwXt1fHTr82P1r1b4oIkWkxpEyLG0NwAM55IYMfwwRXG6Za{+}V4o0RIMFbSzWTHdsNtx{+}uK6f4s6lE0NiYWwTHcSMpGCcu4/mTXz7leaPeStE{+}UvjHG0dhZRM5yjmQKDxgjb/7IK5Dwh8S/Efgaa0fSdQaFbWWSeKKRA6K7p5bNgjqV49u2K7P4z82lg{+}D8yKvI93J/mK8kr6KhFSpJSR83iJSjVbizb1bxrrmu3kNxfancXDwqY4Q0hCxIWLFVA4UbiTgdya9/8LTLqmjQncW82MMD35GRXzKRmvbfg7ra3GkrBISZIWCA57DkfpiscXSXs7xWx1YCq3Uak9z03RIDDcqpUqvRW9a9f8M6o9qEVW{+}UY4Y15lpsYnkJH3l6e9d9olqWVfmHGCD6818zV1PrqMuVaHrWka3IUU7wxI7966iz1t1UAna1cBoSOUAVQR0NdxaaXIbHeAMerda81xs9D0FUbRX1bX57u5hs4fnmmfaiqfzJ9gOtd9oO3SrSKLeWxwW9T3NeAaz8QLL4eXl/rOu/6LaL8ltI4O3aOuD6lqbof7VHh7VYI7mGUTws{+}DIrcD6{+}n41v7GclzRWhKam{+}W{+}p9MmdLxSHQFD1BGaybrwJpWqKTPZRAn1Uc15HaftI6fc4W2EZOepYE1sW37QB6BIHBGRk1LpSR2RwlWS91Frxz{+}zDoeu6c89vaxxSHoyKOvSvCp/2S9UXU2SOaTyC33QK{+}jLD9pBYo/IvLC2miRuhJDfpXT2vxr8LX9vM4011uEUOVWUYIP8QzyfyrsgpJaMwngcTHVwv9x5J4A/ZetNGiWdovNmQZJYf416lpXgm002IIsKIRx06GsHU/2hJDHIlhZ21qhwpcDcw69ycZrjNT{+}NF9Er5vURzk/KAa56kHJ9ztpYCvZ3SR7U1ktvEFU9{+}lUL6d4FAwCOtfN2pftJXNoxT7WzBcbioBI5xXnHjH9sOfRonaeR3kb/V26gGR/wAOw96qGEqz{+}FHLXo/V05VJaH0h8So7TWLAkP5F/bEvFKvUex9Qax/APi{+}4urbyJWcTIdpGODXn/wCzjF4w{+}Ol5NrniO0Gl{+}H0AMELSN5k{+}e56cV7lF4PtdK1B1t4dqoASVrKrD2L5Huc1OXNtsZ2t6tI0JJLAYySO9eAfE3WJbjzVXcAM8D0Ne8{+}L0{+}z2bKAMkdTXz74vxNISADuPHanRV5XFVlaJ5FcWgiLueWY8k9q8W{+}LOpfbPEUcIOVt4gMehPJ/pXt/iy6i06KU7gqIpLegr5n1a/fVdSuruQ/NK5bA/SvrsHF35j4rMJ6KCHafF5soUdx2r0hdWmufh9puXUxWk9xZqe{+}0{+}XIP5HH4159pv7kA/xsCMe2Oa9Zj8MafpnwCh1036PqF74ha2SwCkOkccGWkz3BMijFdFd/C/P/M5MNpzehx0N0JngjIw0hWJiO2QR{+}vH5V678GdXli1EQgqJowUCMMKXTlSf{+}A5/75rwqS6Iu42Und5qEH6E13vgXXDpHiVZ2IULJuJB5BDEfqCfzrOvHmg7G1Cdpq57zrtvFKmoWSjyotguYyTwVY/dH4ZP5{+}1eTfEDTl1XRIZ449s0JeMgDJyOR{+}gPNe3au8bwR3DbNyZQ7f4lPzofb5WQfjXBavpzvql3amHb9otmntwBkEqd2M/Qkfga82jO2vY9OrBNW7nnHg/RgvhKWT{+}JLkOwx0BBH{+}H51z2sWUi3M5I{+}ZGKFc9epB/wDQq9n/ALIgh8PXqQLtWKFeV{+}7nKYHPfANeYeLYQwtblTnzYip7/MjMp/kPzrqhU5pM5p0{+}WCRyaK15ok1ucma2PmRnvt7j8P6iqej30lnd{+}fAQGUCQqQCDgjgg9fmwfoTWmAbFjOjqMYDdwVOB{+}RHFZdswsdQnSMlQ4ePaeoVlPfv1rtj2OKelj7w/4JmeM/tFx4l8HztII3i{+}0wnPCkEcgHuPlOfavqnwdFD8Kvi1c6VHGLTR9dkaQQou2OO{+}xlgo7CRMMPeOTuTX52fsM{+}IpfDH7RfheSWTy470NbPtGN4PyAEfUA/QZr9OfjNojz{+}GrvUrcE3dmhuYAACTLD{+}9h/wDaie4kI71LRk9XY9Wj2hiR0PNLJluOPWsPwRr8fifwrpmqRsHF3bpKNvuAf55rdUFtzE9eAPaqWpg9HYYAP/10Hkn0pzLwKD0piGU054NOI96D0oAaBSdDTwOKaeKAGgUU6igC0OTjpTvu8CjoM0o65oAD05oFBOKPxAoAKCcUCgjOKACgjiilAzQAjsojDE4UEZJ6da{+}Lv2v/ANqTT/A1xqfh3R7xLrWLi3azcRkubMNguDj{+}MnjrxXeftj/tGT/Cjw/FoPh2ZY/E{+}pqds{+}NwtIuhkPv6e9fFvw{+}{+}A3iH4vXtzf6Vb3N2QWe81u6OZZSxOMMx2rxzkZPJ5OawqS6I6qUPtSPnt0eaY3M7FAxLBcY4Pv2Ge/Ws7UtZSKQlmBwMhQMED0A7D9TU3xLN/wCHr82TIIHEkiMxPO5ZXjbv/ejbk8{+}9cYXYylQdwJVS2c7j3/Dj9aIUb6yNZ1fsou3mpzalKWY4RsBUHQY5OKv2MStp9wc4bO8598kfyrO020a4u7ZSCFY4yOxBrZ09Af7QAGYxGjD2yGI/QitnZaGe{+}rJvDVkJte8tiVfyiVI7kf8A1hXoVlpf9n{+}JbGdBtImCsBwCDGR3/D865HQLZI9VtGYlHkdEU/UP{+}nT869RaKRorWdV2PbzkNnkkbcp/7KK82vNqVj0KEfdv2O9sfLF/oBKlZJAys/QlfMBAz9cn8qd8T5kXULZV5MlpMct6mQt/n61kW{+}obdU8PTMv3XZSp/hOEbArQ{+}J1uq{+}IY1k/1flSohHQ4VSM/X/CvLXxo9OXwnzp8U5FurS2iyG8uJGVlHrk/1xXk8i7Dg4yK{+}gr270ObQLrTdfgdLdzE1tfW{+}PNs5MFWbB4dGwAVJHTgr1rzTVvhhqSTqdOkttXsif3dxbzKvGe6sQR{+}o96{+}joVYpKL0Pm8RSlKTlHU4U11nw81Ke01OaKIOyMvmEKCQu3kk47YzVc{+}AtSgVpL8wWMScu8s6Ej8FJ96be61DpNrJp{+}jzOYpMedcn5TJ6ge39K2nJVFyR1MacXSl7Selj6S8IayLkx7369QOvFet6G6lEKFWB6HH5V8p/DvxY08cTFtrA7W56GvpPwbq6TW8a8bcZJr5rE0nCTR9dhaqnFM9a0SVkC5HDc5PFev8Ah62F3o5CgMoAxXhek3YVkwSNvr{+}levfD/WAwMLNyw4yOtePJNM9WMkT{+}M/hHofxM8ByaBq1okscmSp7qckg1{+}dvxb/Ze8Q/DLxFdN4feV7eMlhGj4bGe3rX6ix3BtmVCcjpn2rzv4r6JHeQyXYjBm8vCsPzx/n0rrw2LnQ0WxrSoUa8{+}WsvR9Ufn18IvDl74x8aNpFy0tvceUzRoWKMXUjIP4E/lX0D4S/Zo8VXviK7jeWZLONV8oq5B5Hf1q/oX9gzeKYNQuLIJq1hkw3SriRWPB{+}Ydcg9DxxXvHhT4p6ho{+}pT3FxHHqllNGhihi2xtEQuCQ2Duz74r0ZVY1Hrpc9PFYLH0E5YZ8y/E8u8MfsmeJL/Xr621K{+}uBagq0Lq3JBHOSevNdRrX7Hur2H2f{+}zr{+}4kEjKkimUDCdzwP517bo3xs02PUJmvLOeGIopiEX71z/eyABjt0zWpd/H3QDPbR29rqk0jTKjH7C6qqnq2T1Apv2aWsjw3iM3jJe6zwyX9j7Vbe1ldNXuVcglenHHTOPr2rh9T/ZW1{+}DwlJeXF8V1FY3ZTC7AKQTjI98CvsHUvjDotvakqZ7liCBHFFgn/vrAryvXPinqGq6bcWz6WllDIhUS/aCzcn02jHHfJ5qZezSupG1CWb4iSTTXrofBHxi{+}G2o{+}HNG0qC0jnn1rUbkRrFESzN8pJ47due1aHwd/ZXH9rWup{+}Jyt5eSuD9kzuWPn{+}Inqa9/8Q65Z2UhjiHnsqbTI/wAzD/gXXNafw3imvbtJGXO05HBHHasp4ucafLDQ{+}gngYUIe1xT5pLbsj3HwtpVpomlw2VnCsMEahQq8U3U4FRmccMeckVb05hHAqg4OMZNZniK7RUc7h8oxgda8WWruzwVL3uZnmXjq9JQpu3HvtrwXxfItu7ux2so/DNes{+}NtTRpHJIxz3r5z{+}KnimGwtLieRyAgLdf0ruw1NylZHFiKnLFtnlfjjxrZWWt2sFxH59s8o{+}0w92i6N{+}Pp9K8p8SaF/YmsXNuhL227fBKejxnlSPwIqnq2qSavqk11L1kbgDsOwrc0XxHI9mNOvbdL22Qjy/M{+}/GAegPp1/Pgivr405UUnHXufGTqRxDalp2ZHp9kTBayqpdmUgADJ3ZIx/L866/xRdvHoun6KGLQ2Cktzkea{+}C5H0AA99oNQ6fci2s/NtLaOFoywCld238yfT9KrXsRn0uaYksTlmOerdP61LvKSb2RpGKhFqPU5VTm9hZugdWIPpkV0WiXO6bdg72IXJ6ZHP8AICubgIN6jAYAYEk9Bium0yGSKytnZSGM/nYPoTj/ANlrWpa2pnBan0z4Slt9X0G0SVmb7RYqDzzujYofz2R1S1WwmjS0uVmPn2TRkODkt8i9vTgjPrJWN8Ob57W{+}0xTkqGZUB{+}7g85/8eH510PiC5EV6rQMxjI{+}zSh1wQF27WH4j9DXz{+}sJNHvq04pkp0tD4P1F2Znk2GdXUZDgvwP5/hXiursJvDgdlCtHdsNuOQrAEH8xXvqMYvB1w4ciOCF5HZD/AhH6nJ/lXg9xaGOLVrdt2I1V8b{+}mGAIPr3rbDyM8QrNI4zUQsVq/mE7SoTA643f4H9KwtTHkX8UuPvIPpkcf0rpb9CpRdqncv8YHQgVg6uit5Q6hd6r7detetTPKqxseofA5vsnxM0fUYwv8Ao93FNGCCUXFwgIP/AAEk/iK/ajWrWO70KXcqNiISNuHAwCf5D9a/Gj4CWBubq5m3IZVjkdQ2SVxtGQP94L{+}dftFdRBNMnhdS2{+}AsQeSfl2nP{+}e9F7tnPUVuVnD/AEGy{+}H9ppxOWtXkiTv8glcL{+}QH516cfugD8q4T4VWD2nhnTGlIaV43ZyDn5jLuP8A6HXeY5NOOxhLdkbDgDpSdsVIwpp4FUSMIwPekp/50088dKBjc80mKcR1FJ0H{+}NAhMGilHFFA7FsdKOlFLxnFAgABpQMUtFABSAYpaKADtzVLWrma10y4ktoxLchCIkboW7Z9vX2zV2oriMSKAyB1zyD6YI/rQB{+}YnxvnvoPixb6t4lnn1O4a5InLxgRIUYYSJSMEAYAHU4c9q{+}6fh15dtpL2{+}nol3YtEhi8pFQNGRuXoMEhGX5h1GOvbl/2hf2brD4m6PeS2MwtLyYfMxUEFgDtJ4yMdj1HPUEg{+}O/ALxlr/AMNtfPg/xjNFZ6lYruWSZsLc2QPEqPnJ8skkrzgE9hxzpckvU6m{+}eGnQ{+}Of2yPBMnh/x34gkbyxFa67OqqEKssVwq3MWe2Nzz9OhDeor54t8FiOd4dsDH0xX6Y/t9fC{+}DU7R/Eaxqi6lZ/Z5XAyr3MO54GJ5wTHJMqnuQgr8zoiyXfkyodySMpzwcHiumDvp2M5LRPudJaQNbgRAYaXYynHOCxHX3P8AKtjR7IE6mMjelqpYf8AYH{+}orKS4UXlkwcBdhjYPwMgKwH{+}fSug0NYm1fVowA8htXXae/HB/nXPO6Z1R1QuiOhsoJGcebDPF254J/{+}vXsE0ZuLJkHJZIpj3J2fKB/46eK8a0F2OhXkqruMNyh68gAt{+}nNew6H5k2noUO5ZLQtwMk5Abr6fOa4sStTtw0ty9fLG1nbuozPHdRunOCyspBPH0Wtb4mJJc6j4enCuIrpV7feJQg5/wC{+}VrHuSo062m2b4/LidnPfY{+}Dj/gJNdT4wVbjQ/D8zB91tsjHHdXCH88fzrzZe7JM9Fe9Fnz74mtDe6deQAAOjFQCcdG4/nXlU13cWbkWlzNAoz80TlR29Pxr23VbPfPrUIGJBM4AxyCQSP6flXkF7YF7xvJH{+}tO8jGQF9P519DQ1jqfP4hWldHN3tzPcsDNLJK3XMjk1Wq9qEIWUjBA6ZA4NUipzxg/Q13JaHmSu3dl/Q9Xk0a{+}SdDxn5l9RX0j8N/GUc6RYfdGwyGzn8PwxXy{+}UbP3T{+}VdL4J8Wv4dv0WRz9lY5yedh9RXFiqHtY3W56GDxPspKL2PvbQtRWVI8E5HXJr0Xw3q32S8hkVhtU4yfSvm/4feMftEcQ8wP6YPBGP1617Not{+}GAHGDz06V8lUp2dj7SnNSV0fQg1IXNpHIrZzz171U1qAalYOpO4sPfnviuN8Ka6GtzC7llAOOef8811FtcExHac5FcdrHSmfOHjTQZdA1a4baY9xPlTqOD/ALJrI0LxhNp160UxdYuP3TE5NfRvibwjb6/aMrxgs4IO8ZGa8C8X/DK{+}0uRvLSSSDOVYfeX2B7/jXbSqq3LI9vC5jOj7k9UereGfF2mXl1b/AGuWWKH7zGEBiOnFdTZa/oFtI8l9Bd3shKGNIpVhTH8SsSpP0Ir5j086lpFwCzudmOZBj/61d9ZfELTprX7PfWKrJwfOinY4OecA9umfxrf3U9LHryxeHqK/M9ezsep{+}JfFGjzLDcaTaSWEQHzQy3Al5x/CcA{+}vBrzrxN43nMUsMMeWAI2dc{+}9M8X/E228RtBbafpUNja2/yoYYx5rjaB8zAY7Z{+}p61QstIutUUM8PkI/wB4kct{+}NQ3FSuyY5lRpUuVLX1v{+}Jg6bb3GtyRgpvc8AA4B{+}vtX0B4A8NjS7BC4O8gZbpmsbwR4JitVR1gIAOenJ59cV6fDAttBs29O1cdSpznzuJxE8Q7y2FM3lJnJyfeuJ8Xa2ILdhuxkZwa3dU1NYY3O4AAV5D4514bpPm4HXmsUrs4WcF478RGFJedo67vSvjv4x{+}LW1q8a0ifMUbZkI6Mwr1v4x{+}OjbQy28EhMzjaOenvXztfWzzwyuQWYDJz3r6fAUuX32fOZhUbi4ROfPIFa{+}kp5lyvBO7jisoIOR6VsaQjLNFtHXow4welfQStY{+}XhdM9GSFYNFluFVlEiA{+}vzjhv1wfoaztUAg0LpsM4646Drx{+}NVrK/ku5hafNCu4Boye/Hf0xzXSeNNOgEFvbRnfthDAqMduSPyUVwSfLJJnqxXNFtHmdlEXuEOOCwB9etdrLAtppo3EPswo6AksAf03E1l2GkmO7gwuCr/d7DkDJ/Oug8RBFKWy5CGRkIJ5zkAn/AMd/KlUlzSSQQjaLZ6H4euDajTZQfnYF4l68hFxn67K29Qu/7R8Sala5aOSMkhOckldyNj2JYVz{+}iEw6L4cu85Kz{+}Xu56qRn8Oa0b9XtfFdlesdqXcCE4PBwDz/SvJqK0mevT2SOi8Q6r9h8CXexyq7TEyN2LMPz5rxy9uXTWpVb55LmHkdMEruP/j2f{+}{+}a9A8Zv5XhmaFX3BzHK3qMuT{+}OeK4oaVILu2vGG5kZIm3H1LAn{+}f51phl7rJxL99I5bWQpktiQDhCCR2{+}Y8fkcfhXLXsnzhOQcNyf8AersNXjWS3YYEUqhz7dTxmuIvVMcojIIkXGR9ea9Kg7nl11Y{+}o/2afCc974c1y7WIb3gt4Y5D94hrlVOPxAFfrHqNwyl0DK6qoZySQdozx7c4/X0r84f2adKvrTwj4Ijht/LTW9UELk/MXhR0nwOccbTz/tiv0avo47Tw7eXkzh5hEzyFT0IXt/n8qIu/Mc9VW5bkPga3ktdB0yJgocxySNg9S0mev6fhXV4x161kaDCUsLZcf6uBFHGCDjmtcfNg4xWyOViYx{+}NJgU9qYOgpiGsOfakHP8qccAgfpSAUANoPSlbnFNIoAAM0UpHviigos8Z6U4UmOaXvQSHenACkPBpRQAnHekxg0/pSY5oATHrS4FLQRigCKa3VsYAHI4xxXk3x0{+}Bdp8U/D2LWU6br9oTcaffIC3kTev8Aun7rDGCCcjpj189RQyBuv51MkmrMpNxd0fCVjrOoeOvhp4l{+}FPjWJbLxTpik2fmKNhdBui2A/wAJ28H0bpxX50fEXRY9M1Z7oRGFmkaOSMAqUcDkEHv1/I1{+}y37QnwUXxtax6/pLtZa/YrxcQj5nQcj6kHt3yR3r8zf2gvBEl3q2plIEXVLg77i3jU/LOucvGD1DgEjknn1BxyqTpVFc7larB2PBdQPmSQyw5Ec6LJg8bX5GR/Kuw8IyLc{+}NIJpMBbi238jCsAQW/wDQTXnolxp8ayHm3lMTKp6q2T09iD{+}JFdj4dudt/pjkqjRyPA2D2YHp7Hd{+}hrqqRujKnLU0tHKpZ{+}ILWNhvhjRsHrgMoz{+}WRXqfw31BZrXTQ/3EXYy56gO6Eflg149MTp/j26ib5FvYZImB98gfrXSeANfW1iMTYXZcEOWPRCuTz{+}Z/CuStG8TroytKx65Dm2Z7ST51tJ3tsD{+}JGAKfnlv1rr9QtVu/CFwnyvNbEyLjncGUNx/30fyrgIL9by5kJKySvbpIx3D/AFkTFS3/AHy5JrvtEv1dEjZg6TxbQzDksoKj81YflXj14tWaPWpS3ieI{+}MsQ6/PLwFujG69wWYEDj65ryi{+}TbflTyHbZgg4UZPX26/lXtPxK0SS1jgYknyHG4jJIG/K/h98V5u2lI2pyzSDbuR5PXkZ4/Hj869jDT9y55OJi{+}fQ8y1bc95J1IDEAY9KzmXsRXQXVpIruxU7UGckZz3/UmsaS3dSoYYPfPFepF6HkSWpVIwaSpZUCE4O4{+}3Sou9aGDO8{+}GvjqXQL2O3uHJt8/I2fun/CvrHwV4xiv7SE7wHwAcHmvhaA4lHavSvAHjyfw/corsWhOBkjO3/61ePjMMp{+}9E{+}hwGKaShI{+}5tI1l7e4WRXXB7A16d4b1xbpVJYH1Ar5j8L{+}M4dTiXZIDuUc5zgV6j4T8RhJ1yxGe/bNfM1KbR9RCS6H0RpqJdoqn5j2qS88IQ6ggLRhwepIx1rD8HaysyruIPIHPevTdPVXQYKtn0rnUbmspcqueV3fwdsL1yWgBz09DVe3{+}BFkrCRYTnHQr1r3K2sUJyE59QKspbpHhcZHUdq6FT8zleIa2PEovhFZWjllhUkegHFadp4BgiH{+}ryewHcV6zPaoWJCjB64FQGyQk4x14qJU/MuNds4q10VLNVEY574rP1qVbVG6KwHeuuv0SFWOcGvN/F2pqiPk4GetYNWdjeMnLVnF{+}JtbEKMpYnP4fhivnb4o{+}No9Ptp3Zzxk8HqecCuy{+}JvjVNNinczBVQZJJ6e2K{+}RfGniu58T6k5LN5IY7Iwf1Pua78NQc3foc1eqoaGDrF7ceItTeaXJZm/IelXovDrNbEsp7ZNb/hDwk94wkYct7V3V74X{+}yae5KYOK9SddQajHoecqDneTPl/VbFtP1GaBgRsbjPcdq1NMjJtm8tvnSMyDHop5/Tmu1{+}IvhTzdPOpRJzCF347qTg/rXFaGVke3R2GNzqfQ5HSvZhVVWnzI{+}eqUXSqOLPQvh1p0Os65Z3RTzJFJVo8ZDsflH{+}P4V6FqHhNtRv55RGZY1PlRAHAJHC4/IfkfSuC{+}Dm2311Le4DiJGeaTbwQgHP4nPH0r27QtdhuiwkjEczs8qqOiqq7UUfgCPwrhqO03c9GhFOmjy{+}z8IS210XdQVEpVmPQN6D865XUrY3{+}pHazODkBvYnJP8AL8698{+}I9tZ2Wm2tjZJuwgBZjy7k/O7Hvzn8/auS0fwXLDELiWDiMFjkAr5oGFU{+}y4DH1O0dqxjUv7xrKH2UUnBsvCNkrYCxyTSqM9FYquR6nv{+}FT67ctKLGdSrCCNPL9MbmH8wKPE8UMOn2diG3YRMAEZ5wR{+}ZJzWdqlzHPBYxIGQSxmI{+}gIOQf/AB4VhLWzOiGjsaHiq4Mum2oi2yyvNFEiNzuJZgOPQGqGr6FNppuxh1jikR/l5B27MnjtuyM1r6M9tqPiHRVuIT5cMvmsZGzkpuKD8wPWug8R6fJceHZb9Q7NeSvk56J8xH6oT7ZFVR0gFfWVzxHxFatb2Rcjd5obHfA3YwB25zXn92rXGriHl3Zwij3PAFexeMNMW2s4GyMqNuTxwPm3fm2K4L4caGfEHxU0CxeNSJtSiMit0K7wSPpgEV6NBrlZ5mITTR{+}jXg3wkvhLwv8AD1rYyxXemWVzrsnlAkfZkZYgAuTjehJ/AelfYeqyDUNFsoA5ZLwRcKeqnG79M1wXw{+}8NWc{+}u{+}IPMhVrOx0u00GKNM9AjPLj2zKv4g56V0ngGHfptjYmQu{+}jb7OXf13RZiH4lcN9CPWqgrfM4qjv8juLSDbAhI2sQCTj2qbBAFPC7Vorc5hhGaQrxSn71FADMfnSYxS9yDRt4zmgBp9aaVBPSn4zRjnPpQAw49D{+}dFSYz2FFAE2eaUDIpeKMYoAWjHeilyPwoASjFLxRkAUAJSgY680lOoAQEE0tFFACMoZSCAQeCD3r5a/ar/Z7g8T{+}G7/V9MtlNzaRtceUF{+}/jkgEcgkZ9ecHFfU1VdVtEvtPuYHGRJEy8jPUVnOCmrMuE3B3R{+}CHxA8Py2GqXEpt3NveI6K7R7HV1OQj44LZGM9fwxnmNNlcWyyo3K7CV9Sv8Aj0{+}tfUn7QHh2Ly7q8SGOHZcGK68sNsjkB{+}WYcdDtwR2wfx{+}X7q1fSdXvLWRVQN86Ec4HXj6dfw96mjPmjbsddWNpcy6mx4znafUdJ1bzMrcRB{+}P4SMZX65FWNJuFttYlAOY7iMSD25yf54qtJC134XnjzuazkErJ3weGK/jg/Q1lafeGG7hUtkpznpkHhv6UNXjbsCdpc3c9b0bVxGbadsD5lVyOOSDEw{+}hJzXqPhLVUk06B2/5YzLOFHVUZiHP/AAFSfy9q{+}f7bUPJS8jYDAIYfiM/zGfxr0Hwd4mWyWOeQlom3RyRg5JG3PH5P/wB9V51WHNE9KlOzPTviT4cFz9oeJnBlUCRD0HQhvfnI/EeteD6vp4toNQXBDqFQA/5/Cvoy11p9b0bT2klSVdpt5JT0kZT/ACbCsP8AeFeVfE7R309HZFyZnKhs8Buw/HGPyrDDz5XyM0rx5o8x4XcxE2M0z4ALAsvr1Jx7dK5p1DASEcbjyOK73xRbrFLYaYnySCPzJFb1Izj9B{+}Vc9qdnFa28cQG1sknAzxn6/SvfhLQ8Oce5gzacGdQpPIJJx0AFZjDknINdleQ2tppMjbj5shCrheduM/5{+}lckQjO{+}FPfHP/wBatou5z1I2ZFH98VsWLbX5rJGwDPzZz7Vp2uCAc/N7Upq5pR0eh3XhPxVc{+}H51CszW/wDEgPI{+}le7eEfH6XpieOQHpgk1802MhYAZI46iug0jU7jSblZoGGCfmTs1ePXoKWqPfoVnHQ{+}//AAB4vSZYjv5OM4Ne{+}eGfEMbIoMnP1r86vAHxU8i5VWYoxPzRuefw9RX0v4O{+}JsU8EcizdcZUnpXz9WjKDue9SqKcbM{+}urHU0cJkgMB09aui9jIz{+}ArxPQviDC0a5mGcZznP511Vv4thuItwdcH3rONRpEyoXZ6Abpc8Aj2qvcXaouM4xXGjxXAFz5oB9CaxtW8f2ltEd1wCR79KHO440bG34h1kQq5yPYd68D{+}KHjWHTbaZnkAIG7k8Cn{+}PfjBZ2VvLNLcpEi5G5mxn{+}pNfIHxR{+}Kk/i64khti0dluIyTzJ9fQe1VRpSqu5dWpGlHzML4leOpfE9{+}6o7NbqxIwfvn1P9Kx/CPhObWr2M7SRnJOKraRo8ur3qxxoSTjmvpX4X/Dg21vE/lZfjnFepWqRoQtE8{+}lCVaXNIr{+}FvA4hgiVYiGxzxV7xPoDRWLgLjA5BFe0af4U{+}z2{+}9UwQO39K5nxfoq/ZGG0nIIrxfatyPXcEo2R8/HwsNX8O3luVGXjZcEfl/OvnvxDoqaLPZ3dumYJMMy9lfoR{+}eK{+}ydC01IrC93rhUJ/LFeK6NZw3/irUrBLOO{+}kT7TsgkA2iNjndkkAEZ45HOAOtexhK7jJ9jxsbQU4rucd4NnS2sry8nBClVjJHVh1wD6ksv613PgyM{+}ItXLNmPeQo2dEwdxPsMDr6cd68w1zWTFcNCCqjzWeTACqrE84A4xyAMdhXWeFfEI0223szRsUDN1yoPReO5/oK7asW05dzz6Ukmodj3eeysyqRLMtxdw8JKyggkcKx{+}nXHcke9O1aGOC2azgCPBAg3gt944z{+}POT{+}Ga8us/G76dGLl18{+}VhsjiA5Y8ADFWL3xiNP8AtVzeOZdsRlmdD8u4rwq9uuB{+}NciUnY6pOK1M3VT/AGj5k7ghEXzPlGM4l3Lj2wlV/E8KW1nahCMiXbnPzdAfw6VEt016IQwCzF1Z4gMbQQ{+}VHtzj8aPHtxttbMRklwVL8c5IAP8AI1q1siIvqV2vpFvrDynVQCpBXg8sxP8AP9BXsV75DeDNHUSItxPbH5DkkFl4A/E/pXg2r6nFZRac7MAAoUtjBwCefyIr13wZq1r4m1bRrKJg5lJd3P8AAuTkY6fdTP8A{+}utYK1PmZE5XnY5z4waIbdobSND9oeMBkA5Ygbf5lj{+}VYv7MelRQftAaNLdQqyQvLMExuBKxNt/Ns/0r0nx/MNX1HU9TiDTL5Lm33gDGOEbjoCScfQVh/s3w2ug/GfQ766cPBBfWtsxxgszDj8M5rajK1M5cQuaVz9VfhvoD6D4SZJwGv7m4lnmPJ3SMxJx7U7w3YLoPjTU7YbimqxfbSzEkNMpCyH2yDFx7V02nWYgijyAscCBVI6ZwAf5frWN4hiNnHb6yMq2nXPnSAd4GG2QH2Cnd9UFdtrJWPI5uZvzOmXOMenFKwwKXOGPOQeeKVulWZkZBApKfQelAEZIHakp3eigBnagn2NPoI4oAjPPrRTgCaKAJ8UZ5xS0CgA70vFHPrml6mgAxmm45px9qQDmgAApR0oo6UAKBmg8UL9Kd25oAavNJJwjH2pwGD7UMMqw9RigD4A/aB{+}HH2rXfEVlbtFE0im6RmGCFLnjk85ZnHTovvivzz8YaXc2slxFNC8N5p820qB0jJJXHt1/MCv1g/ac0K40u9i8QwIrk2nlMHB2owkVovxLlh/wMV8LftNeEo7DxLc6jZoC0qs7ryBLDjPfuCP0rhX7urbuekv3lP0PCPCV3HdLJG74LKI5M9wTt/kf0rHvrd9MvZA{+}SYZPmHqp4JFOH/EtvEu4CXgk4bI4Knp/n2re1q0XVNNhu0H73JXfnOR2B/Hj9K7L2lfozBaq3YmSFnnt3TEomj2deG6EHNaHh2{+}kt5Z9Ob/WA{+}ZASe6np/M/gKzPClwt3o7WrKRdWr7kXvgHkfhmrmrRtaTJewDLI{+}8Y9PT{+}Vcb0fIzsTdlNHtPw11iLUtMv9Ec4uY1FzalmwGAGMfoqn/ZU102rWZ8T6CY2QG6RQHHRm9D6g/Lj2Irw/TNcOmahZ6pafJJA4lwvAKHhhjvwa9qS7ia7hvopB9kmUP8h6Z5H4f/XrzqkWpXR3wkpRsfO/iSwli8TXEs3LygjaRgZyent/IjFYOsrGLSzZyyyPu4xzweeOte4/FvwQ1wP7UtmLx45jCgnJH88/5614lJbTIxe6VlZU2hGAwu48f/rr16NRTSZ5dWHJKxl6lZ7bQYLNgKTkdsVzsNszPcsOiISSfriut1Zh9lKxqRlF6/rWDYoyaVqUrHG4rED6/Nk/yrqjLQ4qi95GK4JKgdav2J5Ge1WNA0dtZ1G1tEwr3EgjEjHhMnr{+}FOFr9nuWVTuUMVBIxnB4NaSfQKSe5q2sDKoIAINakKkIFP6UmkwieMDrxVya0aLDc4HrXnyld2PVhF2uJAxDL2PUY6iuz8OePdU0Yqok89ByAx5H41xMYKj5iMdsVZhfyWBJODxj1rnlFS3OqM3F3R7xoXxue2wZfORuhK4YV18Hx/jMajzpRnsBj{+}tfNVvdgA4HStK3vCq57iuKWHg9bHXDEz2bPf7348SuhET3Lj04Fclrnxl1W6R1gXyif43Ysce1ecm6LJhW/E1DLJhcnvURoxXQ0eIm9mO1fV73WJ2lu7p5mPdyT{+}VZdtYm5nVdvB71aYksq9eOmK6zwX4Zk1W9QOmQWBIrZtU0Zq82dt8IPh6bydZvJ{+}XPcV9YeFvCkVjbR4QAgcmua{+}GPhJLOyjVEz8vU8e9ez6ZpqpbggEnHp/KvCrVOdnqQioI5{+}609YogFGDnJ215/4tsfMbYFBAz098V7BqlsqQj5QMdzXn2t2jXEzbVPHQVybM7ItSR5CukMLO/jHBbOABntxXyOmp3vhfxBqtwr/vZo5Y3D8gjdnlc4OCBjPQgHqBj7wk0pldgVwTkjivkH48fDxtAnuNSjuEMdwzg2/R0Oclh6jBOfT3r1cFUXPyvqeZjabcOZdDwKS9ee9d5fn2sWwe7VvWWoPFC7M{+}f42A7n/ADiuWc7bkq3XOD9fXP61dvJ5fssSQjBbCt6j/P9K{+}slG9kfJxnZtm9a6tPcXQkJy{+}SoyflQeg9z39q19d18zNZafDKjKnzTuSfmcduPwH5VxyXn2S1hVfv4OCOvXk1LoarLJOrNkycl{+}{+}7HQe5/wrL2a3L9o37p6T4auGubyWTI3ME5YZCDIJ/ln8Kv{+}MSsuhpKCwWMxr6tjJBJPr0/Ouc0WdbexlfO2SdSijnPHv8AQGuk8RsLvwnPtYM0Z7nJPfHtiuCS9656cH7tjgvF2qxxyRRSDMaohCKcDPIP5gGu{+}{+}GF5cLbxcqHkjYKka/MF4BI59f5V5J4lCG4RzJ/CrEHk857/j{+}ldv8AC/WUgFxJIuXEIh467Sx4{+}gAP510zVqByRlzV9T2rWdYtU057WSARwySxxM8ZDEogcgEZ4z1P1{+}lRfChI9Rj17WDASdI1Oy1t0SPjy1EgK9{+}AdhyeBn1ryq01i51fXls4ywJmAUI3I6gn3{+}6v4Z9a9t/Z30dNcg{+}LWgQyws13oE6R5bl3jCN8vXoFck47cUU4WikxVpps/V6zmF1ZQtGwZJF3Djse9OlgS4NxDIgeF0Csp6EHII/KuW{+}DOs/8JN8KfCeqO6SyXWmW7O6dC4QK{+}P8AgQNdcExLKcfwqOfqa7VqjxtmYvha6kl0hIZ233VlI9nKx6sUJUMf95drf8CrarGtrc2PiTUAABBeRpcD/rogCN{+}a7PyNbZHFCB7kZ4PHSkPQ0pPNIeRTEIBRilOR0o3ZXrQA2ilpKAG9CaKU80UAWMDFAHNKOlFABQaKDQAUUuKMUAGM0KMmgZpTmgBcYoo70oHNACHg04AUmMnmnUAeNftD{+}G/7Z8E3CpGGaNSqgjjIKMhPsGUGvi39ofw4mv8AhXRdQH79HjaBpIkA2DYrKTjoMuw5OOmK/RrxZo41jSLi1OcSo6k9cZBwfwP8q{+}P/AIhaJGngYoYYrV3SWCSROiyQNswV903Y9SorhxEXujvw0raM/MHZFbTXenXTKkIbajHqp/z/ADNJbajN4fvJrK9/49JhskX{+}6SMB1P4fpWx8T9HFhq9ywAI80rlTnJ5z{+}v8AMVlLCms6RGJQZbi2jKEZwTGD/wCy8V1wanFMmScZNIiilk0DV1kyHGcB1PyyDp19x39fpXaOEubdDwyPgI69DnlSfTIyPYgCvNYJ2ERtrjdJGPusOSvuP8K6jwlq6lBp1ywZXB8l88HPO38{+}R71nVg7XKozXws1bJvsTyR8si8bc4ypOAf1wfr7V3fg3V5I7F7HeTNCDJbAn/WJ3j{+}vp{+}PrXIyr5dyXlVmdQA4/vqRgj8f6U2OU2kuFkcSRYZWBxuHUMMd8fyrilHmR2xlyux7d4d8TWuq2ktpIFlfy/MQ5wSpA/MdAfTg9mrx7x1oHk3LXduW{+}zmXayMDuixxg8dvpWzJqfnG11WybyZgcyiNej55YAeo6jvW1qYXxFYPd2aItyYws9srfLICOq{+}nsfw6isabdKVzWcVUjY8WYxz2KTGZG{+}dk2jjAznjPSs2OKOaxhtIWO6admI4PHTt{+}fToa6m98PvEZRHC7xFid4HzA{+}hHHI/WuevY7nS5nmg/wBH2qY/LU/PtZSp469OM160JKWx5dSLSuzoLC4vPhxPYa1piwXXLQtOo3pE/B2gAgg7Wzk4zzjOOOZ{+}wOJhkHk967D4M6JJrB8QzXdstzZSwrbO0{+}f3JJMgkyCDwIiDgMcPkKcYplzZiSYkIq4bO0c8dqmU{+}Vl0oXjcpaBBibZwB3IFdc2iLLEDtyf/AK1YFrD9luY3xxnBGK9K0OJb2LGBjqD7151Wet0enQhdWPOrrSZrUklQMdsZBrLcMG2kYJ7CvYdQ8OmWIkoNuCK4nVPDrQuzKvTnkfyqYVk9zaVJo563Vu44rRt1{+}bcenoKatjIh27cgdCRxWla6Wznpz1yOKqUkRGLuQ/aN2GGRjikUvKduCC1bUWgs5UAE98AVu6X4MlndB5fB74rJzSRuqcmzG0XQJb{+}4iCqRnnOK{+}hvhj4GW28p5Ihnrms7wJ4C{+}zOrNHjAH3u9e{+}eE/DyWyINm0D0FeXXr30R6VKly6s6jw1pQt4o1CgLwea7q3ASBeOnQmsbTIdqbcbSOnFbDMAgUcjJ6cV59zeWpl6qPNDLnOTXLXVnukY4BJPpXUXqliAefXPeqUtnuUkjGOc4rN7mkdDjrqz2SM2D069cV85/tJaB9p0F7mOEubYvNIsfJKAc4454P5Zr6ruLQFwpHBzXkPxp8OR3mnTJiXHkStmBXLnCE/wc{+}vBIB6ZGa2oT5KkWwrLmptH5o6pZSWdy{+}QGVWyG7Yz1q2I0uVjmiYHcNpTOPwz25rS8QWJtJPLaTzWCBtvPykjJB47HINc5b35tHLRRYBPKg5H/wBavuYSc4po{+}InFQlZ7F26jDTKGcCNAQjOuT75//XWpaRQiPGxkEa5Z93ByPx9c1mLr9tuDTQMrZ646Vt6S9pdgG3cDncy52k/h3{+}maJtpbExim7pmtpE5urxNhKxoPusPmJ9sZ5J6/WtLUtYeHRktDHhpJG8xienHP86i0y2VnEcexSCGYgbSR37de2BUvikwXkBkij8uJHxk8Hpz{+}tcTs5WO6Dajqed{+}IQVmYkEs0Q25OSNp/{+}vW54RuzpmiXszYIclRkcggdM/U1ma{+}VubC3mTlhlW{+}h4/mBTYX{+}zeHY03FVdznHQZ//AFV125qaj5nHflquXkdb8Ktajh8RXOp3JAFvGSinnnBwP5V7z{+}yn4li0r4l6REqKJL{+}0uYpWY7VkDDbtbPXIGO3X618p6VNJZxXlv0bYxfI9FbB/PFelfDnxL/wj{+}r{+}H9WQ7WsLkJI2eAj4IP4EEVclbYhSurM/W79knVhJ8MbzQmY{+}f4c1i90xlY5IQSmSPP/AJAPwr2fAZ5mB54X8h/wDXr41/Z/8AGM3g/wCOHi2yLifTvFGn2{+}uQbT8u9f3b47ZIZfyr6x0rxFb3VmcOWfnIPXP0qo6o5JwfMyTXJDb3ulTgjH2nyWz/AHXRgP8Ax4LWtuyK5zxPexyWET527Lm2YE8/8tlrbju0ZRk4OKfUhrQmPNFJvX1pcg0yQpMD0paKAEwKQjBp3ekOO9ADfxNFLRQBMOAKWk20uKADBNLntSfhzS4IoAUUtAooAKKKOpoABT6aOpp1ABTCxPTp60{+}kxQAghDLsJ45zXzp8QvBk80Xi3TbdHa4huhqdtuIwDIADz6ZWVfbcPWvo9DzXnnxEs20vxZpOsrkWtzDJpt2du5QGIaJj/uyKAP8AfrKoro1pyaZ{+}Q/7Svg9tA8W30K8QTAXEQVcBlKgjGe4Ax{+}FeH2F3Lpd4ksBYSxt5iEjr2IPsRxivtz9tvwVZaJ4oEEEE8pWT5HySuHLHacddpLAd8Yr4q1SNZbjYi{+}VJjpnpx3/IVnQ0XKdU/eakTara295uu7MCHeSUj67G7rx{+}ntj8ci9nWxkhkjXdb3Ch2jB5jfOGx6cg1csL4wbWZdySDbIh6HB4I961L/Q7fWNOMttJGL5HANschnTH3hxg4wB17jr26k{+}jMpRfxRNnQtWTX7aMSSh7pVChuhcDgEj16fX8qluyXBAIV05Qjt7fQ15qstxot3G8ZMbqMFT35PBFdpBrsPiKNQWFtfYyrfwy/X0Of1/OueVJxd1sb06ymrS3NPS9Z{+}wXbnloScSQnnYT/EB6GuttbzymW6tH{+}XqQp4XPf3U{+}v9RXll/eOsg81RFcJlQ{+}OGx2Iq1pHiaWwnRt{+}wZywJ/M/T1rGdLmV0dEKvK7M9XldL5ZJ4UzME/exHuPw{+}8P9ociuW1HSI7k/J/owTJ8oMWYnBwc85Htn86u2Wqq6LPbzLGx{+}b5DwP8A61bca2utYkYpDeEfeH3X/wA{+}2a4lJ02djgqhj6HffZnhspbOKeOO1RYlC7VneN95E2BnBDSgkHcQV5443dc0sahq0tygtQJJGKm0i8qLGeNq/wAK4IwPpUkLnT3RZosNghZOMNkY4I6{+}n9a2by7W4KPbRvcSeWu9pWAO7ADYxwBxx3wBUSrNoqFGzOGn0hlY5yWz{+}tdV4TUtdRoRtKqSc8DNOl0u4SD7RPHsLDIXPGO1aWgad9nX7S/zhiBjPU{+}1c86l0dNOHLI7Ky00SxbNm4{+}4xg1mat4ZSdTlACBx9a6HTf8AXRI4BYjGAeV{+}o7V1o0hbi3Bdd2eTjt71x89md1kzw6fwgyycJx64q9YeEhwNvI6{+}1es3HhoblULkHrkdantNBVc/uyQeBgYqvbEqkjjNG8IRyEELyOORXc6L4SRCpEY{+}XjI9a2tM0kRhTt244JrrNNsFX5cgtxnIrCVVm8YJEmhaEkW07MkcHFd7pdmI4/u4I/L{+}VZGmWojcHcPpXRWxYBSBgdCf/rVySd2dCjZGjbOseMY554FWXm4AIGOwz1qquQoBBxnrTJJwHUdugFIixM3zy/e3fTvU5g/dbsDnoAKrQR7nzzz1ArTRCIwD1xjFVEiTsYd1CBhug6ZrwT9pLUfs3hzUIIEE80kQjMaAhhESDKVODk7QR8oyNw56ivovVIkt7SWVwFRULFjxgDk14B49nvBJfnylaxtbIz30ruquvmg4iATIO5Gwck4wM9SKdKL9onYVSa9mfDPinwFqdx/aV/DbyQW2nObZ0YHfu6tngdD9K80ubW2a5kZZGgYkgpJ0P/1q{+}6PGWhHStSubm{+}trbw3p{+}riFmkkMk1pAqx9Aq7mdnx3AbJzk8187fEX4N3GmFdatrNp7W5AlhC/dKlQcn86{+}spV0rXeh87VoOXTVHic1s0YAK7kOcEdKLdZbMq5Yknn5O1dNDp5t4LgXEL7TCzRn0YAj07YpllBE13CmxSqgHnv0P{+}Feh7TSxwOm0zqPC9q9/BtEpSSXBkJJ4HHUf56movHlrd6Sn2doWht2XcOPzz{+}X616jo3wd1a90a38Q6EjXtuFKyiM/KpH16dB16VzfxZh1vVrCxtdU0ldKKBk80QtE746A5{+}VsdMj{+}lc8YpvmNJyaXKzxbSr5Zd9vJgJKSAx7dMVsLYBmhgIOyM{+}oyT2xWfa6H5d6iEqGGSUJ44xnn{+}VbSqthOJ52ZWKhtipk4I6//AF66mktjli21qZMDROLl/KLBR1P1yF/HH61p6cJtO0{+}SCQ8TAbXByBIpDAH8f5mpQwjtGeC3eC2JMhZiQ7dhz{+}tbHgTTYfEel3ujsxjlQia2nznYwyBn1XJ59s1Ll7txpe9ZH2F8F/GtvfeC/hp47mYodEvh4e1U9hBMPLRm/wBlSUPPvX37baKix5T0GPWvys/ZU1mJfEuufDbxCxh0rxXbNYNEWKtFdqSUIPQHlwPXI9K/Rz9mjxrc{+}LPh1Hp2rXKy{+}JvDkp0fVfVpYxhZcekibXz7n0pRCZ1mv2zeRbwNyZJ42wP9lgQfzxWiFlHGST1qzPYtPcq8gBCnILdQP/14/Kri25d/T61aWpk7aFBJZonHJPtVmK9ZTlj{+}Bqc2pCg8mhrJWA4GcUyeVCrqSnggg1KLtSm4EfSq5scDpj1xUbW7IuyPIGejDiglx7FxLnOM4zVjcGXNZKLKD2wO1SpdGEjIGB1wKCLGhj2J{+}lFZgv1X7zc{+}5ooCzNkcUZoAoyPxoJFB5p1NHWnUAFLjPNNJxTloACueaXGKWigBAMUtFFABRRRQAqnBqrrukQa7pFxZzKGDr8vbDDkHPbkVZp{+}7AoGj5y/aK{+}DcPxO8B3ojjC61p9sSJNpV96glNp7fMATjqARX5IfEHw4z{+}RqEMUkFy7SLcRFQqxyhjvQcnOOK/fO806G6WcMABJGUbjqDX5S/tmfBS68AfETWrm3hH2a/kOoWxjHyybsBwFzwwYkkeh{+}lY25GmjqhLmTTPjXTYvtCzwlcSoC6KxwDjqv17j6GrcZJADExshOxs4Kn/wDXxS3U6yXEd5G/lyjKtFx{+}v19aW{+}gW9VrmEMsjclQvHv79f51ruWmWJ5LbXiI76Jo7ngecv8XHHTr9awr3RrjSWJB8yAnhgMEe47ZqZ51ugrcxv2LcDPtVyzvriWMxH96ucFB6e47003FeQnFPUyTqkksflXREsfTe3X61E0nksNrdOQrVs3tjC3zLF5LdMYK59fasdrXHCZ442sOD7UJxewveTsy1Y6u1q{+}YgyZI3ID8prstK8TuGDxzFcjDcZOfcd/rXCxRQGRcxqG/uu20H2z0rTbWI7SI20FiFlPd{+}RnHauerSUtkdNKu4bnp1n43shBtluFkyQGVCF59SCcV2mlRQ6np9rcWl/bM0zMoiV8uoHcjoAc8c54NfMMkz3N5hE813wgwOWJ9MV7z4P8Ea/wCFPFY0e8mjkis3CO8UZKqcsCu713Iy88kggdK4a2GhTjvqdtDFTnK1tDu5tEvAFjklyNvU8jrW5oHhy7vwA0wZFHy{+}Uo/yK62x8NtLZC4MexSPuv19jWp4Z0hrW5KKBktx6V4bkrWPdtYxLGwbT3VVhKrkAny{+}frnvXoWgWYljKhcZHTbgZqe70Xem4x89xjHNaPhmzZZgG9SMHiuaUuxvFXHT6OyrlgMY6YquLAr80Yzu6gGvSRoIliDMmEwOlZOoaGsSkIBx{+}JqbjS1OZtLYIVzk{+}vtW5aKgO0DAbp61WNm8bYI6c59KmswxfDHPpUN3NIo6KzXAVg3PtWrFIEQE4PrzWRakYDbeo{+}lakBViAV28Z4rOxqXo7h3Y7c/405iZGUkc9OT0qa2tvMUFRkAVYS0yACe9Ul0MXoOs4c4Ax6nNasDDoR2/KoLW1CMAAx78VpxWzMSFTBPc8mtlE5JyRy/jmR20hLSEhZ7txGCWx8v3m7HsD2/LrXzb43s7IanbQXV3cLby3E0TzTZaF5cRPG7kbss4WROCpA5Oe/uvxUvJ9TjOlaZIpmgVp7qQKGO1cnajg/KcqAc44YeleG6rYqPDV9eW8smqRy2a/wBoJcDb5flzBXjQkdkV{+}f8Aa962/htJ/MmKdRG3Nq1/rs8M9xq9pBqunXhsV0DRis/nxzArGzy4ZU42/MrZGWxzjHo{+}rfCa0l8LWenzwRSGOEIy4AUNjnFYHgLwFYaHB4fSKxSEu0sz/ZbtYIyPtKsoYY3SBc8D6V7hqcIWPIUg9ga1qzuouOiMorlk0z4O{+}LHwDt4r4S2tv5EEKbTGkXyP1ySeoPXmvHvA3w1eP4iWlhJb7kcnyy3Ibjp7V{+}gHjjQVvi7hFckEAnqK4XwV8Ho7bXkuLmBBLu3oyD7jfjXqUZSmkzkrKMGz3P4RfBfQPDfg{+}0SKwjEnlp5gZeQ4wGIHbpXmn7THwl0Txn4T1fUZIcRaejLb3EK5kZlBLNkn5gTtXPP3eK{+}ltDid9Jjj{+}ZCECyBTx06564rivj3bwxfDvUYtgQSRiBEXGSzEKgA9ctmvW5bRujw1Juep{+}RHiTwxovh/VUkk1BHWMCSWIoUkB7qFPfpyOOtcRPFc6/wCJnuNhET/NGg5AXtz65r7Q{+}P8A8NdJlWBmtIvtKJ5fmAcjHv3NeEaT4Li0{+}4RLeMkn5ck8n/OaFNJeZbg5PyMvSNKsryLyb2CQI0ZQk84Pr0qCx8Ea14W1xJLG2e5hkLLEyf8ALRO657HFe{+}{+}FfhC13b/NE3mMMgkdPWuw0TwPqGmzC2yDuIMbFenPbNQr7F2W589/ELTLvT5NH8daXA9vPaTxRX8RUq0Vwpykhx0OQFP4etfYHgj402Wl6zpXxf0B2v8AQtUhi07xTp1t8z27D5lmZM53xFmGf4kBx2rpo/hBZ{+}L/AAzqFrf2qTR3sXlTrjG/jg/UdQexFfGj{+}AvFnwP{+}KtxoelanLpc84KxXLFVgu4zkqX3/ACYI{+}U7uA2a0Whk9T9edOvLfX7K31CymjubW6RZYZoTuSRCMqQe4IrUS1Veq5OOtfnf8B/2mfG/wPZ/Dvirwbfy{+}HjM0kT2sTMtopPPljP3M5bapOM8Y4FfZnwy{+}PWlfFKUxaNFJMQB5kjo0Sr68OAeOe1aqxzShJbHozWgbiovsu3P8Xbp0rW2YfJOSo5oEQfA4GfSquYqTMz7GccgGmtbgZzwc1rCPjGM0ksIOPl6D1707j5mYE1viUKMfnVSa0bkg/L04rZuoWwpxye{+}KZLbMdqj8aVhprqczNZEv6fWituayLyE7CfoKKLD{+}ZZU0Z4NFKOlIxG455p9FFABQOtGaU880AOoxSA5paAAjFKBmkpcfrQAmMGgHFIWC9SPxqnearBaDLMPegC6TUU93HbJl3Cj61xus/ECG1jYQqWOOMCvGPHvxL164jmjt38iPHBUfMKTZooNnuHib4jaR4ftJZbq8jjVB3YZ96{+}Hf2uPinpHxLsha2MbSXlsfNt5zxtYZ4B9xj8cVyPj7xFqs5keeaa5fJBZnOBXimu63Pbv5sknfg7uTyDWDcmzeCUTwPxFB5OpTr1DncCBgHj07Gm2F68W1Ap3JwDnqDXQeLbNLu5aeJfmDHO0c1lGzBjjLrgOuN4HQ9s10XutSk{+}pWlZJUeWIFg5IeE4BDeoqjHZSMcsSMcq3Un61pWtnPDMS8fmGQYdc/eHqPetqXTBFZj5Vki6/MOVHoT/WpcuUpR5jlJtTvI12{+}YzJnAbqagu7{+}TykRkQPyfM28t9f8au3NmzSnyxhACSx5FZZKuQDjA4q1Z9DOV1pcrTTvdYDct6gVu{+}E7gWOoGa6i{+}0xiCVIoic/vWQqjY77SQ31UVnrCLgLFAP3jdTnoKtQ6lf6M6waa5s5FbbJdR8OzHqN/UL7Dr3zRKzXKhRTj7zOv8P8Ah1YNbsJP7NvLWfTWTUbl508shAQYguf4pHKqPqCOK{+}hfhZqV1rXge{+}mvUF14i015FW5Kl33u7SqxyCu7llBbkADb3ryT4f8AgnWPGUluLXU7m2EOHl1OcFzcSjO1VDdUTJxnqSTxwK9n{+}HWmXPg/xTcwajFEbme0K3E0W3y3xIuyXafuqd{+}CRkhgAOCK8LE1dbJ6o9zD0Wld9T2bQbu213wzDdQgDKfMmDgMByOQM4ORnGPlq14Yt4/PfeOjbeO1cZDOvg3xa9scR6Xq/wA0LsQqxzdcE{+}{+}MDJ712GlXaCcOjghueO9eHNfaWzPcgvsvoejR6dHdQDb8zHjkVlvpjaZfhlXCE9K1NBdpgOcD/wCtXQT6YJELHnisHqdEXys3dCiW5sVP3uOQeAahvNKRgw2YPrWl4HjG427A5PAyetdDdabGZdu0/XFUldHPOpyTszyy{+}0FuqD2yDWfFpGwknKHdwMda9Q1DREUFVTCnv71gXmkAAsSMHr7UnFpam0Kqlsc1DAUOQMLV8IxEe32zge1WUsRFIM5Unpmp7ez3FMHDKeMDis7G/NbU39FtP9EX5OT6Vfi01m2tjg45FT6THHBaL5hyCM4NSy6iUYEYUKcdK2SUVdnBJym3YSLTm8wkZOfU9KzvGmtjwnoTzRoJLqTKRKjDcT0wBnrycfSrUerASHlQMZ5OM{+}1eW{+}KvEC{+}J7m61SQtLpVuTBZRjBExyd0v8PY7V/PnNbxnGMXL{+}v6Rj7OUp8r26leB7jTNGuri7uDDdarIbRJ7gkCIeW0kkuD06L2xnPrWYJFufhb4j167KanG8D28STKEDKWw7HZjqxY8Y4x0p/jpodP0Lw2rXSwG1dppEeEurhyu8YxgcbupHAbHIALLC1tz8LNTsXGYLmOQorNtynPPtXLzaq52xhuxfBHitNQ0zwtLOtmky/abXYLdn4jnCrjrt4GdxP9417BqN/EY8segx1r5/8HzTweGtCMsC7bYyLcGG8EarMHiBJCg{+}YS3OOmeSepr0m71mfUAVhQgZ5kLbR{+}FbtvlivL9SOROcmjbgxqt{+}beBSXI/1in7v416X4f8AA0FtptuGjLyKCS7gZJrzHwtZ3CyxgzSSHeHcq{+}R9MV9GaNYNHZQ7sqdoyK{+}lwSThZnzOPk4SSRgmSPSIvuStjkJDCzc/lXH{+}KtPn1uWO71JRaabbN5sdsSC0rDozDtg4Prx6Zz6fqlkJMgjH06VwfizSfMtpF3MVI6E16NjzIyvqfFnxptTqerzCPIhBOOM4FedeBfB32/xFF5se{+}PfgqB/WvoX4heFiZ5AFG3JJJqh8JvBL/wBsPPJGwGMKT0xWDjqdaloep/Dv4eQiBcpwR8oPFeif8KksJ9jvAC6D1PXj/CtLwbpyW0cYKlSa7rA2EL0x1roS01OOU2nocfovguHT4DGEzzwTXmvxh{+}D{+}g{+}LfJOs6bFdRQPvV2QErxgr67TnkD2NfQkEQ8o5XrzzWZrGix6lbyI/8qLIhVGndnmHgbw7pej6FZ6bas11BbJiI3LGRlX{+}6SeSOOM16D4as7O3hVUtkifOdygCsWz8Ff2dJIYs7T0UHit2zsZbadF6qo60IbldG6xBVhzycdaeybAwHTGM4otoSAu4DOM1OE3jk9TTOcjWEhyRjjkj0odWKHoC3TFSRRkZw2QTiknYq4IHCjtQBUmj3cMvTuKiW3O3J4PvVwnzCqdSeTn0omjAODgYPQCgDOeXyzjaCOo70VZaMZyV69BjoKKdx3Zj08dKbtpehpCFooowKAFBxQRk8UmKcDmgAzjrQKZI6xgliMe9YOqeKIrYMsTbm9qTdhpN7G7NdxwruchRWLfeKIYDtjG41ytzqVxqJ5Y4zzUsFhI3LZI96V2zRRS1LNzrV7eONuVXNVmtZLvPmSFiK07axKoTjnpj0qzbaeVVgeWPUU0u5aOVutGEq8gKK43xH4XWWGQLHtjbqQOTXrc{+}mmVNvQnoKztT0PfCyeuaLIZ8Z/EPwOXhmIUnJPHY18zePvCN3byvtj2queMdM1{+}lms{+}BItQtpxtOV7EV4V41{+}FaS5LQjHbC/WiyKTPz8udLmhdiYwSeoPeqLWIVGBjA5xgHpX1vqPwCa{+}unZFKc9QM1Un/ZjkVA8cTs4/X/PFMdj5e0bSnuLlFaM7dwwT0Fdd4n8MCxsoWgCB2TdIXGQnp{+}desf8ACidT0m8VzaMoBz0Nc98QvDM6XluhDQAIAynAz74rmnK2rOulFvRHzf4iTbJ5UaDys/O6LjcenbpXKyW7K{+}1fmHUMO9fTVj4Hg1I{+}UIhIG4LqvU8/nXQWnwFtLyIE2qtxwwXFc7x0YaNHSsvnPW58m2W/TbyCd1OwN82BnjvXungT4a3PiaRNS0{+}W11O3cAlomBIOOjKOQfYiu8n/AGcocFfJwuM4IzWVcfs{+}rosMt7ZS3enyxoW32kzRHge1cdfFRrfA{+}VnZRwc6PxK6O60AwaFazuitJLaMYZoMFAsg/h5Xk/TI5HNYuja/f6x4k168bUog9lp8jm1IMZTJUCNMkHdlQwXnlN5{+}5VPw9od74ii0/Sb6{+}1f7Olu05VNRkAulZyVaUKRubkjJycYGcAAdzoPg2z0u/Flb2KQE2U4hjhT52PylvbHGWY88ADlq86nZTs3dnfUUpU02tjWvSvjzwXDG7FJnjDK6lgyOOhBPPB4ye4rU{+}H2qS6spivmMWp27eXcRjHzMAPnALEkNnPQDnA6VT8L6LNpGoyaQdqRyMZYF{+}6TuGT7nG36AY9a7K18OXHh7VIdaijd4lAjvo0H34e5x3K9R{+}NcqkneHc6nCSSkuh6R4fjZNo6YHcV3NlArpjBJbvWJpGnxQwoy5fKhlYqQGBGQefUYNdFZRBXTGVXqQf881zP3XZnRo1eJoaPD9lvI8KSOufT8a6d5AJCXXOT155rDtoxJKMHPv04{+}lX7qVtoJUEj9a0T5Vc4Jx55G9awJdR7T8wAxzWDrekxRqzKBt54pbfUiFPBX2HaqWo37KjDacHOea2lVTiZ06U4yumc9cKke9gc4J57AVShmDTgnhV5Ip2pXBYbvlQYxt7VhyX0jS4HIJ9c8VwuSPWUOh2Y1RrjaFPy47U9/MWMnqPas7w6pnCk8Y6mtzViLS3VEhFxNK2yOIhtkhBGVLL9w7ckH2Na01KqzCrKNLocxrTXF5G2m2ryRvMMXEsUjKyRnho2XHVuxB6E{+}ormzbRXWpPZwRAW1oFXCBCAw5K9CRxs7jj1r0tdIh0yxmmuJC3loXlmmfkgDuT6AY57AVxOkSm30ae/uN{+}zDzuW6oOWI6Ae3TsOtVWWqitlsKjNNNv1Z5D4wF3r/im8SC8Ki3lSFrZBlsEFWJ9Fzke{+}a6/wAXaZdXHhDU7WIsZ2s5I1UDuUOKteFWGu{+}M9RhW1mtbUn7SgKGIXC4XY74ciQg{+}ZgkDA2jsa9D1XTIbTRrm5kUiKONiQBntWfI27I1VWMY3Z4Z8KvDlxp8ccEsdzDDdyu8cYCPsVli3b2/hG5SMLyckZGCK9403wvHtDbS2APcVxXglIpfEui209rPbXCWvAlZd8Rd2LIwAGSdgGRwNp4yePZJwliqZIK9zjoa9H2d1G/Y82dflcrdSvpdlaRSRKYsOWAynNey2spS2jU4wFABrynQojf61ZKgG3eCSox/WvWEiZAATx24r6LBq0D5nGyvNIilG4EnkemKwdYtRcKVAGB6jrXTiMAFTzUEtmsgOBj3Nd55y3PGde8CR6tKzMm7nOcfzqx4b8DLYKxEQQA4GFxXqTaYibtwAI/Wli035QFGR1oNfaWVilpGniKNc88VtLAFUjOBiljtDDjgYqdUy3T/69Bk3cVEGOO3amTAyZGeM1MyjacDv1pgG05PHvQSyuYjhvkOKkS33OcqO2PepmO9MZOM05FwOhJ9M9aA2EYhcdsCl25VcdO9CjfuJGRjvT2YAYVeMfnQK5EqlHXjIJ7nvTJQXHI5JzgVIzgLu/i7H2pCdqnJBwOn8qAsRFjGu/G7jAxUbIeOSQepPWpXdeByT1xTC2yItxycAj07UANkBdyc/Wil83A6qP60UCMIjNA/WlHSmkc0AOpRxTVORSjrQAvTOaq3uoR2MZZ2AxRf3qWUDSOQMe9eaa74jkvZyindk8AUm7FpXNnV/Es145jiO1PY1QstNluW3OWwTnB70mlacWG5/v9ea6S3gEajGQcYzRbW5o9NiC10lQQpxnqRWlFarlQVzk9PaiNG80BRkhfyq/ANzDjJUHFMW5XWMlsKuAanhUAqgBz16VMsPzKAcYGc1Oi7NpHBAz0oBuxUNvtZeCRg9aqTWnz/dzmtpE2ur5PTv0pHVJI1kHIzj3oFzGEulJJbS/Lhj3ri/EfheO4jZfKHK8MB9a9RijWWNicE9qyrvTxLGwIGRx0oLTPFYfBxNz8kecYJBrsdO8GQ7CPKVxwdpHArpk0EKyuBkn2ratbHylwqgKy8YNBXMcNL8P7G6f5oFA57ZxXzv{+}0Z8Coje2F/bw4QqwbLBV6569utfZkVh5WXIGMZwe9edftD{+}HTq/gpfIUho5Q{+}VAJHHvWFaN4M2w9RxqJHw7ofg0aRdi3liwpPD4yPwr13QPC9rLGjgKB228ZNc3d2cuI1QSbkUB{+}ACp9x6c1taHrn2NfKkbBBxtr4{+}rK02j7amrwTOmfwbbSkEqAD0FNn{+}GFvqVtJEyAoQQ2R2xWvpOridFYkKVroNP1IbsklhjJwKzXKDckfNPwi8AWd7r{+}oz/AGeCOdVjt0EKbMxKCEyuwAtgDcwJ9D6nv/EnguHRdUsHkjA80Oi71UpnIPIxnOM429D8xztrU{+}Hs6p40trBNKubewsobpmvZhgOzT8ADGCmQ{+}Dk8gggYrtPi5crFpel3Nu7pLHdhSI3xlSrZz6gjjH{+}1W1L{+}JdmE52hyr{+}tTzzXPCqadZWOtx7g1tJtkCqxDRtwQQvLYO0hc4yBniu20jSEubfmPORk7xnP1rq4tCttQ0OfTpvm86MoQvA{+}Ydj{+}Nc14BMsWmyWU{+}1Z9PkazcKoCnbjBAA6Yx{+}Vcji002dftFJNCaHYjRmbSpGysWZLUY5MW75l4XqpYckng{+}3PRWsBmkGQFTI9ttO1Ww8{+}CKWPC3EEnmRPnoe4/EZ{+}nB7Va065S9ihmhVYxLyUzyh7r1OcHIPuOuQa2nHmXN16nPGfLePQ1VtViiyH4POcdfwqKa5IJ5LE9AB0{+}ta8ESyQYJ5x2HBrNv7UAEAAjsQOnrxUTi0tCISUnY524uZhJlMqc/wj0qCWZ2VmYc1dmtsEsRge1VEtv3mT8ykdx1rkaZ6ScUtDLv7aS5jBj{+}UD{+}HHWsS2sne5wdwGepHeu2uIf3HTBxjpiseOECdsnH49afINVNTodCijsrCSRsEINxJ4YnsB7mptCQ3yz6pOAXuci3DAqRFnKEqTw2CBx2Ue9YUCnXLxdPRithCN1yrLw/IKgHOQSR/3zn1FdjbzR2yTSSTQA24WSeN3AZYyeXGeuAGOBXox/dx5ep5NaWrm2YPj64a38PizWSWGa/nS2RlJ3c8sRjk/KDVPWNEli8Ialb2iK87Wzxjedu7K{+}pOBn1J4rf1XVfCetzI51WdJNMjkuzDLbtGsqIpLOu5RvwAB8pIyR9a{+}X/H/jzxL4k1WJmS5SykkZbPSLa3NykoX7{+}YhxIV4yzfKDxx0Oc5LmTIoSc4OMV63PbfhlCdU26pcafJZzxWaWR{+}UiFtrHbsJGW4JJPTkYz1rpfGVvPf{+}HriztAhlmwig9M7hxXyDo3xj8b{+}Evi9p8l/cXUtvdyxQXFrJb/Z0khJCfNEOA4HII5z3xkV9YfEDxRN4b8Pi9sIjLes6iACPfltpIO3vyBWkNUrjmmpcrRgeEpjd{+}OZbwRSw7UWENLglyAWbKjhclt49nxngivSrtzKxXHHoK86{+}GeL/Wbu/C4S5J2spJDCMLGQRuIBVgw6AkdiMGvTJYspg/LjqR711q99Tkk0nobHw9tCNVkIA2Khz2Oe1emHkD2rhvh4ARPKA2CduSfSu7jAA68CvpKCtTR83iHzVGIi5OTQIjuPGB2p4cfSgyAdDW5zMjnt1IHAzSxphh7CnsyuQM4p6YHQcUAyKYjIwe9BXC596eyqxGfyqJ{+}w5I96ADJC9TyetBjyRUYLHA4FWI48EEjI/lQDDOAvueooBz370EBVGcdeQO9KCHz35oJDrwOhPNAULHwe/pQY8A9yTmjBRTx1oArSnjApYwGQ7xnPGfSpAVJ5HbrTPu4AwQRigaIZAcFlUnPTdxRLIqqMHgdvSrDgeWHI{+}6OprOnkymSRk88UAWFkDDJAP1FFQROSnJB/GigRlHpxS9KKM4oAKjmmESknjFOcnBrnPE{+}rCyt2BYBcdKBpXOc8a{+}JVkYWytjJx161kaTpcrp5r{+}oGTWZGr6pdSzv8Adz0I{+}7Xa6dZ{+}dEfL6DB9s0G3kaNjaDY79hj8KvlGcnaDsIzU1tb7zsHygpnA74q9HCHkXHKle1AhtnAEkJ4yVqzHbjnHp{+}dLGnKt0yMA9RT1yWj4{+}ooFcjVNpQsMDHB/pUyuDt/vY5FSNEHVQMHnFCQBio2leO9AmJIFZE75H51HLGwtQOuG5qxtUMv{+}yOlTXEYeMKF69{+}lAitbRmPhj1HSkmhQsFx97g4qaJVaQAnpxUiqEfJHPrQHUqG1AVcYODU0aAAexxTpE4JHHOadwR8pHrQN6itEdvB56Vg{+}M7dbzQ7uJh5hEZIjx1xXQRsGJO4cGmTxRSo5bGfpSkrpoqGjTPjq70e3W9uAiN5ZbcBu4{+}lYGpaQ8D71TaPQHmvUPE{+}nrY6xdRoqRhZG5XvWFc2UchGQfUAd6{+}LxCtUaPusPK9NHIWF1LbuV64Az7V0llqbpGeN64zUTaGI5GmAwpGcCpLC12lgo3Ejn0rjOo4XRNWu7f4htf{+}Uk1gyywBnOHgnMmSV/2TGgGD6{+}5r0jxwwv/AA60YQzEOj{+}WrmNmwQcB/wCH615xf6XqV34o/sr54bOO6FxbsmxA00qlVDMWGU{+}UkjnGOOpz6B4ntN3hmYkK{+}xFb5n2hgMHg8Y{+}taq6lFmKs{+}aJ2GiahI9lA0vySCNVYMwODjkZ7/XvXPzzrovxFWeSTZb6xCIiCOsqkAc9uCB{+}JqTwlczXWiRmTyhJz8kEnmInJGFbPIHTNVPHlpIdAkuYWZbqzImjYZyMZBIxz0JPrwMUVPiaCCukzt9244xlc5PNY9/J/YWowXv8Ay7Tfu39VY9Mnsp56fxAf3jUnhTXF1vRLS7VQrOvzgHgN0Iz3GQefatC9to9Rs57W4RZIpUKMp96cZcrTIkubQ63TblZLJZEOFZdw4qvfZlYke{+}RntXBeBNfuNK1G78Pak7STod8M7YAlRj8pyWJZm{+}cnAAXb24ruJ5Sp55Hfmt5WtpsYRTjLUoXWCjBe3JzUEUXlYf04xirM8nlfMeSRiqryEMSRgA8DvXM4nWm0h10fkwO/rXF{+}INTFhMqRfNLM6xxptzvc5ITPUZAPPQck9K62SZCrM7YHJ2lwuT6ZPeuR8NaefEGuTeIDh7dsxWRZdjGLJIdh0zzj6D3q4xSXMxX1sdl4W0VdI02KEfvJG{+}eRj/Ef8P8ACppfGGm6Zc6vFrctvNLawpJDELXbMke4FQZD94M5wB6ge9aFofKQcc5HT/PvXlP7Sss1rZ{+}G2tiITcTuZJlXL4j27R7jLk49QK1S0uedXfPa5yfjn9oyOe336wsNlYNNugikgEgRcHDgdWI5yelcVL8ZbSHULuS31y6tLi0iwjm0KRhCwJxs3ZBbJwR/EKr/AAg/Z21T40Wd/wCIvGt08Gm3hYWtnCoR/J6JluvAA4BHSuT8cfAH4neA7mfTNE0XT/E{+}k5P2e8kZ0lCk5AcKy57flQ6bb52zphThrFS1N27{+}IC{+}PfiDo/ibW4IzZ6QggtoNnlte3eWaIFfUttJA6KhPevqLXbbHh/Tp7uOaU2qfaZI4TgybIywXjGMkAH2Jr460T9nHx/Dpx8U{+}LY1a{+}06SGWytbZziCN5FjdVjHGcvGQepw2T0r6k8Qa9e31joCxSy{+}a0DPcLC{+}FcKnz7vVcZ/HFPSFRRve5DV3ZdCx8Hgwa8depzJKMEAO7E8cnsBn39K9MklVwEX5ievt9a82{+}EsMun6NcmYjzgVRwG3DcBkkHJ4yT3P612lvqsbzhG3Ek8LjIz/n{+}VddNc07M56isrnp/hBEstNATkO26urhfcOOlZWiaasdnAAuBtB5rYVVQYA/Kvp4q0Uj5OT5pNi4Ap2A3Q/Wjdg7SM{+}9CkE8d6okjDFGIp6ytt4HFNI4J9{+}lKCdoHbrigB7rgZpincB/OpjjFQg7myePYUCF2gEccd81JnGBnkDNREnd17U4DdgjpQFh23eAMdf0qPlWCgYHXHpUw4x60gByGyOn50CGoxAHY5/KnF1XqceoNGeQOOe/r71G6k9RkjgH3oBDC4AJB6etPGCASOfWoHYE7fUVIu4KOcgdvagZFdBvIYKc5NY7z5kCZ6HFaN/PvYKTjHOKwrd994Ox3d6Ckrm0iAqDkZPPNFXFcIMBOKKDMwOe3NLg0Y49DQeB1oAhmJ2kA15d8S77bIkJYqGznB5xXpN5IUwf5V4r441aC/190JB8sbQP/r0FxVy7oExaM7uEZBgGu50p1jEewYUjBrgNLkeRFVTgL/D7V22k58kE7gM5AoLa7nTWAEOJM5bOCB1Aq8mVUENj5sZHWqFqWUscfMeQTV{+}NGVT/ABdOaAslqWlKlBzznqKeEG7B4APr0pIkBRuODggmph0Y7ffFBA3cAcZ79fX2qRcnjIBBPJFMeH/Z9/TNPJwQMYFArjfLLFSeCO3tUysGGMHI7moyfmGDz0pzZKcc59KAEJAkJI4zUsmF5U5Wq6gsAT27GrAAManFAxDtlC4z05pgVVAB65waQRlFLcjP5U1X3N3I65oAcLcnA6AnOM0NFgHnIIqcBjgAg96jlzGp/iPagDxL4iaV9n12QouWkAZhnr{+}H51yJtAC3GPQH1r1v4oaYZLNLtIczA7N2ePxrzBSxOw5Hsa{+}Vx8OWpfufX4Cpz0kZrwFN{+}eVPGP8ACkt7LyZNu0nNajxgfgen9aQgTYCrtIOeh5rzLHq3PJ/ilY3OoX26zuWsJtN8q9E8aFmIBdT095FH416Ett9u8GlZXG8wFXccYIz/ACrnPiRE2nt/a6zrarFbSLkRLJlgpYAhvlx8pzkHjtW/4NMV34LiijeN41R4tyPvXAJ79{+}vWrstDJPXQqfDK4iu9DxBJBIIXKu0AwpJAb8{+}fU565Oa6m9hWWB1PzBlwVPpXFfCGIxLqtqzAlJAVQQ7GAJI5P8XQDPtjmu/ubb5SBgH6VVVe{+}7CpytE4T4cyy6RqeraO6YFuwdGAzkHv{+}Ix9SGPXNd8WVj94njpXnnjOA{+}H9Z07xBEmPKdYbgKoYlDx378kdR1HpXodo8V3axTRNvikQMjr0IIyD{+}VZJGjetzm/HenzGyXVbA/wDEx0/LoF6yRkYePn{+}8uRn1rc8I{+}KovFGiW98pDiVATgEAHGGHIHQ5H4VZbGzHBzkc{+}leVTWkvw58bG8hJXQ9RlHmKBkQyHqRyAM1pBpaMHDmWm/wDWh6/PIxbjgDsOarNuIJGRS2dzHdWqzRSK6t0KEMPfkcHnI4rJ8a66NE0lIYIhdaleMEtrckrv5wSCDkYyD6YHvSSu7MzT1VjL1i{+}PiTUV0C2bMLoJNQdCGXysgrGQRlXLKf8AgJ967jTrKOCOJY0AjUbQMcAdqwPCnh/{+}ybI{+}fIbq9uGM11cSHLSSn7xJ9PSusjGyMjk/Tn8qLtlTtH3YliLBUYGDnNeK/HrWo/EN9H4fgV01XSF{+}2q20jfBIMMV9drJHn/f{+}texK4ijd3fAUE5z0FfLHxZ8K{+}I7y0f4iaE0n9qxXb3EcW0MHgxtKEdSrJgFT71blpZGMaSqOzfp6n0h8Mte0zUfA2my6aUEKx{+}XJEnBikHDIR2wfzGDW5JMsxDMNx4wM8mviDwB8S7bXru3/AOEZ8Qr4R16eRY5tLvmGwy8A8P8ALIpPQE5Fdd4q/aP8Y6PC1qp06CWFVjnufsjIyygYkXazEDa25fwz3p6tWRk4Om2prU99{+}IPiuPS1sNJjlRbrUZlyrchERlfJxyBlV/WuYuBNo{+}jNNC9v5xU2IlmIASJo23v6qpKpyvODjvivNfgxp2ueO9VuPE2qTXGyRPL{+}1NkBxkFvL9c4AyOAM8k8D0jxUhWe4hQx2lu9tDZRyOhIeSSRm8sfKwJKwkc8VlSi5Vky7KMH5m/4QsVtvDNsI2BHJ6bcDOAMdsY6dvwrtvCWmS6hrMWV/chgTzn8K5nRd{+}m6TbQzAOQgBwck/UjrXpHwztnur55mUiJBgH1J9q9rBwUqnMcOMqONPQ9QtU2wgDqMVL9xCO9KpCoNvApjHPfk19CfKoY2QM9c0{+}InJJ/CnLgLjvSp93FAC4DEA9qbsIJIFKxwSBT8gLjGBmgTGBuR6Un3TtHHv6UrEA8cYpqY5Lcn{+}dAWHIArEEdqcxXJJPIHbvTQSW9AaSRh2wMCgRIhVgTzk9P8KDtAHp/nilQAA9uOlNK5I9KBDZe2OMf5xSs3yjJ596exxgkVWuGBQ9aBkIcGRiMFelBmCMVbIUjt6UkMQbd6noTTHDebtGDg/eNAFK/laLzSCNm3HPU8isrTRmdSzA4Oc1Jrk37uZcndlRx9ai0yEs4L8oD0oNFojpon3oMLkDgEminW6lohhtvtiigyMMHNIRyaXocdqbIcAkGgDL1yUW1hNK33VUk5r5uub5NV1CeYuRGHIPUYr3jx5dvB4fvGBwojYmvnHSYGiWaaL51c/Mr9D7ig1gdno{+}peQUDNzjZz39K9L8Mv9rhiywLbSDk{+}leDT6nNbQSYBLRkED2r134eagNSt7eRXXBfDAHpxVWLPSoINwUk7iRgVpWql02HgY/GqFgsikDPAOOelayIY41PcHGKkhtWJY14BB4x/kU7BIAzye1MUbe/U4pCwUKV/KgzJskDa2DkdaaMHpyTzQrbwex7j1pqAq{+}OueOlAIZcDZt4ozsfkHHr7Us0m{+}QbjTZPnGQe350FEjsNhAJIHPSnQs0seBwq1TZnV9vY9xVi3k2nb2I6igCdX3jaR7fhmgREHcRlW4pq/KTg08uyYUgkYzxQA6AY6njpT88/hUY2nIycdaGHlnjJGetAHOeO7RdQ0GeMnA4YkdsV43PbeUThsg9{+}v4V7rqsH2uzkjZSVIIrw6/j8i7kDIFCsQMPkn8O1ePmELpM93LZfZIeCm0c4PPekUDOAAfpTDIWBHA54wO1PB3McZzivmGz6dI4r4vRBNGtZvKS4jExWSCRsI4ZWUgn6Metafwxs4LXw/BZIqB4IwZGgkWSFmbLZjKk5XBAGTk4/AafidrddGklvIo54oWWQh1BBIYYBBrkfg1catDd61Y6vEYbiKdjE{+}4FWjZ32YA4XhTwMD2rRHPN9jR8FEad4r1iwElwVlzIqMuYlIP8J7HB5GPTk8Ad7lmwMZwMcda4K3jaz{+}Jz/u5mDwtukik{+}QA4wHQ59PvYGOmTnj0TysDB7elbTW3oQnuZOtaOuqafPbMVXep2yEfdPY/gefwrE{+}HurfaNNuNOlHl3Fi5jKsV{+}7k46dshh0Fday7VQE9Bn61wur2h8LeNLbVQzLZX37mdEVdu84AYk4I7dPSsZdzaNtUdi4DIeMc{+}maxvEWiW2u6VcWd0A8UoIJP8Pv8AyrcVPmyzZGM81g{+}Jvs5hgt7u4a3tJnIlkQqp8sKS3LegBPGTx0qoQdSaiuonVVOLm{+}hwfw3{+}J9np0OuafqGqWOqWmhtvkvI9QR0jUvtO9txC8kZXrn6113g63m8RTr4ovsP9qjDWiqVKbCBmQY4ywUcj2rT8F/CX4ZaXo3iCx0zQtCubOUeRqMbvHcqWJUKjZ4DlkBHGTuY9Rirmn{+}H7HwTLaaFpSQ2{+}kwxeXb26A/uVyWVVzyQFIHPp04r1sXg/Z01KLvbc8zC4znqSi1a{+}xsRMFZRjBPerakgVFDbJwNwOPSrCqQDk8/WvHR6LOW{+}Imom00NbSJcz6hItogORndw347c9eOlatvpMFtosFgMNHGgRsqPm9cgVh30n9uePoLfyy0Omxb3crlfNfBC88AgbTxnr2rsIoX3L0PbrQ1d3KbUUkeBfGD4LeGLTSb3xDp/h8XGvKVhtI7ZDuaZ84YYIwQEc57cHqBW94f/Zr8B6Lrl3qH9kpfagJ2aWW9czBZs/PgMf72etb/wASvGcvh3V9Jt7cB/JkF68QiZ2lXBi2kgEAbpE565YYB5x2WhxXCaNDLeqq39xmaZMdHclmH4EmhXbsTzySu2Ri1S2Ty40CBRgBRwK8x8RO114p09baQtCk0yypt484BGK4PT93JGwcHpuHcivV7v8AdRuCwfIP515/qM0kmqw2a3qu0121zJCigiKN9giJbrnbEx2/7YJHFdFJJO5jNtpHQzxStHkj7oAVxzx0AAr1n4TW7/2M05G4s3BORivMpZoki3KxA{+}6oHpXt/gu3Wz8O2iKQVKBj9TzXs4BXbZ5GYyaikdAmSvNIUIAYdenFLv6elIz5PBxXtHgAgKHHX1IpzqTj09qWNPlPORTRlcjoaABHGMHtTpX{+}Wq/O/qKeWJAGPpQFhgLMxGMmoNSivZdKvl0{+}dLe/eB1tpZF3KkpU7CR3AODirkJAJJ9akPJJA{+}ooEfPmhfGbxX41v9Nt7BItNi8UFW0aWS33tbC2jY3/AJoPpKqov{+}/S{+}FvjL4o8Xano1vHBDYReJZIo9Mka33GA2qBtTEgPU7w8ae65r6B2qzcfezwaYyBM7jnnH06cUxHgfhP40eKfEGraLaSQQxf21NBp9vtg4jubUxNqoJJ9GuY0HZrVjzmrHxA{+}IUWl{+}MtV1vR/FVjbIPC0V1YRtsmj1KZZ7kpCmT824jbtj{+}c8YIwc{+}6RMFJJ6/wAqNm6QYQYHI46UCseUXnxB1WwsPGms6pqL2ljp{+}pRaTa2tvbwgwNIttiSSSRguVeYgsxVAuSQcCuZtPivf6rofh62k8XWmla1d6peWyyzNaiKezguNpmckYZtmxF8ogM8gIG0Er9ASKBGVZAVYYIIyCKhcLIoGOBQOx86eKPjLf{+}FPBuoSWd0LTWLW61{+}7iidYRBdLb6hcRxxHzW3uTtGViG7nOV{+}UHvvH/iiN/EHhGKw8ULpdqdck0/UPJaEqZfsskiQSFwcMTswvBO8d8V6PKqM6Lt5ByM9qhmQBU2orDdkkjrg//WoA8c8D{+}NdX8X6leQXwji/smNNP1ALHtEupKzCbb3CBVjdfUTj0r1XTIlDAgcngVi3wDXhJ5{+}cc5rcs32xgKef1pGj2NNYMjnOf97FFEcZKD5tvrkUUGRiNUbd6KKAOB{+}KhK{+}F73BIyvUdeteDeGZ3ktpGY5zKVI9R70UU0aRL2t2yRWtwq5Aj6fpXT/CG9kiKwqFEYYMBj6f40UUyz3aykYtMemGHArYVice/NFFSZMeTiMnvkCnNGBx{+}NFFBIkKBj1PrTj1/GiigY2UBsEgHJxUVud5wenIoooKJ5IVwG7jiowdrpj1xRRQBalXDZHWoyxIXJz8tFFAEoG1Cw4OKCxIUnqRRRQBFccpj1HXvXhHi2If21dnJwrkBT0FFFefjf4R6uX/xDIjOCOO1TRudiEcZ54oor5GW59gtjN8ZRpL4W1JWUEGBs{+}/y5rhPh9qlzaeJ5LIyGdLnT4tSlknJZ2lYuMZ7KOcAAdTnPGCiqRjPodLq9uJvHWmXxZxNHAJFCtwMtsIx6YY16AHLQjsRgAj8DRRW89l6GEeoA4A7njk1h{+}ONOh1Lw1dQzbtq4dWB5Vs9RRRU9Gax3LOlSvLpFgztvkaFdzkAFjjqccV598bdOi1LR4rWUssbxTZMZ2kHZnOaKKcHZpi3k16niUPwT8P8AgX9m7xnYabNf7ryS0uZLt5lE26PeV5VVBHzN1B619Ix2g0XxVpelQSSNYpZrPHHKdxjyCmwHrtAUYBJx2oor0swlJ06d33PPy{+}MVUqadjurcfdxxk44{+}hq/CoPB5BOOfpRRXlI72cR8PMTXHiC5dQ1w{+}oSq0hzkhQNoP0zgV20JJRCepJooq6exFX4zzzxJKL744{+}GrGaNHt7LQby{+}QY5aUysvzHuAAMD1Ar0C6uHaPqASwGQOeRRRRH7RjDoY{+}pMWDR5IVsA4NeYG3Fr8RtTjSSQot40ADOSAqRgDjoDjAz7D3yUVcdn6G/2kej2EC3OqW0T5KF14HpnpXvViAkSqowqrwB0FFFe9gPgZ89mPxo0YjkA{+}opQN0mD0oor1TxyeMc47Diq0zHIOetFFAkJGNwJJzinkZH1oooBj84K{+}9PxndyaKKBMSNByMdASKcyjdjFFFAIiKhCQPXFCkgj6ZoooBiyOSSM8VCrHgduaKKBjFUNICeuP61Wu2OdoJAJzxRRTQjlrljJfoCeAe30rfsQCVGOMUUUi5bGrCA0YJA/CiiigzP/2Q==')");



        //    driver.Quit();


        //}

        //////Tirar foto
        ////driver.SwitchTo();
        ////    driver.FindElement(By.Id("btnCamera")).Click();
        ////    Thread.Sleep(200);
        ////    driver.FindElement(By.CssSelector("#pnlCapturePreview > div > div.section-content.buttons > div > div.imagebutton.accept.dark")).Click();
        ////    Thread.Sleep(20000);

        ////    //Bot�o Avan�ar
        ////    driver.FindElement(By.Id("avancarBiometriaFacial")).Click();

        ////    //Cenario 7 -- Formaliza��o 

        ////    //Validar P�gina Outros

        ////    wait(By.Id("avancarBiometriaFacial"));


        ////    Assert.IsTrue(IsElementPresent(By.Id("panelAssinaturaEletronica")));

        ////    driver.Quit();


        ////}

        //[Test]
        //public void CT007_Cadastro_de_Clientes_Outros()
        //{




        //    //M�todo de Print
        //    ScreenShot print = new ScreenShot("VendaCartoes", "CT001_Cadastro_de_Clientes");

        //    //M�todo simulador de teclado
        //    InputSimulator simulator = new InputSimulator();

        //    //Execu��o

        //    //Acessar P�gina do CCM
        //    driver.Navigate().GoToUrl(baseURL);


        //    wait(By.XPath("//button[@type='submit']"));
        //    wait(By.CssSelector("img.png_bg"));
        //    print.PrintScreen();

        //    try
        //    {
        //        Assert.AreEqual("Login", driver.FindElement(By.CssSelector("strong")).Text);
        //    }
        //    catch (AssertionException e)
        //    {
        //        verificationErrors.Append(e.Message);
        //    }
        //    Assert.IsTrue(IsElementPresent(By.Id("login_username")));
        //    Assert.IsTrue(IsElementPresent(By.Id("login_password")));
        //    Assert.IsTrue(IsElementPresent(By.XPath("//button[@type='submit']")));



        //    //Validar Tela de Login
        //    Assert.AreEqual("PSF Security", driver.Title);
        //    Assert.IsTrue(IsElementPresent(By.CssSelector("img.png_bg")));
        //    try
        //    {
        //        Assert.AreEqual("Portal Lojas", driver.FindElement(By.Id("tituloPagina")).Text);
        //    }
        //    catch (AssertionException e)
        //    {
        //        verificationErrors.Append(e.Message);
        //    }

        //    print.PrintScreen();

        //    //Realizar login
        //    Login();


        //    wait(By.Id("btnSearch"));
        //    print.PrintScreen();


        //    ////Alterar Filial
        //    Thread.Sleep(2000);
        //    SendKeys.SendWait("{ESC}");
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("labelFilial")).Click();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("changeFilial")).SendKeys("54 - L 054 JD / SP");
        //    Thread.Sleep(2000);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);
        //    Thread.Sleep(2000);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);
        //    Thread.Sleep(5000);


        //    //Validar P�gina Inicial
        //    Assert.AreEqual("PSF Security", driver.Title);

        //    try
        //    {
        //        Assert.AreEqual("Buscar por:\r\nCPF\r\nCART�O", driver.FindElement(By.Id("tipoSearch")).Text);
        //    }
        //    catch (AssertionException e)
        //    {
        //        verificationErrors.Append(e.Message);
        //    }
        //    Assert.IsTrue(IsElementPresent(By.Id("cpfSearch")));
        //    Assert.IsTrue(IsElementPresent(By.Id("btnSearch")));
        //    try
        //    {
        //        Assert.AreEqual("Concessao do Cartao", driver.FindElement(By.Id("menuCollapse3162")).Text);
        //    }
        //    catch (AssertionException e)
        //    {
        //        verificationErrors.Append(e.Message);
        //    }
        //    try
        //    {
        //        Assert.AreEqual("Atendimento", driver.FindElement(By.Id("menuCollapse3164")).Text);
        //    }
        //    catch (AssertionException e)
        //    {
        //        verificationErrors.Append(e.Message);
        //    }


        //    //Cenario 1 - Cadastro de Clientes

        //    //Consultar Cliente
        //    driver.FindElement(By.Id("cpfSearch")).Clear();
        //    driver.FindElement(By.Id("cpfSearch")).SendKeys("5207901117");
        //    Thread.Sleep(2000);
        //    print.PrintScreen();
        //    driver.FindElement(By.Id("btnSearch")).Click();
        //    Thread.Sleep(2000);



        //    wait(By.Id("iniciarCadastro"));


        //    //Validar P�gina Cadastro de Clientes


        //    Assert.IsTrue(IsElementPresent(By.Id("formPreCadastro")));

        //    driver.FindElement(By.Id("dataNascCad")).Click();
        //    driver.FindElement(By.Id("dataNascCad")).SendKeys("25/10/1995");
        //    driver.FindElement(By.Id("iniciarCadastro")).Click();


        //    //Validar Mensagem de Retorno CPF

        //    Thread.Sleep(8000);


        //    //if (this.IsElementPresent(By.Id("dialog-message")))

        //    //{
        //    //    wait(By.CssSelector("#modalContent > div.modal-body.row"));
        //    //    string Retorno = driver.FindElement(By.CssSelector("#modalContent > div.modal-body.row")).Text;
        //    //    // var setMensagemRetorno =
        //    //    Marisa.TestDataAccess.ExcelDataAccess.SetMesagemRetorno(Retorno, Cadastro.CPF);
        //    //    Thread.Sleep(2000);
        //    //    print.PrintScreen();

        //    //    //refazer o contador
        //    //    //int counter = 0;
        //    //    //int NumeroCPF = Marisa.TestDataAccess.ExcelDataAccess.GetSemMensagem().total;

        //    //    //while (counter < NumeroCPF) ;

        //    //}


        //    wait(By.Id("avancarDadosPessoais"));

        //    //Cenario 2  -- Dados pessoais


        //    //Validar P�gina Cadastro de Clientes


        //    Assert.IsTrue(IsElementPresent(By.Id("formDadosPessoais")));


        //    //Nome da M�e
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("nomeMaeDadosPessoais")).Click();
        //    driver.FindElement(By.Id("nomeMaeDadosPessoais")).Clear();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("nomeMaeDadosPessoais")).SendKeys("Maria Aparecida");


        //    //Estado Civil
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("estadoCivil")).Click();
        //    Thread.Sleep(2000);
        //    //new SelectElement(driver.FindElement(By.Id("estadoCivil"))).SelectByValue(Cadastro.EstadoCivil);
        //    driver.FindElement(By.Id("estadoCivil")).SendKeys("CASADO");// PEGAR DO BANCO
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);



        //    //Sexo       
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("sexoDadosPessoais")).Click();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("sexoDadosPessoais")).SendKeys("MASCULINO"); // PEGAR DO BANCO
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);


        //    //Deseja Receber a Fatura por email?         
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("faturaEmailOutros")).Click();
        //    driver.FindElement(By.Id("faturaEmailOutros")).SendKeys("N�o"); // PEGAR DO BANCO //


        //    //Email
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("cliEmail")).Click();
        //    driver.FindElement(By.Id("cliEmail")).Clear();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("cliEmail")).SendKeys("starline@starline.com.br");//PEGAR NO BANCO

        //    Thread.Sleep(2000);

        //    print.PrintScreen();

        //    //Bot�o avan�ar
        //    driver.FindElement(By.Id("avancarDadosPessoais")).Click();

        //    //Cenario 3 -- Dados Residenciais 

        //    //Validar P�gina Dados Residenciais

        //    wait(By.Id("avancarDadosResidenciais"));


        //    Assert.IsTrue(IsElementPresent(By.Id("formDadosResidenciais")));

        //    //Cep
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("cepDadosResidenciais")).Click();
        //    driver.FindElement(By.Id("cepDadosResidenciais")).Clear();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("cepDadosResidenciais")).SendKeys("06033050");
        //    Thread.Sleep(2000);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);



        //    //Numero
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("numeroDadosResidenciais")).Click();
        //    driver.FindElement(By.Id("numeroDadosResidenciais")).Clear();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("numeroDadosResidenciais")).SendKeys("1010");

        //    //Tipo Residencia
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("tipoResidencia")).Click();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("tipoResidencia")).SendKeys("ALUGADA"); // PEGAR DO BANCO
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);


        //    //Telefone Residencia
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("telefoneResidencialDadosResidenciais")).Click();
        //    driver.FindElement(By.Id("telefoneResidencialDadosResidenciais")).Clear();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("telefoneResidencialDadosResidenciais")).SendKeys("1135926538");
        //    Thread.Sleep(2000);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);



        //    print.PrintScreen();

        //    //Bot�o avan�ar
        //    //Thread.Sleep(2000);
        //    //driver.FindElement(By.Id("avancarDadosResidenciais"));

        //    //Cenario 4 -- Dados Comerciais 

        //    //Validar P�gina Dados Comerciais

        //    wait(By.Id("avancarDadosComerciais"));


        //    Assert.IsTrue(IsElementPresent(By.Id("formDadosComerciais")));


        //    //Classe Profissisional
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("classeProfDadosComerciais")).Click();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("classeProfDadosComerciais")).SendKeys("Outros"); // PEGAR DO BANCO
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);

        //    //Atividade
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("atividadeDadosComerciais")).Click();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("atividadeDadosComerciais")).SendKeys("Outros"); // PEGAR DO BANCO
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);


        //    //Renda Mensal
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("rendaMensalDadosComerciais")).Click();
        //    driver.FindElement(By.Id("rendaMensalDadosComerciais")).Clear();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("rendaMensalDadosComerciais")).SendKeys("10000");


        //    ////Tempo de Empresa (ANO)
        //    //Thread.Sleep(2000);
        //    //driver.FindElement(By.Id("tempoEmpresaAnoDadosComerciais")).Click();
        //    //driver.FindElement(By.Id("tempoEmpresaAnoDadosComerciais")).Clear();
        //    //Thread.Sleep(2000);
        //    //driver.FindElement(By.Id("tempoEmpresaAnoDadosComerciais")).SendKeys("10");



        //    ////Tempo de Empresa (MES)
        //    //Thread.Sleep(2000);
        //    //driver.FindElement(By.Id("tempoEmpresaMes")).Click();
        //    //driver.FindElement(By.Id("tempoEmpresaMes")).Clear();
        //    //Thread.Sleep(2000);
        //    //driver.FindElement(By.Id("tempoEmpresaMes")).SendKeys("10");


        //    //CEP
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("cepDadosComerciais")).Click();
        //    driver.FindElement(By.Id("cepDadosComerciais")).Clear();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("cepDadosComerciais")).SendKeys("06033050");

        //    //Numero
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("numeroDadosComerciais")).Click();
        //    driver.FindElement(By.Id("numeroDadosComerciais")).Clear();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("numeroDadosComerciais")).SendKeys("365");

        //    //Telefone Comercial
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("telefoneComercialDadosComercial")).Click();
        //    driver.FindElement(By.Id("telefoneComercialDadosComercial")).Clear();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("telefoneComercialDadosComercial")).SendKeys("1136811452");
        //    Thread.Sleep(2000);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);

        //    //Bot�o avan�ar
        //    //Thread.Sleep(2000);
        //    //driver.FindElement(By.Id("avancarDadosComerciais"));

        //    //Cenario 5 -- Outros 

        //    //Validar P�gina Outros

        //    wait(By.Id("avancarOutros"));


        //    Assert.IsTrue(IsElementPresent(By.Id("formOutros")));

        //    //Vencimento Melhor Dia
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("vencimentoOutros")).Click();
        //    driver.FindElement(By.Id("vencimentoOutros")).SendKeys("10");
        //    Thread.Sleep(2000);

        //    //Confirma��o de Vencimento
        //    SendKeys.SendWait("{ESC}");
        //    Thread.Sleep(2000);
        //    SendKeys.SendWait("{ESC}");
        //    Thread.Sleep(2000);
        //    //Cobrar tarifa de overlimit


        //    wait(By.CssSelector("#tarifaOverlimitOutros"));

        //    driver.FindElement(By.CssSelector("#tarifaOverlimitOutros")).Click();
        //    driver.FindElement(By.CssSelector("#tarifaOverlimitOutros")).SendKeys("N�o");

        //    //Conta B�nus?
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("tipoCartaoOutros")).Click();
        //    driver.FindElement(By.Id("tipoCartaoOutros")).SendKeys("N�o");


        //    //Campanha
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("campanhaOutros")).Click();
        //    driver.FindElement(By.Id("campanhaOutros")).SendKeys("931 - Desc Primeira Compra Marisa Itaucard 10%");
        //    Thread.Sleep(2000);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);


        //    //Bot�o avan�ar
        //    //Thread.Sleep(2000);
        //    driver.FindElement(By.Id("avancarOutros")).Click();

        //    //Cenario 6 -- Camera 

        //    //Validar P�gina Outros

        //    wait(By.Id("avancarBiometriaFacial"));






        //    //   Assert.IsTrue(IsElementPresent(By.Id("fieldSetBiometria")));

        //    //Ligar Camera
        //    Thread.Sleep(2000);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);


        //    Thread.Sleep(2000);

        //    driver.Quit();

        //    SendKeys.SendWait("^+(J)");
        //    Thread.Sleep(2000);
        //    SendKeys.SendWait("acessoFlash.faceDetect('data:image/jpeg;base64,/9j/4AAQSkZJRgABAQEAYABgAAD//gA7Q1JFQVRPUjogZ2QtanBlZyB2MS4wICh1c2luZyBJSkcgSlBFRyB2ODApLCBxdWFsaXR5ID0gOTAK/9sAQwADAgIDAgIDAwMDBAMDBAUIBQUEBAUKBwcGCAwKDAwLCgsLDQ4SEA0OEQ4LCxAWEBETFBUVFQwPFxgWFBgSFBUU/9sAQwEDBAQFBAUJBQUJFA0LDRQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQU/8AAEQgB9AH0AwEiAAIRAQMRAf/EAB8AAAEFAQEBAQEBAAAAAAAAAAABAgMEBQYHCAkKC//EALUQAAIBAwMCBAMFBQQEAAABfQECAwAEEQUSITFBBhNRYQcicRQygZGhCCNCscEVUtHwJDNicoIJChYXGBkaJSYnKCkqNDU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6g4SFhoeIiYqSk5SVlpeYmZqio6Slpqeoqaqys7S1tre4ubrCw8TFxsfIycrS09TV1tfY2drh4uPk5ebn6Onq8fLz9PX29/j5{+}v/EAB8BAAMBAQEBAQEBAQEAAAAAAAABAgMEBQYHCAkKC//EALURAAIBAgQEAwQHBQQEAAECdwABAgMRBAUhMQYSQVEHYXETIjKBCBRCkaGxwQkjM1LwFWJy0QoWJDThJfEXGBkaJicoKSo1Njc4OTpDREVGR0hJSlNUVVZXWFlaY2RlZmdoaWpzdHV2d3h5eoKDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uLj5OXm5{+}jp6vLz9PX29/j5{+}v/aAAwDAQACEQMRAD8A{+}ysEEd6UKMUoFKB6VzmoEZ7UpBHNKBjp1pcHrigBQgxSr{+}tHPrR3680FJAQc04DJx3pcZwe1KR0x{+}dADSMYP508LkDOBSgH604Ad{+}nvQITGf8aAuOOxp2D74pwHAFAhqjAxjinbenrTgABRj0oATZxmlwKUDINKBngUANIPbpTgvJ60/Hy0DNADQm0U4Lx0NL1Jo6CgBAvB7Um3jjpTyKMZ7UAIeeKMY5pQRz60uc9eBQA3qKbg9qeRjGO9ZuteIrDQLfzbucKTwkYxvc9gB3NAzR6UY/CvJ/FP7Rvh3w/Ym4jhu53UkPFJF5RXB{+}bO7HTjpnrXjJ/bKtvF3iY2dhetoIiRntpn5hmJwCkwI454BHfrjurlKLPr48LRXzp4U/a10yW/GmeIJYLe5kcxx3ETDZ5mf9U2DwcBiM44BHUZbgF/bwXSfG82k6jbxXNqjGDfapklwcbgSRlTlT2PWlzIfKz7KB9aQ4r5z0b9q2114XF3bqrWMR42LkhDvUFs/dYtG3GemPcjJv/2r/wC15mfS41NqjYJdShAAB388YJBwMcj60udDUGfUQUEZByKTHH614t4P{+}N/9rwLHO8bzkKyrbKcMCOSR1B//AF9K7rT/AB9a3yZgZZl3EBklXY56YDHr{+}HFNSTJcWjrsACgjOQRWfZ63bzIGdhDnqHPH51qDBGTgVRJERwM5FGB16mpCucZ5z0pNo7/SgRGPel28cUpGRSbSQBQA3bgY60DnFPH1oxnkUAMxgCmhRg9vepMU3nsKAGbeKTbgVJ25pMc{+}tAEOM4oPHSpdtNx69KAI/wAqRhkipD1pmMEmgBhXHHQUmNpx6mpaaVLexoAjOCO/4UzGD609htoI5oAjOV7UHJPtTyMmmkcUAMYYNNIAFPAyPfPrTTzmgCNuKYRgZxU2OOcUbaCrkA4xzRjt0z3qRlweKYcfWgLjCoB5NFPwDyOaKBloAk88e1L0ODxS529qXGcZoJsIBgDjFP8A50HhvUUu0EmgBOWpwUKM4z70u3sOtOCY/GgBAM9KcBjHenAY96XHGKBAB6UY/P3py/r2pcZoAAKVVJJPSnKo7Ud8UABUD3oHSnAc{+}1OAoAaAeeAaXHPvS9/0oIA79KAEJ74pRyRijPalFAwxyaNtA4OKO/vQAE88UvTmgAk8Vn6zrlpoVsZruVYx2DMB{+}poC1y3LIkSlnIUVz{+}t{+}NtN0clbi6WJ8Z2Zyx/AfhXknjn45RQG4eKXMcakpCmVY{+}{+}OCM9iR{+}FfL/jz44re3chsSLuWSJlaPYWCH/gWR19qydRLY2jSbPbfip{+}1XcaJqSw6JDcXVt9151Awpx6EjPPvXz547/ah8SeLJlsTd/wBnbyY5JI23YyMfdPAz1z7da4K98WatNJJJPOgWUMrRsVcbDn5cdB17CuJ1Ywx3LyyLuUkEN0z06UlJyNuRJHUf8Jt4zsNGu2GrHUI3fy5bXUHEsTcfKCjcq2A2GBOeQMV5t/bk{+}m6g15ChjLffQEyQyAjBBPcHPIPr69Ld14n0u0EkSQqrMuxnExYFfQjGKxtQ8Q20loI1BRMYXfGOc9cYGa0WpDXYt3V9c6hNJe2kzrI{+}BNCzlnBB4YE9ccZ79euTUGurN50N4gKO3zFASdvHPPX8Kx4vEAiKsrpJtPGVK9a2x4ogeUSbg25Nx45zzkGrsTddTQ0nxLqmk20wW4kRLtQ7KhPY7hkfifzq5c{+}L9Wg0uKBJJ5PtDqzxY{+}V8Y2rjvyP51Ui1rS7pYXnB{+}U79sePmAPAOemeK1LbWZtSjDpbxRpyPMdsAd8etQ0aRV9mdR4a{+}KHiPSdCmtrm{+}nneQEN821FyMFSe4AxhRgdc9Oe2{+}GvxNh026dtVv5ruJcqENx5SIeehPzH8Oeeteax3kKwGKS/2x9dsKbQc9eWIHr2q0jaPIuBdT8AgbQTn8RkYPvWZSjY{+}xvhv8bfA99G9leaxZ285YgK7HbKuONpbkntz6dK9r0bW45reFtIujFnBCRzAwP2GCOV/AAeor8yG0PS7gs/2yNhySkpKN79scVs{+}GPEvijwiBJ4W8TyafEcbYSziFvqCCPrxT5kiXC5{+}p{+}j{+}LJZ3MVzbypKpw6ylQyfXB5Hoe/wBeK6WG4WZcjHX1zzXwB4O/bF1XTHt4vGenxIQAi6lY4ZR/tfK2Meo/IV9U{+}BviQnjGwt7zRruzvkkTdEIZBukHQjacbh9MfnVKSZhODR60cY6UY6VkaV4ijv51tp4JLK9xkwyc7gOpU9xWuDz61ZmJtFJ0PtTu1FAhCAB0zSFSBQcjgUoOetADCuaaUINS{+}1Z3iLVRoHh/U9UaIzLY2styYwdu8IhbGe2cUAXCtJgEHpXmf7O/xwtf2gPhZp3jWLTG0KG9nmgWzmuBKQY5Cn3sLnO3OMV6cyYHWgCM8D1pCvX60u5c7cgsOoB6U1ZElBKMrgHB2nOD6UAN9ecYo9K5vXfAFnrOptqVveXuk6myhWurCYoXA4Acchqoz6V4p0SymlbxXaTwRIXMl/p4GxQMkllYfniueVScW7w08mv{+}Ackq9Sm3zU3bumvxvY6yW4himhieVElmJEaM2C{+}Bk4HfAp2Ac56V8/T3nivU9B1jxzc2uo6/HplrJ9h0jSN0E{+}o4IzHFtUsobAyQMnGO3HsPw88RXfivwRousX2hXXhm7vLZZZNHvcia0P8AzzfIU5H0FKhWdePPy2XQjB4mWKi6jhyxe192u5vDPQDgUh5FKkiSLvRw6k43KcivG/2l/g/42{+}L{+}iaNaeCfiDe/D{+}6s7h5bi5spZozcIVwEPlOpIB55zXSd57Fjk008E{+}1fmR{+}xb4c{+}MHx81h/Ecnxm8R2{+}meG9YtftmmXupXc63sYYSMh/e4wyqVIII5r2P9n/x34q{+}Gn7Z/j/4TeLfEmra7pepxNe6A2r30tyY0XM0aRmRjj9y8gbHUwe1VYm59osDzjFJjOQcYr5N/wCCiPxf1/wP4H8L{+}EvBuo3uneLfFOqLFBJpk7Q3AhjIyqOpBUtI8K9eRuHrX0p8P/Dd54R8DaDo2o6nc6zqNjZRQXWo3kzTS3MwUb5GdiSdzZPPrSsM3Svp0puOPwp/8RpCOwxSGRZC8Y/Kin7d3aigdy2OTz0o6ilC807HpQMQen61IBQq8A07oKCWJjBx39acBRt5zk08AnHFAhqipFX3oAHHrTgcnFACFQKMY/GnhBz0NLkD60AJ0ApwHXNIBxinY5oAQj0NKB0oI3Z7UYoGLSdM8ZzThzRt96BpCY5NAGR6Up9KcBxQDG4xRjLU7O2vOfiv8V7bwJp7QwNHJq0qM0aOSUhUdXfHQDsOM0m7Ak3oi78R/itpXw/s281jdai4Jhs4cNI57cfXuePWvjP4vfHPVtR1JG1K6Tz2fdDpNoykRAcgyOfunPcc{+}nGK8{+}{+}JXxpe8ur2LTbqa{+}v5pMXOpzEHeeDhQOAuQcKMjivHrvV5J2bYfPvJGy1xISVDexHLH{+}tcrk5Ox2QpqKuzuNd12a/RzfX8cClyzRQAqpB67j1J46MfTgVz8{+}viRZE02Jyqr887ybcH1JPA/n6VUsLc3siwRQNq1ymczz/JDEfoD264H41sXljpOgobjxPq4mmABis4R37BYx29249B3qdIvzNldmFGWv4WEk4H{+}3EjON31JXP61y{+}s6ZZacrPdPlucBsKx{+}uc4{+}ldJrPjzV9VRrTQdNj063xta4kO6UL0GWPC/z{+}tefyaG19ds91cyagxOA6ElSe4BPJ/CtaalvLQym09I6mfJqGntI8iQ{+}bjj5cnd6CqV7ezzHesQjJHRu1dBdaZFpcCmZ4rdv7gOSv8An1Nc5e3lvG5EOWH/AD0f19hXWmuhzSutyuqTN/rnGPYVqabYQzXCwuxYAEsQcHODgVlWxN1couCQCM59M1t{+}HrS4u73bGoLP3HNEnoTFXEtNtrclJTjYSBnnHpWjBqwkPlpyx{+}7xnPr9D{+}FVNes/7PuXgkBLZ5Kj8v1rBeZoZw{+}4kAjuR/Kl8SLb5XY7IeLZokeAQwuoYksyq34/dzSDWLW6iCFDCwPXoT/WuQffOzTLCSMljsGRgfyxU0LQSHcztH6FcHB9x3/MUuRIFNna21leLkxs8iHLhW7{+}mfX8c1asfEc1g3lyQBOMNKh2t{+}XXP0rB8P67f6BcpPEPOt9wyUG4EjsQRwfpXf6a{+}heMYN7TRWN2/G9f9WTnow7ZrGUuV{+}8tDogufWL1LlhdnxDbF7O5tp2C4NvLIN2M88Hr{+}FXvDHi7WvhxqkV74d1KfS7tWJNpM{+}YCw7Ecj17VyGp{+}G73wzfEL{+}5lyGWTGVYeoPQj9PXFbdl4i03xIqWNwbfS9UPUSswikbHGRg7CfUcf7OOazdlrEvfSR93fs/wD7U{+}nfE22i0TxMyaZ4lgAXc52rKegKNnr6EHn8a{+}n7C6EkSB33sR8r9Q3/ANevxde5k03Ultr1WsLpMPDKrBSM85Rx1B/EGvrv9m79q680CSHw54uumvdOY4h1E5Z4TgZ3jqV557gc{+}hq1O25hOn1R95Y60mDnsKpaNq0Gr2kcsMqSoyh0dDkOvqCKvkYrY5Rh4IzSNk9KeRnGRSbT68UCGHp1rnviM3/FvPFA/wCoXdf{+}imroyvX6VleJ9IbxD4c1XS0lELXtpLbCUjIQuhXOO{+}M0AfnN{+}xz{+}xF4W{+}PnwDsPEvjTX9buvNmubfSrGxvBHDpiLKwZlUqw3tJuc9sEcZ5rH1Px94qs/2Rv2h/hvrWu3OtzfDzXrDTbHWXkbzXgbURHs3ZzhTAxAJJAfbnAAHsPgj9iD4zfBXwwNG{+}G/xtj0q2vt7apbXWngwiQsQJbfIcxsY9gOMHK53YwF7dP2EbPS/wBl/wATfC7TvEjSeIPEt3Bf6p4mvrcu086Txy58vdkLiMgAsTlmYkkmtLk2Pnzxx4R1n4VfB34OeAfDXi3VINY{+}NOo21zr{+}v3VxmRS8dshiRhhhH/pC5GdzCPGcMRVv9rT9jfQP2cfgJqniXwD4n8QabIj2trq1rc326LU42mQKWVVXDrJscY4wrcd6{+}p/jD{+}yZp3xf{+}B3hXwVeavJpuveGba2Gma/axndDPFEsZbZuB2NtyQGBBCnOVrxrxl{+}xH8ZfjP4RfRPiT8bU1W3stjaXbWungQmUMAZbggI0hEZcDOTls7uCGSYWPrX4WlpPhn4SdiWZtHsySTyT5Kc1yWt6hJ8VvEx8PadIy{+}HLFw{+}pXUZ4nYHiNT6ZH6E9hnM8T6tqWh{+}HdG8A6K7vPaWdvY32qFCkcahFjznnaGxyfwGT06nw1qnhnwBocGl2tw1xs5mmihZvMkJZSScesbADsAK8mrVjWqeyvaK38/L/ADPCr4iGJqvD81oR{+}J9/7q/X7jmf2tIxpn7L3xEWz/0UW{+}iSiLyjt8sADGMdMV8Z32reIPix4f8A2aPgfDr95oOheJNBGo6zeW0pWW7jXzT5ZY9flgkwDkFnUkHaK{+}5PjVosHxT{+}FvjPwfHdTaa{+}oWUtm189nJKkROBkKMF{+}TjA68{+}hrwPxn{+}ylpHib4UfDjRtO8ct4e{+}IHgOBYtM8SxWkkKtgglXUkFRuVSDuypzwckH0VUprS6PX9tSWnMjxP9rv8AZh0r9mjwr4R1DwL4i1u00TUvEFtaX2i3d6ZIpZgkjR3C4AwwVZFPX74xgZB/SwLjtXxL4v8A2IPip8Z7bSL/AOIvxki1vUtLuo5bCC201Vso4ersQnl7pGwnzY4APJyMfSem/D/xfafHTVfF9x40mufBlzpy2tv4UMbeXbzjy8zBt2MnY/b{+}OtHqjoifKf8AwSdGfBnxF/7CsH/otq1/{+}Cg{+}g3Xw68XfDP456PCz3nhvUorLURGOZbcsZIwT2U/voz/12Ar1r9kH9l28/Zh0LxLYXniCDxA2r3cdyskFsYBFtUrggs2etem/GX4Y2Xxj{+}FviPwbfOIYdWtGhScpv8mUHdFJjjO11RscZxRfW4W0Pj/wjd237VX7f0viS0lXUPBfw{+}0{+}I2ky8xSz4JjP186SRwe4tx9K{+}7T3B614h{+}yN{+}zFB{+}zD4I1TSX1OPW9V1K8{+}1XOoRwGEMiqFjjClicL8x69XavczyMmkwRFijocDHNKw9OKTAH1FIBrIc8GilNFAFvoen4U8D1oUZNO2jGTQO4m3JPangY{+}tA7U4DpQIQCnhBnpRgAdOaeCPWgBAuTnPHpTqTG7qKcORzxQAEcgUBR260fdpwoAbjnFOUZpSvB9aUcdaCrCY4FKDn2pQe9GMmgYntShcmlC4NL0oFcQ9MUD1pc84qlq2q2{+}i2E97cvshhUs2Bkn2A7k0CMLx54xg8K6XI5eNbgoXUOeFAIG5vbJH16V{+}cv7Svxxi1nVrvRdNuGkjLZvr/ILv2CL6ep9zjgDn0L9rP42NpSTWFrdJNrF7hrl1{+}ZYQMhY19AmSM/xOSeq18PzXD3E7PLI26QmTLnknu2O/f2GaxfvM6YRUdTXk1J7{+}4dUQQwgn9zv6DPc9f8{+}9b2kaPLeeXJLI0duQTiIhS44yPZRnk1k6PpQeNJbiJlhJyFb7x{+}uOnX1zzgZOTXoFlpn2ezWedTawyR7sv8u1QRh39v7qAelZyfLojpinLVmffaxdafa/2foCN9oC7GnP3VXPb2HqfwFZVl4btrVXmnlN3qLnzJb2V{+}hJ6DJ449OTxW3YW6a9dMI7aSG2Y56ZeUHpu75PPHU/y7p9L0fwbbw6prsSXEykSWlhwyRDpvcZwT7HjnvWTnyaLc0UObXocZD4JiksEurxzbaWBuNxLHgOuOSkfUnHQtXI{+}LvHdjBE9h4XsjbWwXbLfTnfNKPd{+}ij2GB6ZrV8XeI9X8aXE9zP50FmzkpbLyzjt9K4zU/DtxIxMwxHjiBGz{+}eP8{+}9awgr3m9TOUrK1NHK3Km7PmlzPLk5y2fzNZps3BLEhf0roLgJpcbxzRhHf5RGBk/U/p{+}ZrInLzsElAiQc4H3vxNdq8jia7jbOJVDlSSeQuD3rsfAOnvLq0K7Q4Y8kDkD1Brl4QoJRCu4cheR29K7L4eh01KL940bZGVCckVE37rLprVGv8UNKFpqUpSOREIG0tjO3HXHbv9eTXkl4PLcqecHoeRivob4gW9vql2GDxiSQBfndSwOOmMgDHHUivEdY0aS3u5V3R7Qfvhs5/Ks6MlaxpWjrcw4JGRgI3ZOc7c96mluWmdTIoUr8u9RtLH1PbNJ5VuG/eTkt6BDx{+}eKvQWVtOFY5mX{+}8eo/Kui6OZJ9CbSbmQvuV9roMgYBBPv7V0VhewTOjCOOzuAT{+}/jO0P7MOn4HrWDFZz21tK1ukM8C/eVic/lmrui62A6xvZwSbuNsiZYe3PUf5xUO0jSLa3PVtD8Tm8sxp2rwGeA52SLnKdOQDyPw9sg4qDxL4J0{+}{+}soLlvmhlbbDdR4zEf7rcH8ue9clpt5cTTJNZ7XAwTAiDeuD1QfxD26/Tt33h3xJbamxSZI1lCeXLbO21ZV6kg5{+}Vs4wT0PXjmuRxcNUdsZKSszgbuK/8PmO210NqegO{+}Eu4jueA84IP8JHXB4PtQLuTQHt2F695pRYG31KMfNC3bcBzx6fXHfPq2q2YsGgjhn{+}3abcf6ppEXEyn70cg6Bwexrz7U9At9KtpH0uXzoptxl0{+}dThQDz69sHrxj8QJ8wnG2x9RfsrftLXWi6nD4f1WU/ZjgYVsqQSNska9{+}uSo98eg/QHSdUg1W0jnhkVw6hvlOQQehB7ivxJ0m8m08oLZmgER82Fsk7O7LkdRjJ{+}nT3{+}5/2Yf2n5NRW30fWrgR3UbbC0pAHJA5J45zx2/PIuL5dGc9SnfVH22wGM03n86hsL6O/tkmjP3hkrnoas9TW5yEZHFIelSN1ppGRigRHgtk0oHHTBp2OKBjNAARkc9qaVB5p{+}MimkDmgDkV8AWlzqM11qDNchpXkjjV2A{+}YjlueSMDHYY{+}mOktrSCxj2QQpCg/gjUAfpVhl5oK4HFZxpwh8KMadGnSu4ojIzzUckayqVdQynggjINTAY96b0Oa0NrEe0KAoAAA49KYeAe9SkA5GOaaUwKBkZOMdD9KYQc56VKwwOlNJxQCGA8U0/ep{+}M9KaQc89aBjXGabtp/WkYZoEMxRQ3WigRe28YpQM9qAOtOHygDvQAbexpQo496cBk5xTjxQAgXb0p2MZoxmlHIoAAO{+}aXHWjGRSqM59KBoAM9qMd{+}1PIxSdRQOwZ4zSZzSjp1p2BjHagY0LT6Tbgj0p3eglsTHWjGKMc0cHjpQIaxA{+}ueleKfHn4hWfh7Tr6S6neKy0qMT3TRnDNI2BHEn{+}0xIGe3zH{+}HFeua3fR6Zp1xdSnaka54OCT2H4nA/GvzB/az{+}LMmoXg0iG7eZY5GubjaTiS4bJyf90HaB257msqj2SNqcb3Z4X8RfG02ua5d3tyPMvZ5CwgUZUc8KB{+}P6/nj6VZXDOuSrXDENNIfuxjsPqOcD{+}fAqnoto{+}ral5rfNs7ngKD/k/5Fehx2Nn4d003F31XkQMMNuI9D/Ef/HB78DOTUFZbnVFc7v0NDSNOsdDtGvNRCuTtNvCxJBPd29vQc5OfQk1rm{+}n1/Uz5m6O2iJU8End06d2x2/lzWMmqtqlx9ovc7m/1UEfAUY47ccY{+}mPrXQR3MWhW8cXk774LuEanCwAgfMf8AaP8Ah9Kzs4{+}pumpeiOutb9fCGnJHDAI9TdT5KkZaIFR8x9SQKzbPRn1i{+}e61S8kuHxubbllU4yAT3bHp09B3zNHjn1PUJprpmk3ktJOxIbOei/ljFdrd{+}K9H8J2Ky6mUs7dFDtAiZYg92P8AQcmsX7isviZovf1eiHPo1rPAPN/0OwQENcEDZGo/idgMsfYcDjpivJPG3j6CeWew0CARW2djXYULNIenBx8ox{+}PoKp{+}PPitq/wARZmhg3aZ4fiYiGBBteUDoSAcA4xx7CuafR106xguL9njjlz5MKHEkvvnsuep784ranR5bSqbmFSspe7DYoym03NlJJmjGN2/5U9cnHJOc/j{+}VKSK3vFyWBVeRyc49O1SN5{+}ouIY4iFHKwxj5VHqfb61BeWDQxqjOCXYk4HHH{+}TXcuxyO{+}6J7Z97jafKReAAQM/SvR/A5TSQ0wmieZsEhXVmHPrj{+}teX2dvIsqmGIs2eDivTfCVlNuhRYZEkYfflU7cd8g/wCevNZVGkjalFt3I9bmW/uHlLiWXaWHdl9sdvwrk5547rcrbCQeA/b/APVXqXifwzFY2yfuTJLEg3vEhI6jBU9wQea8q1SD{+}0LmUFwJM5HQbvyrOlJSRpUg4swNXs5I58jeVb7u7ncP5GqVte3GnMrhQ0YOCAuM/wCfpW2kVwsTQuBcQ9dpPTtkDsfccVXl0zed3LA9Mjk{+}xxXUpJo5HBp3R0Wh{+}ItN1CFYrvdaSLnbMqjK/wDAhz{+}DAj0Fb1/4ImnsBNAsF9GfmSa3OCw6{+}vX615xNozW/76GTcmOUc/Mv{+}I9DXV{+}AvFdxoV{+}qxXAi38NFIu5H5/iB/mMH3rGUWtYG0JJ{+}7NDrWCXTLxeJLOUDOJE4Y45z3z{+}FdjpgTU7dL2Nle8Ti4RByB2k68g9D0wR1rrW0zQPiPF5bkaZqgUeS0BJRz9DyOfT864jUtA1jw1cMmpQOyRsUS6T7{+}MYwSODwercEd6yVRTWujNXBw9DvNOuEvC9tcI3mEfvLZ{+}OBjGD2IPQ9jjsRkv8ARRPIVMqsCgeKfH3mz7dDg4I/oa57RxJdQE2VxiaEBvKbBJHTABzke3XqPSuwsNSju7cw3AWMmQJvA4jkGcHnv1PuPxrF3TubxtLQ8j8QaLd{+}FJf7StAWs1f/AEi3UYMR/vD/AGD1q9pOpi01O3ngdRHKge3bPyvGeSjZ5BXsfY9sY9J1zQheWi3Fsoa/VWj8liBHOMElQGx1GSBnOQR3ryjVdM/sK1EY3tZb/MiQrseIscMuD6NjH1I7mtVaaM5JwZ{+}nv7LnxQk8WeF7eG5u1uLiJdpAwGAHQ4yfp6Zr6DRwyggg{+}/rX5GfBL4qah8M/EtlewTmS14aSJmOyaPPzKfoOfpzX6i/Dn4gaX4{+}0SO/0yYSRMivtzkqD2P45FVSl9l7nHWp8vvLqdf1xSH2pxGOnIo74roOUY3IAoxTgM9uKUgA0DI1JHHSlK96UqCeKMYxQFhpzjpTSvNSU0jn2oERlSTTSO9SEdfpTSOo7UAR4OfejnGMU8DnNNblsigCNl9utNPAxUjcnHT3ppHNADMZPpTHHOR1FS7SAcDmmFcUDI{+}2Mc0gJXOKkcdwKZjI96BiADFFJkDg8UUEmgq8dKWlAJPtSgY{+}lAC{+}2aOM80AZJpw6CgAPANAFLtz7U480FWGnI7flSgZzSgZNOA4oAQ9PWg9Bx{+}dLtyKXFACKnJpelL2{+}lGM0CEx9KMg5FFLjI4oATtQ5wpINOA46VS1C9W0jlkY/JCmTjrnt{+}NAjwP9sL4qJ4A{+}H81ukgF1dBkCj124A/DIP{+}RX5OazqNx4h1iUMzSSM5LEc45/z{+}Jr6J/a7{+}Ltz8QfHdzZROPsenF4FCnIkmB{+}dh68kDPoorxTwtpo0wJdS4E0mHiJHUDIyPx6HufpXPzWbkd0YaKJteGLaLwlG0jsiyw4ZpGAO1vTB4yP5gDsc5N7q//CR373c5KwrxBE5Lbz3Y{+}pPU/Sq2qXzalcm3XcYi37wjox9Pp2q9penLKzlhtjQDe/YgH7orK9vee5vb7KNOxjTT4F1CVDNcyDEEeBhRuxkjv/8AWqWzEt9MdrSNKRveQdFOQevfA/KotQujdFCGOIsIGUegwPpU8Nz9jtBLO5WNMFkX{+}X1P{+}NJNrXqytH6HSfbYtEtFSEtK6KWBC5APPOPXn6D615zqzz{+}LdXD3qulnE3y8ZcdSOO7Hnr0rfvZLmSeOaRAk0uG8sLkQRkZGR69D{+}IrV0HQjEH1C72QQRjKq43HnnJ9z19TkCrilT1e5Mr1NFsc/LYQ6XbxXD2qG4bi0tCpIAxne31/WsiDw9eazem5vpljThnneQbV77Rjr9BgfhXZLZG9E9/c3rWlmOATzPIPT0Qe3P9a4bxB4n3PsiYxwtwgbJZ{+}eDipU3J2juaeySV3sT6le6ZpVmba1ZpnYAskS7Qx9Wb/61ZVvEb/yFjtUUsDgbN7fePTP{+}FSaJ4e/taTK2iydSSWf169a9u{+}HfgaOyuI7i4s4IVSP5FfdknPufesalaNGL6s6KVCVaWqsjlvDnhBnMDXkCW0ZJyJI{+}uB2AHt6V3lpG9tsdYZU0{+}McM0GAxz0G/rz1Oe1ehWngmS9kSeaaS36kOUHTpwOv{+}RVxPCRuddhjtGlPzpAr3PJRc5JHoeK8aWK5nc96OFUdEeean4f1XxPHF5Ud7FaxjJlPl8tjpnA9OuPz7eSa98PrqNZJyrSQBvMLxjlARnI9ffFfdukeBIdT067t1LBZ8KHT7xQdTk9Mg4/GqHjb4ORQ2VwtrGArHCArkYAA/ofaqp4xxIng4z33PgAxNp8{+}25TzYj9yUH7x{+}v8Aj{+}nWtTTdJtr4vjLSMCD8mHQepXuPcV6j4u{+}GF3pO67Fj/oLtieFUBMZz99R6EYPHWuNuPA7Ws0bQMUVv3ilW4PoVPUcV3/WIz1TOJ4SUdJK5xmo{+}FLuwJnVRPbOCRIBnA75Hp7/n61zWpaHNCouo0zFnG7PzIce3avarGaRn23SBzxlZOrdRk46/X9TRqfg4m2e5tIDKkg{+}eDHUe3/1un1raGLcdJHJUwfMrxPKPDWr3NjNGBdZiU4wCSAfTjJH{+}cV7t4Z8XWfiTSo7PVJHuMkxghh5inHbuR7fQV4jq/hkWzS3VpFNFMv8ACvyt15DKeD{+}GM{+}tRaF4theRLa6TytpwtxGNrBuqn2wa7JRjWjzROGLlSdpI9U8Q{+}F38MznULNvMtS4ZZ4uOeD0z{+}nB4{+}la1vBLrFtJqEQErDas0SqAGXtx2I9faoLXxO95BGtwq3KygKSR{+}6uAQcBwejHnnv9aqafFLoc0V9pbPPbM2DFIfmjP8AccHuOx75rC72kb2W8TYsdWOqW8lncHy5ojtQk43Ln5Tnvzzn9a43xQxsrx0uo/tFhPlJ43XIXPBI/PPHsfWtvUr1kuIr1EQRTMG44aNuMjOMbSc9f6Vo3{+}nReINNunhaEswG2Juq9sc{+}vzDqeo9KV{+}R{+}oviWu55VpswtbybSZz9kuUfNtKcmNmAzj1AbB9eRX0Z{+}zR8XNY8AeK7GG3YTafdt5awPL{+}6yxGUDDgHptz6j3NfOnifRR5dvJHJ5cwO1Wb70cyfwk{+}42/XrUWh6vc28yS20wifIYeXk55z/iPwrWUb{+}9E5094S2P228K{+}IrbxTo1vf2rbo3HQjBB9DWsV5r5K/Yv{+}OK{+}KoDomoTp9t24BJwzN1yR3JwfxU56ivriumEuZXOCpBwk0N7Ubd31p4HOMUoUgetWZoh27ck0m3cKkbnr{+}tIRtPtQVcjbjtSbcdakI3AiowCaCRKaQKcaRgMcigRGVwP6U3BA{+}tS9RTWHFAEeOKaVNPoPPFAEbA5ppGaeR/kU0jHOfwoAjK854puCD1qRgDzTcbgc9KCkxuKKTBPUZooEaABH0pQCSaUdRSnGfegQYOelGM4pRyacBgUDsJjHTijB9aXg04CgYKMe9KBgUUo5oEJR0pSMUdCKBCDoKUdqOATS9hQO4ADFKB6UgpwwOaAEZgiEnNeBftZfFMfDb4d3YgcR31/mKLnqxByfwH8/avc9RuBDbu2Rk4VeO59u9fl3{+}3f8XD40{+}JsukWE26w0jNijB8h5gcyvxxwcL/wABNZzeljWmryuz58uLldY1RrmR5JFJKZHVnJycfiT{+}JHtV3VJPJcRsoE0h{+}crxsAGdqj0A5/Kq{+}nqllB5xjx5ahYUPQnB{+}Y/zP19qoTz3ErCeRmLvlY/U5PLfj/hXLu/I71ojS06JDOkUaY3nLyN/CPTA9{+}P8A9fFtj9sP2a1{+}QZLMw4yO5J9B29/pTLdVsLX92SJSPnkznYg6kn2/U0spj061WWTPny/MV6bEA{+}VPryPxNRd3ujRLRXFunt9Ot0AYKQMk7ulXvD9g19B9vu482kZAhtG/5bSH7uRnpnH6dsVi6DpMniTWPKlZEtYvnmLcDthfzPJ9M{+}lel6RppkYXM6ywQJkRqR07Zwe55x{+}HpW1uVXe5HxPTYqWWlzNcJbB40PzSTyZ{+}VR3OfbB//XWxLYLdwC4kkddItjuCn{+}Nv77e57D0x71fstHazt0nuV{+}RiRb2pXiQqDgn1Rc556nk9BnM8W6{+}2l6etqXS4kdv3cbP{+}7LH{+}JmHbP4n{+}XHObk{+}WJ2U6aj7zPN/GOsTahMqxRh{+}R5Nqik5HZmHpnp6kdhUHhb4ez6/qAmvJN8rnllw2T3xjv61uaPocurXf2iWR5VkHySOu0NkfMUTsOwLdu3GT718NPh8JUjZoSQcfe{+}8w6e3H5VjVxHsY8sTtoYb20uaRneBPhV5EUaWlsUZWyZiu7p2z3r0vR/hu8{+}rwy3SO6RKykAdTkHIJHH8{+}DXp3h/QGtY02xohUAcdh6V1VjoxkVX55JyPxrwJVZTd7n0EacYI5Cx8H2yKAIQqgYJcEkj6mtaDwLvRBbw{+}Rl95ZV2gnGBz{+}P45ru9O0/orREjoM{+}ldRZ6cuc7Dx0yOlQotkTqqJy2keHUsLe3SPBKf3gAMnrnP4Vp6n4dS5tDE3zYHOBx9a6qHTV3Kyqpbp09KtzWCNHtYDJ5rpjS0PPliLs8N8UfDy31eCRGjRSylc7eo9CK{+}aPHnwwn8IyPG8HmaS5yjlSfJYnrn{+}6fXt16V93ahpwJOBjjOVrkNe8J2{+}r2c0M8SOrDoRn6VKbgzrjUU1qfnxq3h8KFWaNnTdlWP3kb1B6g06xN9pdwuzN/C6BSH4PfHTv7{+}2CK9v8AiD8MB4VOxoy2muxMc2M{+}Tnorf7Poe35GvLL7Sm0yeSGX7n8LDtXQp8xbpqxRu/DcGpmW5sURJcAmKRPlOc5DA{+}/9K8o{+}IvwqKCXUbCzkVkOZoI{+}SV4yVPHQ{+}uPwr3XT8XEyOjIlywKuSMrKMd/U{+}1adzpTTBvlQGIbXCjLJx1AHUf0rSniJUZXTOGvho1Vqj5I8OalJpaMt1IskKNslU9BnhXwex9exxnrXpOk6o04LLlkcAOrfd9sjn25HQ89DUvxM{+}FDoJNZ0aNWOGM9vFyrryWKex7r2J9q4Twvqr2ToGkcxMAEdjhl9Vb3/z617kZxxEOaJ4DhOhLlZ6fLCLcSLcFJbKY{+}W8bgAj1B9COuenGe3Oe{+}nzeE9XNlta9spvmgdshivI/qR39agtPEFtbqY70GWBhscY5A5IYeo54PYjHIxW5dWxvrNbaOdJXVPOsLndww6lQfQjP0IrG7jpI1snqjkNb0yOOKW2WQyCdRJEzkkFwMr/ADKn/eHpXnNxJEpW7i3wNGxWdOu0k/MPzAP1Jr07U5IdY08ysxSUqQ4x/q5AecY6A/ex/vV574mt/s1zHdyriO4zDd7P4WHDN{+}ob8RXXRd1ZnFWWt0d38IfH83gTxdaahby4fzElhkGQGCsCf5YwfUV{+}w3gfxVaeNvC2na1YMHt7uISL7eoP0r8NdNuMu9hKyxzQnzIX{+}mQ2P54{+}tfoB{+}wL8dW3t4J1abAkO{+}0d2{+}UHncv1zj86E/Zzs9mY1F7SF10PutV5FOwO1KBn2pQuDmus4SMgHoKZtBqfHHSmMuKBEJTmoyvNTMCaYRg4oHuRMB{+}NNIwcU5v5UNg/WgBnXtTSMnHangYpGzmgREy44pBjHvUrAGmbfUY{+}lAEec/SkK5HNSEcU3GOnSgCIg0EcH0p7LwabyBzQA0ADrRQfpRQO5exgf1pcD1yaKUgYoCwADrS8ijHFKq56mgoRRT6OlL0oJbDFAzQDjrS4x0oEDUgGaX60DOKAF245xRjPNHagelAAOaacucDge9KRztzRLKttA8jsFVBkk9KBnjf7U3xbi{+}EPww1HUYpVXVbhTa2QPXzGHL4/2RzX5EW5l13VpbuUmWMMxLE7t3djn9Pxr6k/b/wDim/iXxhb{+}HIZCqWzHehPIPQ8dsc/jn2r5lVlsdLVCuM5wq9QAf6n{+}QrklK7udsIW0Y3USbh/K3BTIwQAnJAyTj{+}Z/P1qVGWK4ZI8tIoChj6{+}ntgcn61k6WztI1w{+}FMZKgHpn1/T/x0eta9ko{+}0QwBd0svIcjnGev1z{+}tQ1y6HSnc1II44bcvKu/YQ7h{+}jkfcTHpnk/gK5i6vZta1bZE5ZASQxPUDO5z{+}R/Wr/AIhvhHvtV4jiDb3GSS3A6foPoKoaIqNdRjaEUphwpwCB2/kPwqqUftsU5X91Hp/hrSI7fTEWU{+}TC37{+}WQdflHG714PT1PNd5aF7mSzt0jCM/z7JQSsY5Blf1AHQd2PoK4vw3Zi48tZ8yHPnyI3ORnCp7kkDPsM9q9W/sxNE0jN4yte3eJJnfk8dIxj{+}FQQT7kDjnHLXqcqsjqpQTZia5K9zeDbIsdsibVct/AMZ5PHfJOeSfpXl8NxFrPieWRHjunt03eaykW9v/AA5A79uT78ei{+}MvFU{+}r3b2sUotNNjIR3UndKQfuqP4jnjjjJPetLwjo/9pvBbWtr9ntY2Hm4OTJJ15PfHT0rD{+}HDmludcI{+}1mox2O5{+}HHhh9Q1DzCWmVOBLL1b39AK{+}mfCeipFHFvUjHoOwrifh/4aTT4YwQB34xzXsei2mCmAOfWvArVeeWh9NRpqnCxtaZYZ2gDg{+}vauq0/TFBCf0qppkI{+}XAx3zXU2ap8vy59xSppM5q07C22mrHjIAbtkVp28A29CCeeadFEDgHhT3NXoIwPY/rXbCB5c6l0NitBgD1GPxp6W4Jx3X0qyqlTgHHHYf8A1qaSVbIYn/gNdNkc25lXtodvc4PUDrWTdWwAJIBz3xjmujmOd{+}TnJznpWfNHlCdoPsKwnFdDeEmcR4h8OQatYTQyxhgwxgjNfMHxN{+}FsvhyQyRZOnOfkYjIiJ42n0HPB/Cvr{+}86n5SMdK4vxdpseoW00bxhw4wVZeDXE5cp69Kbe58MOtxomobJUG3djaRkEeldVpepQz4eGVY5AOPM4Jx2J9vWtj4i{+}C10{+}SRYwRGW4BBO32{+}leewq9vdhQ/lygbtjnhse/Q962upq5clZnTXNibmbz4jh3fDJ0VyB3A6N7/wAxxXmvxF{+}FhntLnVtLikuEibdc2apiVBn5mX1x179OODgel6Jdx3QVCgYYx5bD/DuMnB{+}vrVu6aW3mbyn35IVZJGzvJ/hb0PbPQ9{+}lXSqTpSvFnFXpRqKzPluxhMskmnSTpdXMMe{+}2lXjzEPVWHX1/XuK6vwJqPm2MmnSTGONX3R7yN0bA9j25A/n0Jrc{+}K/ws/tCy/t7w/G8GoWcjma0yAxxyQPQ9ePf8K8vtNXF2kd/GzRurCO4{+}UKVYcK2Mcccfh7V7sZRxELx/4Y8CUZUZ8r/4c9C8TWBsrldT8s2tu58u52DKA9N3t1B{+}hPpXBatZx29zdWs0fnQTJ91TlUcdDkdiOM/4V6p4c1KHVtJmtLlRcQyIY51K8{+}XyA{+}PUZ56dx3zXmusaHLbzXWnXO5bq0AeGUjl4{+}2fXGcVVKTTs9yKsdLrqec3vm2UysVYXVueGYcnb2x9MH867PwF42uPDviGw1OxmMMiOJI8How6r{+}v41zV/5sllKUZhdWLfvQWJLpnhvyPP0rM0{+}4MMrwhsrIQ8LHqrfw5/kfpXfOPPHTc4U{+}SVmfuX8D/ihZ/FrwDp2tWzfvmjCzxk8q4yD{+}oNeg7D/AI1{+}bn/BPj40/wDCO{+}JZfDGoTbLLUGLorH/VzY5AHoQuePQe1fpMo6f0pU5XRzVYcktNmRtgUx04qYqfSmsM{+}1amJWYYqNgBz2qy4x2qJlA4oAgZMn1FMOM4zipSKYy8ZoGMpCMjilz8tJyBQIaRk4zzSbeOTTxTTyTQAwgmmnr0qTPHFM780ANPJxzTCO1OYfNxSkZIxQBDjPr{+}dFObg0UAXR3p3WjHFABz34oLFH6U4dPajHOaXrQSwAzSYxTl6UEc0CFIzQenFFA5NABjJxR6igjvnFOA5NAxuTjHpSjtijaO1OA44HPvQA1Rl2/AVynxJ8UW/hjw1qeoTsFhsoGlZh1L/wAIH4/0rpp3a3BZnVMjOdvT3618s/tx{+}OY/DHwvfTVm2X{+}qyeXGpfBVBn5uMenP19qznLlia043lY/O3xXqh8ZePNX1u4n3hpDtLryx4zjn0/rXLa3eF7gwLxghMKeP89KvXjrYwNKuCiqdnPOff3/xrE0ovNNLduNywkY6nc56D{+}v5VhFX959DtdlojV08pEuJvltLdC0sndiecfXoKvQ3r2di{+}ozp5M1xwuOPLjA4A/z6VmQI1xdQaefubt0vQ5x1J{+}nP61T8RaodRutkRKxkiOJR2UHrRy8zsNNJXKtxqPnyqWTKK/HvgevfGQPrn0rpvA9k19fidyVtk{+}dwe6AgKo{+}rY49jXGRvv3JGflYhEz6ZwPzPNe0{+}BNHisbaCViGtY9rFc4DPg{+}WGPbu59M1rVkqcLImknOR6T4H0hUeW4vSimIG7ubh/uI38IOOygZx3A9a4r4nfEue7ikaGQ2kdwfLhGzEhjGeSeuSSWOO7AZrZ8QajcSQ2{+}mK7Q2UuLm4QHAkHBUHvg9eey9q8Z8Q6o3iHxPdTxLmKF9sSgcEg/ljPP1{+}leVTXtJ80tj037kbLqXfDVlPq2qxR4LyggMx52DuB74/XNfTXwx8IrYRI5jy45JP{+}NeW/CPwi8I8yYZkYhiSPXqa{+}ovDWjCOGEBFUYAPoK8/GVrtpHt4OjyRUnudV4bslVUyAO3vmu409MbeelY2mWarGu304rpLCHcwGMfSvFPUeiOi0wYQNyMnp2rpdOBKgE9{+}lc/p8fK9OK3bEfvABzg9q6aZ51bU3Lc4yOoHpWhbkuAMD6ZrOhBrRtcg/Nge/pXowZ5MiztDHbnlenFRyoRxjFWEIPfb25FMc9eRnHGe1b2IuUJTtJUgY9CKozkkHaQBjk1emDOenI9sVWkVQp56c1zTuaRMa6UnIyD71z{+}qW6ujbs4HPFdPcgKT8oz2FYWpIHJHBAPB7V580elSep498QPDy39tONgzj8a{+}cvFGlNaXh3pvZeq9M{+}49DX1/rdmLmORWHzEZIFeGePfDIM0rhMEDOQP6VMJcr1PRa5onkmkzpbzmdHKxnCukg6fl6V0l7bGOL7QoM0UqAMQcq49R2BH68VzF5byWt/tXCy9m6A{+}x/x/A1qaZfm3YDZiFwVntie3qp7HrXZ5nI/MqvdPayBpJhLFgAXJU7Ux/BL3Hs3OPUjpwfjz4dRia517SrIlpFKajp0WBlT/y0XqMd8jofY5HW{+}Kr1tCn80fvYJV4mjO3cnqfQj/HrWT4d8app99FZ38i/eCw3cQwkiHGEdB3x/d4PUDNddPmh70Tzqyi/dkebeGtYk0S7khQN5nG0PjLD0xj0x{+}B{+}tdX4g0{+}HxBpFvq2nIfNg/wBYrEk5zhlb/H3GfbovGXw{+}s9TZbqz22V5GytCRwkyEdOB94Z7dR0HauQ0nULnQdSZbyIxwSkpcRn{+}Ejgke3t6Gu7nVRc8d{+}xwOHI{+}SR5Z4jgFhrcV4U2Wsq{+}TIPUHIyf1H/AfeuOu4JbKaRQpzA5Ug9QOeDXuPjTwzHdyywxphbgFADxtk7Z9iRXj{+}rRMrRzujbwnlTKeSxXjP124/HNelRqcyPMqwcWdD4I8TXXh/WNO1qxl8u4tpkkzz1ByG{+}nSv2l{+}A/wATLX4rfDTSdbt2xKYlinj/ALsgGCPzB/Kvw2sZTpl2Yg48mb5onHTOOPwNfb//AATx{+}Nn9g{+}Mh4WvJfLsdWby9hPyxzj7pA7A4x{+}J9KbThO/Qya9pC3VH6XD86jZRzTweuO9DAHg1ucJXfnOahYcVYYcc1E6jg0AQd6aRUpT2qJhjFAEZ44700g5J/nUuKZjGaBjeaaaec5xTWHFAhMH60w49KkPQUzoTzxQAzmmnkelPb0FMPANADD164op23Pc0UAXQNvvTvxpBgGlHIoGw4{+}tH4Y96XHFHJNAWFXpQTg07OBTQc0CFpTx070AUoGTzxQMbtpwOOtLx70UAA4Jp4GB/OmjNK54x{+}H1oApXWZDz9xefXLdvy5r83/ANv7xhHr3xG0rQopSy2sWWx6uwH/ALL{+}vvX6Ma7OLWwmYcYQkMPXBH8zX5AfH7xavir42eLdVMvmwW0rRwuoOMINi8j6fpXNV10Oql1PIfE168dxLArgrDkluuewH86dCV07SoYyMMkW92zgmR84H4L/AD9qoRBJp42uQ2xm86UZwSB0UfoPxqfc2ralHG4AijO99vTJPIH8hWlkrI0WrbL9pKdO0xZOl3dfNhOqxj/Ej/Oa5ydiWMhYD5SoVPQcH{+}ePxNaGs3pmI2/KWACqPuov{+}HU/lWVdsIo9qAlioA7kAHp9SeT9KqmurFOVnY1PDWmnWdegg{+}7HGMu2MbfX9P1xXtKyLAltbbwttbKTIQAMkcuxHucL9Aa4z4caT/ZmnSXk2PtUgJPmDIBz/QAH8BWrJqitCyAbUYfaH3HJKg4RT{+}Iz9Aa4sQ{+}d2R24ePKrsb4i1{+}VbC/u5HdZrk5LFumcjH5cD3D{+}tYPgfSvtl8jOCF35Ix3/z{+}uaxvFGrSXd3bWmclSJCAeenGfpyfqTXpfwu00nY/TGKxmvZ0vU6aT9rV9D3f4e6LHbpGCi4wOnP5ivctB08Ki/L6celeZ{+}CrTaIwmMkD/61exaRbiOMZPK8mvmKmsj6mm9LG7YW2FUBc{+}tdDp9rgZwM4xms7TojjcQM9jmugtVVSuORnrWSi7lyloXrW3BbhcnpW9Zw7MYJH1rOsTlgwH4VtWq85PU11U4nmVZdy/GmUzjPfHpV{+}2VVBByOwPvUEChTu4GOlTySgdRjjk9q9CKsjzm9SUkKuRz7U2QkKMrweetV/NyAc8fXFKsxdfr0HaquhDpkLR54OOKoTLsJAH5VqxspGGOSPeqtwiqDzyex61E1oEGYN0jEFdvOcfSsS8gYghT19K6WdFk65YfpWdeQbV6Dj9K82cT0KcrHHX1qz7gOevHtXEeJdHFwjAISTznGcGvTLy3XJIHQc81z{+}qxxMCvR8YAzWDielCdz5b8beGdk8sgXHPp0NcpCqTt9nmkENwDiOZiNrdtrH8sH/I9/8YaVBLDKoHznPLelfPvjNP7PuJCFKjJ7c12UrtWMqndGZ4gmkhhmsrhAksZKgSchTjkH1B/{+}uK8Z1i6bTYpFt0W6tg3{+}okHMZz8yA/qD9etep/21BrGn/YNRl8iTbsgvmGQnor46p78kfpXinjye/wDD2s3AmDQTxBRMo5DofuuMcMMHqOoxXs4aF3Y8PFzsrnsvw88e2V1ZC2vLnzrKSQKryjMsTntIOTg56/jWz4q8JfaZbeGeSNRKuy2vtwAfP3VfsT6H8Oh4{+}abbxFc6JfC/twsto6jfGTnC91LDkjoQR0H6/RXw58f6d4j0lNOuiJbKX5GSUhmQkZ4Pcdx69RzkUVaMqEvaQ2MaVWNePJLRnP6ppt3PaSW08Zt9YszskV1/1i44Pv0/T1ryzx3poJkvYo9pmkJljByPM9c9sjOD6nHavoLXLA6XdQWVzcvcLtBsNRcjdKnaJz6jsx9gexrzrx1pELKbqB0WKdtssSrkRuBggr6Hr7EVrRmk01szOtC8XfdHhBj8{+}F7ZfnaP54gfvAdcfnXSeEPEV94b1DTtZ06aSK8tpQ5ZTkq6nIOPfFYGoQ/Y710bK{+}UxKleSq9156gH9Kl0{+}4VLuRSA6SDehJ{+}Vu{+}PavVkuaOh5a92Wp{+}53wQ{+}Jlp8W/hjoXia2kV2u4F89V/gmHDr{+}efwIruzyK/N//AIJsfFx9D8Yal4Hvrgta6oguLJd{+}VEoXJGOxKn/x0/3a/SHrRF3RzVFaQxhnrUDjFWTzUUg61ZkVjk54IpjjIqVziozntzQBF6UEDmnNnpTSOOtAxjKM5/lSdqfjBprcnFAiPGTSMB0708jA4OaafWgBoXrTSBTj9aRuhOPxFADePeijBPQ0UAXdmT1pyrkA0DkU4crQAhXikA45OB604dKOhoGJtBAo28{+}1OpV4GaBCgYo28/Wj9KUjFAARim45zTqCMUDFAAFNKFjkt8vbHFPAwRS7cnHSgex518dNf/4Rr4c61qX3RBay49MhcgfiRivxl1e9ZLXUZrjc0kr7nOODznHvk/54r9X/ANszUJk{+}GaaTbz{+}VJfySeZwP9TGhkfr7AfnX5DeKrh7nUYrASZO7c2Octk/y5/Oud{+}9UsdULqnfuUELR6Y9zKczXDfL7KPu/hnn8BVyCBreykU9SvmynuMdB{+}oH1NQn/AEu8SH/l3txwfp/jVu/mSJUzt2AhtpPDYPGfx5P4Cm9XYqOiuYt2s2/c3EsnzFfQHGB{+}PH4D3p{+}k2bapqUNpCC2Xy7E4JUfyzj9arSztO0sxOR97J7nt{+}pJ/Cup8C6U0kMk7Hy2u2Mfmdo4EGZH{+}nb8K2k{+}WNyIrmlY71bqJLGC0TbGHjMrsvQR8KPzGSPYisi9dPLnaUFIYV{+}0TYHOMDav/AHzj/vqrFuol3zMNyyP58qkcLCmFjj/4EcLXM{+}NdTdbH7PkCS8ZpHI6lQe/1Yn8q86KcpWR6UnywuctpkrXusee5JaR9x74Hp{+}HT8K{+}h/ht5cEK546Ee9eBeF7IyXanA4PevZNK1tND00zsocxrjFPFrmSSKwTcPeZ9J{+}HPE1jo9m1zcTxQxR9Xc4x/9ar7/ALQehxOsNrcC7JydyMME{+}vp17V8z{+}DvB3ir4wXbAyy2ml5wXK5XHoBnmvb9I/ZXgt7ZAuqToMcqUAO73rwqtKlTdpPU{+}gpVJyV0tDv8ATv2krMZXyTtXjeTj{+}fb3rc0n9pvSpJf37BFH90jGPrXiOpfs66rayPHDrSyoW3bZY{+}3PA5PFcnrfwm1zRy8kcEUg2/ftm4/75NZxpUpbSNXVkt4n21pfx78NXKDGoRI56DeM/wA66vTfi7ok08Fs17Eks4zEAw5HtX5gatBq1kwRrZuBydmCPetHw/491LSVtUeV0KS7lMhyY/Xb7Edq1eGlFc0ZGMalObtONj9XrTxhZzFSlwrB{+}hBrZTUFdMqwI68dMV{+}eHh/4n397aRNHdSQ3FpKzkhuo2kqPzJ/Ovp/4X/EObW1YSS74oykfB6sUyawU5rSQTw8HrBnuougFbIx6YphnUcFvmPYVjwX4kAII2kdzTzccZyTWnOcfIai3hVD0z2I7VXvNRWHl2JwKyZ9SjhjYkkKBXi/xQ{+}LUml2GptbTqpi2pGP9rqx/AHNZzm7aG1OjzO56vf8AjfT9OSSWWZdqIWyWx0OOteaeIv2kfDumoUWUyztnaic8Y6mvlvxp8TNQ1XyUxJJCA5EeTh8scZ9hmvO7yHVrx/NKvI55d1B4P1ohTcviZ1uMKeyufVmsftS6aZvKjGwEbhLJwP5/zrkdX/aVimjLJKjY4JA6{+}uOa8EtPCerajKzPGuc5LMMD0x2rprD4Oz6haq897hepQPgY9MVo6NJatihOo9FE0PEf7RVxPGxjYnGQWx09s9D{+}IryPxD8Zbu6ldLyEupyQ6YGfqK9zsPgro0UBF3tkcKPufMSfxzWH4q{+}EHhZIJJEsgz4yWHBJ61tSqUYu3KZ1oVZK6keMReJbfVLRngkBweUJ5XHNcneeMbZ5F0rXkabTAxEF1Goaez3ddoP34z3jPHUgg9eg1TwkNAvLh7ffHbsc7G7c15f4qjLuT0y3AFe7h4QbvE{+}cxU5qOu5sXmh3mgXFvYOYpbe5UyabeRnfDcA9EDHqrcjB5B64OcaPhXxB/YNxE8B/cE{+}VPayEhocuflbIHGeh7Z9RWR4P8app2iz6LrMDan4blmBMAOJbZyP9dAx{+}64HUdGwAexG94j8MyALr{+}m3C6pazZzeRqdl4h4YOp5SUDlkPXkjPU9M0vhkcEJfaifQXg3xfp/irS49J1QsbKVfLjncnzLWbnhiO3r9c8jNZ2veDJNDvLnTLwq77Pklydl1FxjH{+}2vp6YxXinhzxNJ4d1E3lu73FnJhL21J5UDOGz7ckN1BB9efofw94g03xtpSabqFyJHKBrS8I547exxwR0IHHSvGq03RldbHs0pqsrPc{+}Z/HGkyWV65cYlgwrYGdyY4auQdntrcGNmSSF9yt7ckD{+}f5ivor4i{+}DZroTwXEZXULVSY5ccXMWeeR1I4P5nvXz7qeny2zMjqSCMBlPyt6EH8sj1Ferh6inFHm16bg2d78M/F194O1/SvFujyGO{+}0m6iuPLBIG0H5gfYjI/Gv24{+}H/jC08feDNI8QWTrJb39skwKnjJHIr8GfAGopb3yJL{+}9hBKvHjllPDAD6E1{+}m/wDwTs{+}IFx/Z/in4f6ldGSTRrnzrISHloWJBI9uFP/A6v4ZWOaouaPMuh9oA802ReOtPApGGe9anKVXT3/SoHBz6VbcYqF0GKAICOlN2g1Ifu4pjDbQAxh/k0hGehB{+}lPPrSDjNAEW3H9aQjAzUhGabigCNlzg03aR1wam6A96YTjjigCMjBop3HrRQBcC96djNIBxQBtNADsbRTR0p45FMHzGgBwHrS0uOBigryTQAEYpRnJpBkmgDPfGKBgMk9qcBnpR0pVHegLCjrUgH40wLT5WWCB5G6KpY0CPjr9t3xMtqlxIJSkdpZTwAZG0yPER{+}vmKf{+}Ae1fllDMTLeahJ8zFisYJzn3H4CvuH9v/wAQtLp0UUblrjUbpiCp4x2B99pXp0r4oCBruCyjVTHApyf7x68/p{+}lc0d5M7baJFi0AstPd3UAv8xOORx/nFYup3b3ThAv71xyqn7o6Bava1fZ2oW3KhGAP4j9P1/KqtgRb3AZ0zKTvO4cgjOB69iTW1NfafUJv7KFNksk0NnCOXbGD2xx/Pcfyr0rS9Mit7G4CtiIFLKMDoFX5pG9{+}ef8AgVcr4U08u82oSKGVPliz/EBx/wCPHA/4FXbTY03S47edVRo/lZT/ABsSGb/2VfoDXNiJ9EdVCF9WU47nLOrq0ayMJ3B6qACI1/4Cpz9WPpXnuu6gl7qN3IOVj2pEvoqtXbeI5DZaPeXDkm425ye7E9fzx{+}FedW6mS9ZMYG3kfgP60YZXvJixTatBHc{+}ENN82WNgBufjnPWvT7Xw0lykVvKNyD7w65Nc74K0tsW{+}AACufxNejW9lPFMGK5HHzY4FcWIqe9Y78NT91NnqPgrVrfw/osVvboqR7QRhsEGte7{+}Kj2seEZiADkjtzXm7LILZWYgNymOueK878e{+}I20iORQVkcpho2JBQ{+}p{+}leUqLqzsj23VjShzSPSfEvxzlsV2y3JAflAOp9sdc1iDxr4z8SRCWw0id4Tt2TTkRqwPAxvI6{+}1fNk3xGm026eaAC8u85W4mXlfYDpipLX4ua7BZG2k1No4925QTll{+}YHjuMHtXrLL5RWh5DzKnd3/AAPoi80vxpFF9pu/D01xEVGWgZJ8f8BUkkY9qTRIrDWFIktVWVThkI2sp9COoridI{+}OfjD4Ua7p8HiK1YfabCG6jhuwYzNDKm6OQZ6qy7SDXsPhnxVoPxqRbi1KaT4kjDHzGUKJFB4Vv7wx3rmrUalL4lodtDEUqy9x69mVbTw5bWUvmwoyKVwyq2MgHP9T{+}det/CjUV0y4GZMq0hkKg4{+}YjGf51gab4flv4DBLBsvI{+}CvY/T1FXNN0u40HUI90bBCeozmvLqao9KLWx9S6HqZmhQg7gRn1rZeYgEgn8RXC{+}CLrfbRAcZXr/APWr0Ge3H2ZWzg1jGTaMKqUZI5rX9RMFrJn7pHPavlr4izR6vqFzbhS4dzwnf8a98{+}IN{+}1raS7AT2ANeOaToE13dXF1LEGAJbcaUHd3Z0q0UeZa1YW{+}jWH2mdVhSNQFwvPsBXIRad4p1adHL23h3TirFJdTcl2UY5ES8jr3xXoHj3VbXw6v9t6ttKRNjTrAAl55Ox29/b8{+}wr50{+}OzfELSTZa1rljNpVtrCSTWqSMcqmV7dj04PavawtF1dTzMViVS3PUdRhm0fRby4ufHVjttyf3aWhDSEEjCjzM/wmuR0n45XtjYxI91Z3jEbUjDMsm7HU7hjA6de1eAaz4itNT8P6Ja2unXFrrNuZ/wC0b97wyR3hZy0ZWIr{+}7KqSp5O7rxXffDP4D6t8TPBWta/HM9uLCUQo3l5SQ7dxB9McfnXr/UqcV754izGrJ3gj2rSvjnFdtHFu8uZhgh/lJGTnHPPGOlbi{+}O4r8SLIxdgueTxmvi5b240W{+}ktmciW3crkMSFIPavXvh/rVzr0UcKK67Ttkkx1H8h/n61yV8DGC5obHXh8xlUfJNHX{+}MZBdRyGMncema8U8YWjLCNqgDd0x0r6ZufBEyaHNNcRqpI{+}Uda8J8baa0dkXQAlZME{+}h9qeDmk7GONhdXPLEcgqueA3{+}NdP8O/HNx4Q1KeJ4vt2lXYAubCRyqy46EEcq47MOQcVy7r{+}/dcYODgVJEuJ0ZRgMMg{+}{+}a9qSUotM8CLcZXR7J4i8JW0{+}mReKvDFw1/pU0ebiNT{+}9tHHDLIo6MAoPHB25HBIWHwN4nFni18zyFc7kCHgEc/L7jqB/jXIeA/HOo{+}C9du/sZR4J2AmtJhmGZc8q49CO4wRwQeK7DWdB0/ULCXXvDbMmnlwJ9Omb99p82fuNgcocHaw7Z6HiuKcbLllquh6NKet1ue9ad4gtfH{+}mQ2N4Fh1ZP9XKvG/0Knse/I788E48E{+}K/hC50DVJZola2DEjZj93v7gg9uvB6fiKuaJrjxpHKHdZY8bwpwR6MPQ9Qfp6EV69Z6npXxN8P/wBnaskct8YfKW4bP7wYxz3B/P8AGvOinh53Wx3ztWj5nyZYzompkmNrV26mE4AI56H/ABFfW37O/jSfwP8AFT4e{+}OVmH9l6oy6JqUrvjcx{+}Ukg8/daFs{+}qmvnnxz8NNU8EakIp4z5II8i5YZRx/CpPTPbNdf8ODLrvgrXtEEbNcWUy3tsyZDK6o{+}OnTK5H1A9K9OTUkpxPMUXG8Wft3BKJUByDkcEd6krzD9nP4hL8Tvg34Z14tm7e3W3u17pOg2uD{+}INeng5UGtU7q5wtWdiJ1zULrirJHNQyjOfSmIrOM800rtGKldeKiZc8mgBgxikJHPFLjHSkzz7UAIeOlNJ6045NJgYwKAGbhTR605hg0YyfSgBpGaKBRQBcAOef0pVGG9aXGTShe2aBiGkwO1OCYGRSjp70DsGMClHSjpzQBk0AHSk6dqdjn3oPFAhtPB6UADNO2jOfagdxyjJFUvEF19k0a8l6eXE8n/fIzWhGAa5j4nX0WmeB9ZnmdooxbOpdDhhuwuR{+}dJ6IErux{+}Vf7aXib{+}0PHlpaAh49Pt1IiJOPMdVA6/7Cxn6mvnSCM6fFI0jhZGy7leu3/6/QfhXonxo11/EvxK17UJ2BSOZs8lgNvygA{+}wGK8u1a5NzMIgQjMfMkJ7emfoP61zwXMrHe9DPuJ2lPm7CshJ8tf7vv8Ah/OrlhZPNN5MeQ4UDJ7Z6n8ADTbKEbHvXxsU7Ioz39P15rY8P2sjeWQMT3B3l84wucL{+}Z/RT610SkkjKKuzstBs1jt7eFFBLuqqoHPBwPbGTn6LTb26OpeIBCMyRROzMynO585c59ug{+}lWhcGxsrm7UbjBAVi9CWGxePYEt{+}NZ{+}nW40qyEsnySMfvN2z0/Xn8DXlP3merH3YmH4yvNzG2U5jjIDY5Hf/AAz9GrnNN/dnzGXLN8oHrzz/AEq9qEhurbzOS0sjSbcc44Cj8sflVO1njaSOLhjGSfzr06UeWFkeVVblO59B{+}ArNLhLUrwVUD3H5V63a{+}HhJDuYYLA5wK8t{+}F6gQRDIII3Zz/KvoPw5brcQRcfK2DkjIr5vEycZM{+}qwqUoI8y8T6XqmmaaDDGGfaSGYZB9OPX/Oe1fK3xM1vVL3UpFuwdn9xVOM1{+}mNt4Gg1jT3jOF3Z5A5NeV63{+}ybY/bJr1R9scncFmbPPsOmK1wmJhTd5rUwxmGqVlaMrHx5{+}z78MYPiF44s4NXYWmlKd8rE/MeeByeM88039oL4Mz/C/4g6nZwqZdImkM9jdp8yyRtyBn1HQivoi9{+}BENtqEkpiaxljPyxxjaMjr09au6d4d13Sba5M2ofb9ibY4byISDcxGMA9B2r0Xjru6OOOWw5bPc{+}VtG{+}FmpeKNc8O2VpctqsmpxRhGSOULB820ozOoHy45Iyo9a{+}8fi18BPhto3hqx/srUzpfiqytUijvtOlOJZFQKN6jIOeeeDzyavaRreqmyfTmjsEit0jMMlvZqjAk/OFx0BwDgY613{+}jaZea2TbWcIED55EYOM47446frXNWxspbI66OAp0le55f8AAXU9W1LU18P{+}JTFevGFEWsWYLRsuBjfkAg9unb2r2jxd4Mh02VogQwRRhwOCCM10{+}i{+}GLjTY1Q7YxxnYAPxJrP8AElv5aspcsRwSSTxXj1He8rWuepFpyST0Kvw{+}R4jEnUg4r2PyC2m9BkjvXk3gtBHcKvbPFexI/wDoW3GBjuKwpJNNkYvRxseL/EK0Z4WUsFG7qKr{+}DLC1GgXEcsAeRiQqg/fOelbnjVFeTGOlV/BjiCXY4G0t37VNO17HRO/KpHimqfBDxNdfElvFM0VhfSwbXsYbpi0dqMHICA4yT368CoPi1oV78W/DsPh3xn9mtoYJs2l5p9m/mK{+}DnbkkbTwDn8uK{+}rrvwpBd2oMJKSEfwniuE1v4UanJIJLO8STBzsfIAz2r1VOpBKzOGM6NR{+}{+}kfC{+}r/s5eDl0TTdLbxVcxQ21zKzQ/Y0MpmaRY8FxGGCnCEBieORjNa1l8Tb3wR4KTwL4R0WaHTBuZ7kplp2P3mY45J4H4V9H6/wDCnWt1wr6RFcM0iyblA5KkEHP4CsCb4QeIb1P3enxWqhyQTx15P86uWKk9JBHD046xSPkKb4C3firUFm1OGG3u5cN{+}4QIGVu{+}ABgjNerfDX9nvVPDN4kVxAjxxH5JlQDcOx{+}vFfT3gr4JxaUI7zUXWe6RcbQuQv0zXZX9jHaxbI41AXoQKwrY2o48t9Ap4WmqnP1Pn7xV4aaz0aZGbLBc7VH9a{+}QPiJaoINQiQEFZAy{+}5zX3j8RLcSafNxksp5xjtXwd8WZGsptQyTjrg96vL5OUzLMIcsDxS8/dXc0p5xweO9RbCohA{+}7nH4Hmn3z7tw5wTuPqeMUkZZ7Xy85aMh{+}B2P/AOqvrNj461mSW7O2pRiRmIZghwexGK6Xw/4yutM/fROokZQkiOMrKO6sO4OK567gaO4icDAYA/Qgkf0FRthJpF4TLMP54qJRU1qaqTi9D27Q9LsvENk2qeHZR9pQZutKmGZIfof44z0z1BAz6nFg1WXSLxJ7csqk8xk/MPUE{+}ua4rQtZntLuK7tHkguYujqcEjH{+}f0rp5PFtl4pUxaggstTQbWuoVxFMo4yyDo3TJXj25rilTab6o7oVE12Z7t4d8TaZ450BtP1tUuoFwFyPnx3Iz355B4P1wayNI{+}F7{+}FfFEj6dOb/Rr{+}IwLKjlXH8SxuOueCozk/NzkYNeTWV1c6PPHuPyOMrLG{+}6OQeoI4Nd/4b{+}JUun3cYmV5Ivly6H515GCR0ZfryPWuVc1N{+}7sbu01rufZH/BPTxulveeOPAlxP{+}8guRqNvGScgN8kuPYOoP419qo24Z4GRz7GvzS{+}COp2ngX41aT4qV/s4nzb3qqcK8cnBY59CQ34c{+}tfoxousW2tW3nWkyTxN8ysjZH{+}c5rupTUkeXXpuEr9GazCo5ORT92QT2xmmtjFbnMV26VExxUz8VGQCaAItvY{+}tNPpUjcGmnkUAM74zSYxTzxSMuaCugx{+}lNx0qQjAHFMIxQIjKk0U8UUAXOQad060Y3fhS47nrQA09KXGcUqrilNA7iEUDrSkZFAGTQICM0pHHTNGDnpQM4OKBAFxT1GaaoJ5JxUqDB{+}tAD0Tj0rxH9rnxoPBHwl1TUvMjX7Oo2RuM{+}bMzqI0x9ckj0r22SUQxM56KM1{+}dX/BR7x9c3es{+}HvDCTlbeCJ9QuVUnDOxKqD2OArY9mFZzdomtJXkfDmtXOBI7vvd28yRs53HsPfkk1zEcbXTkA5knbbnP8I5Zv5flV7WLtrq4MSDdjksD/n6fhVcH7NCFH/Hw3y4AxtX/E5pwXKjplq9C1HAl9qsFlbuVtVwC/8AdUfebP5muy0JISZJpEIVF7D7oIwB9dvH1NcroNs0FrJIB89yfLHHOwHnntkgCurs3/s6yDFQ7oPNck/ekb7i{+}/qfxrCq9LI3prW7LdqH1q8mtpBtgilDTkdAcEsPwAQfgaxvGmpedc{+}UiCN5JAEwOkYAwfyH6muk0{+}EaN4Sdjhbi6c5kfqAcFj{+}QP51xkLNqutC4fO2NS3TPc/0Fc9JXk5dEdFRtR5erKes/6K2xcnCCMgfQk/qSPwrB0{+}I/2jcnH3QVA9MHj{+}VdDLFLeMWADsTuIA5HI/mafBpzWd5cWdzbSQSXSCWF3UjcccgV6CkkrHn8jk7nrXwq1MS29uW4{+}RRwTX094Pl32cXqBzxXxr8LNQNtLHHwArFSv4//AF6{+}vPh5dq9vGzHtng/0r53HxtJ2PpcBLmikz3DwvP5UCq344ru7TSkvVQMoYNxg1514cuUOD1B9q9K0G72bOjY7GvGhLXU9mrD3bopar8OLS/P7yBWHXLCsNvgjpEzgtHIpLZwGOK9at51lXgZz6rTwcnlO3QivQSR5bqS2PPdK{+}E{+}iae6EWXm46eaxau2ttMhs4QkUSKFHG0cYq8WAPIA47VUu7ohDjA/Gm5qKM/em9TP1LbFCVx8x/KvOdfuPMkKjr7Guy1e5PkOK8/1IFmZs4brXnVKrbPUoUrK5q{+}EWLagpGQcjr6V66rZtiD6da8j8ILtvUbtmvVN5EJOcZH4VdHRMyxavJHnXi4hpSpOeT3rO0WRVZVONx9a1fFg/f5xuHNY{+}lMu/PcGsW7SOuKvBI9H0TUf3SwyAHb0Nb8YEwX5u9cPpkp3r0Jrr7CQbBxwe1dtOr0Z5Veiou6Jzp4lZhwT05702XR1UcoOnOec1dUgjj9R0oduDjgY6dq69DjvI5y/gjhjbC4PqK4zXMKGIOSOK7vVDtVgSB7CuB8RH5HPTArzKzPXwybPJfH16rWsynBwD{+}VfCPx4Gye7kLbg5HQdDmvtPx5diKO4JOMA18h{+}PNGHibxFFZld4dtxz0wP8iu/LXyzuzLM4c0LI8JudIm/s03QHyIwGB15FZ8an7Sq5IDDyiPfHFeqah4McM0Cq/wBmI3gDpuAPHv1JrzO5jEN7OD821xyOMEHivq4VFO9j5CvRcLM07u2MiWbEkZGTn9f1FZE{+}8yuc4bllz{+}orsJ7VJtLPy4IOQe21uf55rCvLQRgFVCqATuJ/A/y/WnGRnOJBbOYg7HKrImcemTj9MilnMtu5ZsxyZOG/ut/gahlBhgXGFcAqQepCn/ACr0UkV6j78kqu18enQH8OlN9yYvobXhzX5ba3YDY8IbL20g3Kjdx7exrpLObTtUEnlTGwlI4jkBKKcY69QPrke4rzdJJNOu84O5PlkA/iT1{+}orVFyWmDqSFIDLs4Izzx/ntiuadO{+}qOqE9LM{+}qLG21WXwxoWsXkqW32J7eKSUEMJImAUg9iBtGPr2r7Y8I{+}D9Y0DQ7XWPC2peVNPCjz2n3oQ{+}AWBTkgZzyv5V8IfBTxpDrnwt8X{+}H7x0kuba3S7tw6hmkwy5IB46HrX6L/ALUU8S/DLSb0skhyV86PgqTgnp23FhWMI80rPcVV8sdNjU8P/F5Y1W08SWk2kXI{+}X7Uyb7WQ9Mq4PA9j0rvbLUor23WWGaKaMj78TAg1hat4ZTUfMWUZcjByudw9GHRvyzXAHwpq3hOeaTQ7y5snI3/AGdF8y2bPcow{+}X3/AJ11Lmj5nE1GWx7AJlkfA59{+}tK3B4/GvJbT423ehv5PifRZFVeTfaYnmKQOpMWS/HtkV3/hrxpofjO0NzomqW2oxg4YQSAsh9GXqp9iKpSTIcWtzYYd/0NNxxShge/FB9qokacAdKbj3okcKORRkNgj0zQMD0ph46085prA0DQ00UHAooAvKDgnPNKeB70oGBjtRj05oEGDRg96cBQTQA0gnpSgYoHt0p2CKBCEY75oAJpccetOGFwKADbgDJp{+}KTGe9A{+}c47Dqf6UAZevXht7E7CBI5Cx56fU{+}w6n2Ffjd{+}1x47/wCEr{+}L3ie8hkb7Olz9mhDHHyRgRrwenCZPua/WD4x{+}Jj4Z8G{+}I9Y{+}UJpulyTLvOAZHBVOfoCce4r8OfF{+}pvq2p3kzsSXm3c8nnP/wBYVlL3pJHVS92LZT01YvmkkywHLE/y/H/Glt4jcy5kJXgzSPn7q/54/GqZcrGLdW{+}VW3N/tNV{+}5L28ENtGMXE{+}Cxzx6gfQcmtHozRbaG3pTGZycHyeU{+}XsMc4{+}gIX65rdjt/t1/HZgBlQiWQHpu6Bf5D8/eq2mw2{+}m2ALHCxjHI{+}8{+}OB/NquW7/wBj209xI2ZD8zf7319QD{+}Z9q86b5nZHbTXLFXGeMNT84tZRHGwqi4PYE5P4k5{+}gFHhOxH791Ut8oB9yTt/lmuYtZpr{+}9uro/OpAX8CP/r/pXqXh7RxYaHdXIxkSRjPXACEkfma0a9nFRRMf3knI5zwZopv9dsbEqCZJMkgds9P517X8U/hTbatolk9sTBqVvtkRQOmMHP0rjvgtpQvvH2nNIAcqPlX3wP5mvqbxR4Ul1ad5bWGWf7NanzDEhYKgHJbA4H1rzsRWcaia6Hq4WknS95bnwro9tJo2tXUTIVKzbwPqP6V9MfC/XBJaw5wDjpmvGPH{+}jtp{+}vi52lDKuGwONw/8A1103w41v7PJFG7gZJzRXaq0uYeGXsazgz668OXu2NCG7jvXqGgXQPl7mOK8J8G6qJ1jweAAf0r2PQLoEJz2r5qV4s{+}nS5o2PTrCfai47jGRWoi5yd2SfauW0y6wnBzj0610VpP5ihhx7V2053R4tWnytkk6jZ2Hqc1mXQBQgt9e1bL7ZUPI4PT1rLv7lIYn3EdM1cloRTvdHJ645gjYnhQOoNef6lfqznZjB44rX8aa{+}J7y3tI2ILtg47ism6topbiKNFztTLZ9a4XHU9iOkTrPBimQBgv516QsZe1wy544xXC{+}CIGTZwCenXrXqi6cTp3mZB4xgV2UYOSdjzsVPlkrnlniOPG7vweorjbXUPJu2jxjHWvRPEtsY5XBOfSvOL63jtdTRzht/H0rmnGzsd9N3jdHd{+}H5DMqlW3CuzsIyq4A6DnPGa8j0HXXstVFtkBW{+}ZR9K9S07VhOihm564FaU0upyYiLexvwNuPJ7enaklOMAH/wCvUazqIwUOB61XurgGIkHg98V3XsrHnKN2ZGpzHaSSRnvXCa9LiJskdCM5rsNUk3KSeQK4HxRMEgkweQCa82o9T3MPCyPC/iLcBVm9Rn6Gvl3WrmR/FTpASknKjafpX0J8TNWWIzHdtAJJJNfLtxqlqni61ur3zDaR3kbSLHnLJu{+}YDBBOQCOD37V6mAi7NnBmU17q8z2HXdAtdI{+}Ff9p3AjiESLKD6YBDfmCR9a{+}NRci6ubqdvuSOTjPPXIr0X4xfGS48QQXfhrToprHSYr2YmGUncE8xikfJOAowDyfrXmNhGJBICTg4UH0Pb{+}VfRYSi6cXKfU{+}Wx2JjWnGENl{+}Z6Ho5e90iJic5j2kAf3Sf8TVWezkHmYX5VOB754/{+}J/Or3gVy6ywHkQtv6c4xk/oWP4V0d14fNvNIuB/Eqg8g91H5hRVSfLJoyiueNzzK4sT9qEIyWZWwW6k7Tn9Mfkaz9NvHtpzOOUYlv97nDD9f1ruNR00W2qW0/wDyzDA9MY55H5Y/P2rjbK1DXN9bY5hYyL64HDAfhg/8Broi{+}ZHNKPLJFy{+}tj5aTxMWKD5Wb7zKPX3HSorWdEkVM/u2GVwfufT{+}f/wCqrUDhZjC5zG3Ksf4WzjP45AP4elQXFqLWRonQhJOVbvGfb{+}R9vpUrTQ0sdD4c1m40XVY7i0YJNDkMoOAyH7yn2IJ4/DtX6p/sG{+}K7PW/hVPaW8ikW9wAYgeY89m9/8K/JCBpHjBHEq/Jwev8Aj/8AXr6d/Y7{+}PM3wm8YwXl4z/wDCPai6W2pfN8sQJ{+}WQj2PX2z6CsWuWakU7zi4n62uD85A{+}ZAOvfrTJEinKEorqQfvDkd6i0{+}9jukjmikWWGZA8cikEOpGQQe/HenyjDoo67yD9Of8AGtzg2Me/8I6fqbFniKknIHBU/gQetcZq/wADtHuLxryyggs7zqLm2VoJwe2JFNeodDTSO/Sk0nuUpNHlmlab8QfChZUu7fxJagfLFfERyr7CRRzx3KmtWL4mNZgrrPh3VdKcZywi{+}0Rn3DJkkfgK7pRktnnnv9BSlAynPIPrQlbYfNfdHO2HjfRtSjVobh8MOFlt5IyfwZRWhZv5oLAEqSdoPOBnNXFtYlfIXB9uKkKimLQZjIpMc89qkwMUxhnpQIZtU9s0U4LRQUXMehpV{+}lL19qVQe9BAmCTQVFKcge9KVNACAcUEZ4pQPwpdvORQAijHUU7PrS4prkqOBmgA6nHJ96VuAqD{+}I4pyLtHr7{+}tRXUqwQXEzEDy0JyexxQB8i/8ABQH4hJoHwalsEbZLrt44IHXyocqB{+}JX{+}dfk0JjOtzIRlgQ2R{+}X8zn8a{+}3/8Agpd42km8T{+}G/DaHEdpp/mMAOrSZ/{+}Jz/AMDNfEFjEwJIGBgjd1yTUwV7yOpaJIktbdYoHnlIdA21Qf4mra0Czkvbx710JYnZH/U1nQWMuoXPlxoViViAFHP0H17mu5gtmsLOGHKpIy5B6bEH3m/PP61hVnbRHRTg27vZDoo96EKAViJVc/xMTkt{+}g/KuT8Ta4bu8isYSwjjyXx/E3{+}c/ma6XxDqK6HoysjbLqRMxR45TPQn3PX24rz3RVa51ZNxJJDEnqelKjDRzZVaeqgjvvDuisdOglKnY7gZ/Aj{+}q17BDo/2Lw/LFMjAO0koA4B2xjg/jXK{+}E7QzjSbYgBZXLALxuAZVH{+}favYfiFp66dokbbQNiTEeWvB/dZI/OuGrUbkkdtOHLBs8Ik8b3Xw21jS9UsYRcPvSPYxI3YzwPfgV9PL8dfEnhAT2sk{+}naHd61DaM9rfQx3iNaTRybp1kUlVCk7WB7n2r5R8ZaWb{+}x0e/d4ba2iu2dnnb5QFVWAx1P3vbrWj45i0vxtef8ACVWovNGiUpaRLoyMbdMhnAX7xyx3Ejd60p04T5ebS99SoVakU47rTQ7n4jatDrGmaVeoTHJcyMiBgFZyqA5IHADKyMCOCDx0rlNDvmsrqPL5IY89KZrXw91/S/Amj6he2OoTSLcW9xb6hf3Cp59oxcKsUbHccY5AyQCDgAjM01mUhhnSM7COx5/lWUYxjT5U7nQ6kpz5mrM{+}h/hnrqzrGpb/AD1xX0B4f1JPKQBsn1FfG3w91prWRMMwGc9f0r6T8H{+}IUeFCXyfr/KvCxFOzuj6XD1FOOp7bpl6AoG4qQAM10tnqgAx37{+}9eX6brIZVKnPrk10dnqu7aCeMdaxg7Dq01LU7aXVQkJC8Z/SuH8V{+}KVt0MaEszHA2nkml1LXBFbuQwAAz171z3h/SP{+}Eh1Rr26/wBTGcIhPBPcmtZSb0MIQjDVnDeKNZn0rXtOubsbIpiVBY8A9cf59K15/ElvvWeOUNlccVH{+}0h4IvtY8DySaLzqFqRLDgZOR2/EcV8O{+}IfiR8R7GP7P/AGVcWbR5DSJExJ7cZziuilQ9uvddhOryLmaP0P8ACHj2zSQR{+}epYH16/4V7Jb{+}LFudNUo48rGQQcV{+}QHhX46eL9HvV/tJJLlQwJ3x7HH4jGfxFfSHh/9pC7bR0kWf5duSJOCv1FbuhUw7tujJuniEm9LH1v4s8TQW1u7zyoo6jJ5NeTXfiyG{+}mEu9VTdkfSvlT4ofH/xR4iza6PuMmcB1TcR9B0/OuO8Nx/F/V5UFutxhjgvcEMPyNH1W65pySNo1FH3YK59t{+}FtYTxJ4z8uBxJ9mQyOf0r0htcGl3KxyybC33Q3pXkn7O3w71Pwfoc13rE/2nVrwhppW9B0UegFeoeI/D9tr{+}nNFOWjfqkiH5kbsRXBJpOyNZb6nYab4gMqKTJlenBzWg{+}oZQ4PHtXguheJ77w3q0mk6lICyf6uXBxIvY16Xaa6LmFSGBGOMU1J2IlSW6NjUrlWUgYznpXmvjPVFSKQcH8a6XV9UMUR55xzivGfiJ4iS1gndnAyCBz1rBpydjqptQVzw/4veI1ikmAfr0zXibTXOmeEde1qyFy19CiQxGBAyL5rFGMmeQNuSCOQ23p1roPG2rvrerzBSSgOOe9czqvibS/Di2umavC8mmX{+}BcBGYFQpxvGOSVDsR74r6bC03CFkrs{+}Vx1ZTm23ZHB/FC1Ex0XUNxa4uLRVnTYMIwA6uPvMWLk556daytI00yWtw6j5kdSD9Bk/zp2r{+}Jor3V91hbtFp/lxo9tLIXDsqAM3PTJBPtmuqj0QadpBMXMFwqXUUxP3kY7efcEEH6V7UW4QUZHzziqlSU4kXhi9TTNRikbAjkIQqeMkHGPxXIr2s6JBrOieeGybdwrY{+}8eMBvxwh/wCBV4beogWNQCPLYSk56qMH{+}te7eBb7{+}2dHkttw2Pb7HZQDuZAGXA9wAPqK5sRpaR2YfrE47xdoU0Vs8wiCRoUljZOcrjJH6kfhXkd6h03WZroANtYSYHRsjJH5E19LazHG{+}lwyFVEaLjGMbju6H6Ekf/WrwPxtpv2WKWNeAJdqhu646n8KdGd3YmvBJXXQqT2kcqL5LB18sHPXPyn9cc/WqyA30CxsQJMDaxPfHGfqP1zWlogNxoli4GZIt8fHcjBGfw4/CqcoWxu5G2sIgADGOu3g8e4Nap62M{+}lyrbYVWwQj4OQfUV3fwV8Qxaf4xitb{+}L7Tpd6ypPaE4EgY4K57HkkHscVxWrwqsquOVmAww7nsR/nvVe1kLR{+}dG5WWD5yR7Y4FaP3lYz2Z{+}u37Mnj5vDV9J8L9cuZXezhS70K7ujte5s2AO3P95MlSP9kjqDX0lvKNlv4m{+}8OnTivhXwWs/wAW/wBnvw94t8PFpfGXhlvtFk0bHcZVwZISe4fJwPfGea{+}wPhX8QbT4o{+}BNI8RWQVPtaAyxA/6twCGU{+}hDAj8KiL6GFRdTsTg80mPl7AUgjBGVyp9u/4UPvVSSoYD0rQyGJyT9T/h/Slb07UifKqgg5x6U5gce9AhvFNJFPIyvpSYwKAGDFNIwfapOnbNI3PtQBEMtRTiDn0ooGXjnj{+}VKB3pCOBTsYFAhCO5pRwT6UvajFACYJ6ilIxS4pSM0AJk4pHBK5AzjmnAYpPrQAqtxnqKxvFshi8PSIrhJLiSOBSTjl5FH8ia1m/dqT/A3H0rj/AIkaotk2iRFQ6m5e5cEZASKNmz{+}eKT0RUdz8lP23vEq{+}I/jv4klREnit5o7KJGDBkEaBSDj3/ka8GsbKa7fcpEUIwOOMD2H{+}etelfEqb/hMvGms6lcS{+}Q8moXRlYHOcynjP8RwMZ56Vmwiz0uAOY2lwQiA8mRh2Vfb3Nc0qvJHljuejGlzavYdpukw6Np63MygZOYomHzP7n2qNL9LWO51C5OeRiPs{+}DwP8AdHp3wPTmGW/lZmu72XIPJOenoormNW1N9Va4ABWGMqiL1Gc//WrKlScneRpUqKCsjP1C/uNf1mWWVy2ScA1N4Ti2{+}J7fK5T5sg/7uaTRrDzdUjVh97HGeevaul8HaUx8WWW1Mks3UdD5ecV2VJKMWl2OSEXOSbPZfCGjsdT8MwvwXt1fHTr82P1r1b4oIkWkxpEyLG0NwAM55IYMfwwRXG6Za{+}V4o0RIMFbSzWTHdsNtx{+}uK6f4s6lE0NiYWwTHcSMpGCcu4/mTXz7leaPeStE{+}UvjHG0dhZRM5yjmQKDxgjb/7IK5Dwh8S/Efgaa0fSdQaFbWWSeKKRA6K7p5bNgjqV49u2K7P4z82lg{+}D8yKvI93J/mK8kr6KhFSpJSR83iJSjVbizb1bxrrmu3kNxfancXDwqY4Q0hCxIWLFVA4UbiTgdya9/8LTLqmjQncW82MMD35GRXzKRmvbfg7ra3GkrBISZIWCA57DkfpiscXSXs7xWx1YCq3Uak9z03RIDDcqpUqvRW9a9f8M6o9qEVW{+}UY4Y15lpsYnkJH3l6e9d9olqWVfmHGCD6818zV1PrqMuVaHrWka3IUU7wxI7966iz1t1UAna1cBoSOUAVQR0NdxaaXIbHeAMerda81xs9D0FUbRX1bX57u5hs4fnmmfaiqfzJ9gOtd9oO3SrSKLeWxwW9T3NeAaz8QLL4eXl/rOu/6LaL8ltI4O3aOuD6lqbof7VHh7VYI7mGUTws{+}DIrcD6{+}n41v7GclzRWhKam{+}W{+}p9MmdLxSHQFD1BGaybrwJpWqKTPZRAn1Uc15HaftI6fc4W2EZOepYE1sW37QB6BIHBGRk1LpSR2RwlWS91Frxz{+}zDoeu6c89vaxxSHoyKOvSvCp/2S9UXU2SOaTyC33QK{+}jLD9pBYo/IvLC2miRuhJDfpXT2vxr8LX9vM4011uEUOVWUYIP8QzyfyrsgpJaMwngcTHVwv9x5J4A/ZetNGiWdovNmQZJYf416lpXgm002IIsKIRx06GsHU/2hJDHIlhZ21qhwpcDcw69ycZrjNT{+}NF9Er5vURzk/KAa56kHJ9ztpYCvZ3SR7U1ktvEFU9{+}lUL6d4FAwCOtfN2pftJXNoxT7WzBcbioBI5xXnHjH9sOfRonaeR3kb/V26gGR/wAOw96qGEqz{+}FHLXo/V05VJaH0h8So7TWLAkP5F/bEvFKvUex9Qax/APi{+}4urbyJWcTIdpGODXn/wCzjF4w{+}Ol5NrniO0Gl{+}H0AMELSN5k{+}e56cV7lF4PtdK1B1t4dqoASVrKrD2L5Huc1OXNtsZ2t6tI0JJLAYySO9eAfE3WJbjzVXcAM8D0Ne8{+}L0{+}z2bKAMkdTXz74vxNISADuPHanRV5XFVlaJ5FcWgiLueWY8k9q8W{+}LOpfbPEUcIOVt4gMehPJ/pXt/iy6i06KU7gqIpLegr5n1a/fVdSuruQ/NK5bA/SvrsHF35j4rMJ6KCHafF5soUdx2r0hdWmufh9puXUxWk9xZqe{+}0{+}XIP5HH4159pv7kA/xsCMe2Oa9Zj8MafpnwCh1036PqF74ha2SwCkOkccGWkz3BMijFdFd/C/P/M5MNpzehx0N0JngjIw0hWJiO2QR{+}vH5V678GdXli1EQgqJowUCMMKXTlSf{+}A5/75rwqS6Iu42Und5qEH6E13vgXXDpHiVZ2IULJuJB5BDEfqCfzrOvHmg7G1Cdpq57zrtvFKmoWSjyotguYyTwVY/dH4ZP5{+}1eTfEDTl1XRIZ449s0JeMgDJyOR{+}gPNe3au8bwR3DbNyZQ7f4lPzofb5WQfjXBavpzvql3amHb9otmntwBkEqd2M/Qkfga82jO2vY9OrBNW7nnHg/RgvhKWT{+}JLkOwx0BBH{+}H51z2sWUi3M5I{+}ZGKFc9epB/wDQq9n/ALIgh8PXqQLtWKFeV{+}7nKYHPfANeYeLYQwtblTnzYip7/MjMp/kPzrqhU5pM5p0{+}WCRyaK15ok1ucma2PmRnvt7j8P6iqej30lnd{+}fAQGUCQqQCDgjgg9fmwfoTWmAbFjOjqMYDdwVOB{+}RHFZdswsdQnSMlQ4ePaeoVlPfv1rtj2OKelj7w/4JmeM/tFx4l8HztII3i{+}0wnPCkEcgHuPlOfavqnwdFD8Kvi1c6VHGLTR9dkaQQou2OO{+}xlgo7CRMMPeOTuTX52fsM{+}IpfDH7RfheSWTy470NbPtGN4PyAEfUA/QZr9OfjNojz{+}GrvUrcE3dmhuYAACTLD{+}9h/wDaie4kI71LRk9XY9Wj2hiR0PNLJluOPWsPwRr8fifwrpmqRsHF3bpKNvuAf55rdUFtzE9eAPaqWpg9HYYAP/10Hkn0pzLwKD0piGU054NOI96D0oAaBSdDTwOKaeKAGgUU6igC0OTjpTvu8CjoM0o65oAD05oFBOKPxAoAKCcUCgjOKACgjiilAzQAjsojDE4UEZJ6da{+}Lv2v/ANqTT/A1xqfh3R7xLrWLi3azcRkubMNguDj{+}MnjrxXeftj/tGT/Cjw/FoPh2ZY/E{+}pqds{+}NwtIuhkPv6e9fFvw{+}{+}A3iH4vXtzf6Vb3N2QWe81u6OZZSxOMMx2rxzkZPJ5OawqS6I6qUPtSPnt0eaY3M7FAxLBcY4Pv2Ge/Ws7UtZSKQlmBwMhQMED0A7D9TU3xLN/wCHr82TIIHEkiMxPO5ZXjbv/ejbk8{+}9cYXYylQdwJVS2c7j3/Dj9aIUb6yNZ1fsou3mpzalKWY4RsBUHQY5OKv2MStp9wc4bO8598kfyrO020a4u7ZSCFY4yOxBrZ09Af7QAGYxGjD2yGI/QitnZaGe{+}rJvDVkJte8tiVfyiVI7kf8A1hXoVlpf9n{+}JbGdBtImCsBwCDGR3/D865HQLZI9VtGYlHkdEU/UP{+}nT869RaKRorWdV2PbzkNnkkbcp/7KK82vNqVj0KEfdv2O9sfLF/oBKlZJAys/QlfMBAz9cn8qd8T5kXULZV5MlpMct6mQt/n61kW{+}obdU8PTMv3XZSp/hOEbArQ{+}J1uq{+}IY1k/1flSohHQ4VSM/X/CvLXxo9OXwnzp8U5FurS2iyG8uJGVlHrk/1xXk8i7Dg4yK{+}gr270ObQLrTdfgdLdzE1tfW{+}PNs5MFWbB4dGwAVJHTgr1rzTVvhhqSTqdOkttXsif3dxbzKvGe6sQR{+}o96{+}joVYpKL0Pm8RSlKTlHU4U11nw81Ke01OaKIOyMvmEKCQu3kk47YzVc{+}AtSgVpL8wWMScu8s6Ej8FJ96be61DpNrJp{+}jzOYpMedcn5TJ6ge39K2nJVFyR1MacXSl7Selj6S8IayLkx7369QOvFet6G6lEKFWB6HH5V8p/DvxY08cTFtrA7W56GvpPwbq6TW8a8bcZJr5rE0nCTR9dhaqnFM9a0SVkC5HDc5PFev8Ah62F3o5CgMoAxXhek3YVkwSNvr{+}levfD/WAwMLNyw4yOtePJNM9WMkT{+}M/hHofxM8ByaBq1okscmSp7qckg1{+}dvxb/Ze8Q/DLxFdN4feV7eMlhGj4bGe3rX6ix3BtmVCcjpn2rzv4r6JHeQyXYjBm8vCsPzx/n0rrw2LnQ0WxrSoUa8{+}WsvR9Ufn18IvDl74x8aNpFy0tvceUzRoWKMXUjIP4E/lX0D4S/Zo8VXviK7jeWZLONV8oq5B5Hf1q/oX9gzeKYNQuLIJq1hkw3SriRWPB{+}Ydcg9DxxXvHhT4p6ho{+}pT3FxHHqllNGhihi2xtEQuCQ2Duz74r0ZVY1Hrpc9PFYLH0E5YZ8y/E8u8MfsmeJL/Xr621K{+}uBagq0Lq3JBHOSevNdRrX7Hur2H2f{+}zr{+}4kEjKkimUDCdzwP517bo3xs02PUJmvLOeGIopiEX71z/eyABjt0zWpd/H3QDPbR29rqk0jTKjH7C6qqnq2T1Apv2aWsjw3iM3jJe6zwyX9j7Vbe1ldNXuVcglenHHTOPr2rh9T/ZW1{+}DwlJeXF8V1FY3ZTC7AKQTjI98CvsHUvjDotvakqZ7liCBHFFgn/vrAryvXPinqGq6bcWz6WllDIhUS/aCzcn02jHHfJ5qZezSupG1CWb4iSTTXrofBHxi{+}G2o{+}HNG0qC0jnn1rUbkRrFESzN8pJ47due1aHwd/ZXH9rWup{+}Jyt5eSuD9kzuWPn{+}Inqa9/8Q65Z2UhjiHnsqbTI/wAzD/gXXNafw3imvbtJGXO05HBHHasp4ucafLDQ{+}gngYUIe1xT5pLbsj3HwtpVpomlw2VnCsMEahQq8U3U4FRmccMeckVb05hHAqg4OMZNZniK7RUc7h8oxgda8WWruzwVL3uZnmXjq9JQpu3HvtrwXxfItu7ux2so/DNes{+}NtTRpHJIxz3r5z{+}KnimGwtLieRyAgLdf0ruw1NylZHFiKnLFtnlfjjxrZWWt2sFxH59s8o{+}0w92i6N{+}Pp9K8p8SaF/YmsXNuhL227fBKejxnlSPwIqnq2qSavqk11L1kbgDsOwrc0XxHI9mNOvbdL22Qjy/M{+}/GAegPp1/Pgivr405UUnHXufGTqRxDalp2ZHp9kTBayqpdmUgADJ3ZIx/L866/xRdvHoun6KGLQ2Cktzkea{+}C5H0AA99oNQ6fci2s/NtLaOFoywCld238yfT9KrXsRn0uaYksTlmOerdP61LvKSb2RpGKhFqPU5VTm9hZugdWIPpkV0WiXO6bdg72IXJ6ZHP8AICubgIN6jAYAYEk9Bium0yGSKytnZSGM/nYPoTj/ANlrWpa2pnBan0z4Slt9X0G0SVmb7RYqDzzujYofz2R1S1WwmjS0uVmPn2TRkODkt8i9vTgjPrJWN8Ob57W{+}0xTkqGZUB{+}7g85/8eH510PiC5EV6rQMxjI{+}zSh1wQF27WH4j9DXz{+}sJNHvq04pkp0tD4P1F2Znk2GdXUZDgvwP5/hXiursJvDgdlCtHdsNuOQrAEH8xXvqMYvB1w4ciOCF5HZD/AhH6nJ/lXg9xaGOLVrdt2I1V8b{+}mGAIPr3rbDyM8QrNI4zUQsVq/mE7SoTA643f4H9KwtTHkX8UuPvIPpkcf0rpb9CpRdqncv8YHQgVg6uit5Q6hd6r7detetTPKqxseofA5vsnxM0fUYwv8Ao93FNGCCUXFwgIP/AAEk/iK/ajWrWO70KXcqNiISNuHAwCf5D9a/Gj4CWBubq5m3IZVjkdQ2SVxtGQP94L{+}dftFdRBNMnhdS2{+}AsQeSfl2nP{+}e9F7tnPUVuVnD/AEGy{+}H9ppxOWtXkiTv8glcL{+}QH516cfugD8q4T4VWD2nhnTGlIaV43ZyDn5jLuP8A6HXeY5NOOxhLdkbDgDpSdsVIwpp4FUSMIwPekp/50088dKBjc80mKcR1FJ0H{+}NAhMGilHFFA7FsdKOlFLxnFAgABpQMUtFABSAYpaKADtzVLWrma10y4ktoxLchCIkboW7Z9vX2zV2oriMSKAyB1zyD6YI/rQB{+}YnxvnvoPixb6t4lnn1O4a5InLxgRIUYYSJSMEAYAHU4c9q{+}6fh15dtpL2{+}nol3YtEhi8pFQNGRuXoMEhGX5h1GOvbl/2hf2brD4m6PeS2MwtLyYfMxUEFgDtJ4yMdj1HPUEg{+}O/ALxlr/AMNtfPg/xjNFZ6lYruWSZsLc2QPEqPnJ8skkrzgE9hxzpckvU6m{+}eGnQ{+}Of2yPBMnh/x34gkbyxFa67OqqEKssVwq3MWe2Nzz9OhDeor54t8FiOd4dsDH0xX6Y/t9fC{+}DU7R/Eaxqi6lZ/Z5XAyr3MO54GJ5wTHJMqnuQgr8zoiyXfkyodySMpzwcHiumDvp2M5LRPudJaQNbgRAYaXYynHOCxHX3P8AKtjR7IE6mMjelqpYf8AYH{+}orKS4UXlkwcBdhjYPwMgKwH{+}fSug0NYm1fVowA8htXXae/HB/nXPO6Z1R1QuiOhsoJGcebDPF254J/{+}vXsE0ZuLJkHJZIpj3J2fKB/46eK8a0F2OhXkqruMNyh68gAt{+}nNew6H5k2noUO5ZLQtwMk5Abr6fOa4sStTtw0ty9fLG1nbuozPHdRunOCyspBPH0Wtb4mJJc6j4enCuIrpV7feJQg5/wC{+}VrHuSo062m2b4/LidnPfY{+}Dj/gJNdT4wVbjQ/D8zB91tsjHHdXCH88fzrzZe7JM9Fe9Fnz74mtDe6deQAAOjFQCcdG4/nXlU13cWbkWlzNAoz80TlR29Pxr23VbPfPrUIGJBM4AxyCQSP6flXkF7YF7xvJH{+}tO8jGQF9P519DQ1jqfP4hWldHN3tzPcsDNLJK3XMjk1Wq9qEIWUjBA6ZA4NUipzxg/Q13JaHmSu3dl/Q9Xk0a{+}SdDxn5l9RX0j8N/GUc6RYfdGwyGzn8PwxXy{+}UbP3T{+}VdL4J8Wv4dv0WRz9lY5yedh9RXFiqHtY3W56GDxPspKL2PvbQtRWVI8E5HXJr0Xw3q32S8hkVhtU4yfSvm/4feMftEcQ8wP6YPBGP1617Not{+}GAHGDz06V8lUp2dj7SnNSV0fQg1IXNpHIrZzz171U1qAalYOpO4sPfnviuN8Ka6GtzC7llAOOef8811FtcExHac5FcdrHSmfOHjTQZdA1a4baY9xPlTqOD/ALJrI0LxhNp160UxdYuP3TE5NfRvibwjb6/aMrxgs4IO8ZGa8C8X/DK{+}0uRvLSSSDOVYfeX2B7/jXbSqq3LI9vC5jOj7k9UereGfF2mXl1b/AGuWWKH7zGEBiOnFdTZa/oFtI8l9Bd3shKGNIpVhTH8SsSpP0Ir5j086lpFwCzudmOZBj/61d9ZfELTprX7PfWKrJwfOinY4OecA9umfxrf3U9LHryxeHqK/M9ezsep{+}JfFGjzLDcaTaSWEQHzQy3Al5x/CcA{+}vBrzrxN43nMUsMMeWAI2dc{+}9M8X/E228RtBbafpUNja2/yoYYx5rjaB8zAY7Z{+}p61QstIutUUM8PkI/wB4kct{+}NQ3FSuyY5lRpUuVLX1v{+}Jg6bb3GtyRgpvc8AA4B{+}vtX0B4A8NjS7BC4O8gZbpmsbwR4JitVR1gIAOenJ59cV6fDAttBs29O1cdSpznzuJxE8Q7y2FM3lJnJyfeuJ8Xa2ILdhuxkZwa3dU1NYY3O4AAV5D4514bpPm4HXmsUrs4WcF478RGFJedo67vSvjv4x{+}LW1q8a0ifMUbZkI6Mwr1v4x{+}OjbQy28EhMzjaOenvXztfWzzwyuQWYDJz3r6fAUuX32fOZhUbi4ROfPIFa{+}kp5lyvBO7jisoIOR6VsaQjLNFtHXow4welfQStY{+}XhdM9GSFYNFluFVlEiA{+}vzjhv1wfoaztUAg0LpsM4646Drx{+}NVrK/ku5hafNCu4Boye/Hf0xzXSeNNOgEFvbRnfthDAqMduSPyUVwSfLJJnqxXNFtHmdlEXuEOOCwB9etdrLAtppo3EPswo6AksAf03E1l2GkmO7gwuCr/d7DkDJ/Oug8RBFKWy5CGRkIJ5zkAn/AMd/KlUlzSSQQjaLZ6H4euDajTZQfnYF4l68hFxn67K29Qu/7R8Sala5aOSMkhOckldyNj2JYVz{+}iEw6L4cu85Kz{+}Xu56qRn8Oa0b9XtfFdlesdqXcCE4PBwDz/SvJqK0mevT2SOi8Q6r9h8CXexyq7TEyN2LMPz5rxy9uXTWpVb55LmHkdMEruP/j2f{+}{+}a9A8Zv5XhmaFX3BzHK3qMuT{+}OeK4oaVILu2vGG5kZIm3H1LAn{+}f51phl7rJxL99I5bWQpktiQDhCCR2{+}Y8fkcfhXLXsnzhOQcNyf8AersNXjWS3YYEUqhz7dTxmuIvVMcojIIkXGR9ea9Kg7nl11Y{+}o/2afCc974c1y7WIb3gt4Y5D94hrlVOPxAFfrHqNwyl0DK6qoZySQdozx7c4/X0r84f2adKvrTwj4Ijht/LTW9UELk/MXhR0nwOccbTz/tiv0avo47Tw7eXkzh5hEzyFT0IXt/n8qIu/Mc9VW5bkPga3ktdB0yJgocxySNg9S0mev6fhXV4x161kaDCUsLZcf6uBFHGCDjmtcfNg4xWyOViYx{+}NJgU9qYOgpiGsOfakHP8qccAgfpSAUANoPSlbnFNIoAAM0UpHviigos8Z6U4UmOaXvQSHenACkPBpRQAnHekxg0/pSY5oATHrS4FLQRigCKa3VsYAHI4xxXk3x0{+}Bdp8U/D2LWU6br9oTcaffIC3kTev8Aun7rDGCCcjpj189RQyBuv51MkmrMpNxd0fCVjrOoeOvhp4l{+}FPjWJbLxTpik2fmKNhdBui2A/wAJ28H0bpxX50fEXRY9M1Z7oRGFmkaOSMAqUcDkEHv1/I1{+}y37QnwUXxtax6/pLtZa/YrxcQj5nQcj6kHt3yR3r8zf2gvBEl3q2plIEXVLg77i3jU/LOucvGD1DgEjknn1BxyqTpVFc7larB2PBdQPmSQyw5Ec6LJg8bX5GR/Kuw8IyLc{+}NIJpMBbi238jCsAQW/wDQTXnolxp8ayHm3lMTKp6q2T09iD{+}JFdj4dudt/pjkqjRyPA2D2YHp7Hd{+}hrqqRujKnLU0tHKpZ{+}ILWNhvhjRsHrgMoz{+}WRXqfw31BZrXTQ/3EXYy56gO6Eflg149MTp/j26ib5FvYZImB98gfrXSeANfW1iMTYXZcEOWPRCuTz{+}Z/CuStG8TroytKx65Dm2Z7ST51tJ3tsD{+}JGAKfnlv1rr9QtVu/CFwnyvNbEyLjncGUNx/30fyrgIL9by5kJKySvbpIx3D/AFkTFS3/AHy5JrvtEv1dEjZg6TxbQzDksoKj81YflXj14tWaPWpS3ieI{+}MsQ6/PLwFujG69wWYEDj65ryi{+}TbflTyHbZgg4UZPX26/lXtPxK0SS1jgYknyHG4jJIG/K/h98V5u2lI2pyzSDbuR5PXkZ4/Hj869jDT9y55OJi{+}fQ8y1bc95J1IDEAY9KzmXsRXQXVpIruxU7UGckZz3/UmsaS3dSoYYPfPFepF6HkSWpVIwaSpZUCE4O4{+}3Sou9aGDO8{+}GvjqXQL2O3uHJt8/I2fun/CvrHwV4xiv7SE7wHwAcHmvhaA4lHavSvAHjyfw/corsWhOBkjO3/61ePjMMp{+}9E{+}hwGKaShI{+}5tI1l7e4WRXXB7A16d4b1xbpVJYH1Ar5j8L{+}M4dTiXZIDuUc5zgV6j4T8RhJ1yxGe/bNfM1KbR9RCS6H0RpqJdoqn5j2qS88IQ6ggLRhwepIx1rD8HaysyruIPIHPevTdPVXQYKtn0rnUbmspcqueV3fwdsL1yWgBz09DVe3{+}BFkrCRYTnHQr1r3K2sUJyE59QKspbpHhcZHUdq6FT8zleIa2PEovhFZWjllhUkegHFadp4BgiH{+}ryewHcV6zPaoWJCjB64FQGyQk4x14qJU/MuNds4q10VLNVEY574rP1qVbVG6KwHeuuv0SFWOcGvN/F2pqiPk4GetYNWdjeMnLVnF{+}JtbEKMpYnP4fhivnb4o{+}No9Ptp3Zzxk8HqecCuy{+}JvjVNNinczBVQZJJ6e2K{+}RfGniu58T6k5LN5IY7Iwf1Pua78NQc3foc1eqoaGDrF7ceItTeaXJZm/IelXovDrNbEsp7ZNb/hDwk94wkYct7V3V74X{+}yae5KYOK9SddQajHoecqDneTPl/VbFtP1GaBgRsbjPcdq1NMjJtm8tvnSMyDHop5/Tmu1{+}IvhTzdPOpRJzCF347qTg/rXFaGVke3R2GNzqfQ5HSvZhVVWnzI{+}eqUXSqOLPQvh1p0Os65Z3RTzJFJVo8ZDsflH{+}P4V6FqHhNtRv55RGZY1PlRAHAJHC4/IfkfSuC{+}Dm2311Le4DiJGeaTbwQgHP4nPH0r27QtdhuiwkjEczs8qqOiqq7UUfgCPwrhqO03c9GhFOmjy{+}z8IS210XdQVEpVmPQN6D865XUrY3{+}pHazODkBvYnJP8AL8698{+}I9tZ2Wm2tjZJuwgBZjy7k/O7Hvzn8/auS0fwXLDELiWDiMFjkAr5oGFU{+}y4DH1O0dqxjUv7xrKH2UUnBsvCNkrYCxyTSqM9FYquR6nv{+}FT67ctKLGdSrCCNPL9MbmH8wKPE8UMOn2diG3YRMAEZ5wR{+}ZJzWdqlzHPBYxIGQSxmI{+}gIOQf/AB4VhLWzOiGjsaHiq4Mum2oi2yyvNFEiNzuJZgOPQGqGr6FNppuxh1jikR/l5B27MnjtuyM1r6M9tqPiHRVuIT5cMvmsZGzkpuKD8wPWug8R6fJceHZb9Q7NeSvk56J8xH6oT7ZFVR0gFfWVzxHxFatb2Rcjd5obHfA3YwB25zXn92rXGriHl3Zwij3PAFexeMNMW2s4GyMqNuTxwPm3fm2K4L4caGfEHxU0CxeNSJtSiMit0K7wSPpgEV6NBrlZ5mITTR{+}jXg3wkvhLwv8AD1rYyxXemWVzrsnlAkfZkZYgAuTjehJ/AelfYeqyDUNFsoA5ZLwRcKeqnG79M1wXw{+}8NWc{+}u{+}IPMhVrOx0u00GKNM9AjPLj2zKv4g56V0ngGHfptjYmQu{+}jb7OXf13RZiH4lcN9CPWqgrfM4qjv8juLSDbAhI2sQCTj2qbBAFPC7Vorc5hhGaQrxSn71FADMfnSYxS9yDRt4zmgBp9aaVBPSn4zRjnPpQAw49D{+}dFSYz2FFAE2eaUDIpeKMYoAWjHeilyPwoASjFLxRkAUAJSgY680lOoAQEE0tFFACMoZSCAQeCD3r5a/ar/Z7g8T{+}G7/V9MtlNzaRtceUF{+}/jkgEcgkZ9ecHFfU1VdVtEvtPuYHGRJEy8jPUVnOCmrMuE3B3R{+}CHxA8Py2GqXEpt3NveI6K7R7HV1OQj44LZGM9fwxnmNNlcWyyo3K7CV9Sv8Aj0{+}tfUn7QHh2Ly7q8SGOHZcGK68sNsjkB{+}WYcdDtwR2wfx{+}X7q1fSdXvLWRVQN86Ec4HXj6dfw96mjPmjbsddWNpcy6mx4znafUdJ1bzMrcRB{+}P4SMZX65FWNJuFttYlAOY7iMSD25yf54qtJC134XnjzuazkErJ3weGK/jg/Q1lafeGG7hUtkpznpkHhv6UNXjbsCdpc3c9b0bVxGbadsD5lVyOOSDEw{+}hJzXqPhLVUk06B2/5YzLOFHVUZiHP/AAFSfy9q{+}f7bUPJS8jYDAIYfiM/zGfxr0Hwd4mWyWOeQlom3RyRg5JG3PH5P/wB9V51WHNE9KlOzPTviT4cFz9oeJnBlUCRD0HQhvfnI/EeteD6vp4toNQXBDqFQA/5/Cvoy11p9b0bT2klSVdpt5JT0kZT/ACbCsP8AeFeVfE7R309HZFyZnKhs8Buw/HGPyrDDz5XyM0rx5o8x4XcxE2M0z4ALAsvr1Jx7dK5p1DASEcbjyOK73xRbrFLYaYnySCPzJFb1Izj9B{+}Vc9qdnFa28cQG1sknAzxn6/SvfhLQ8Oce5gzacGdQpPIJJx0AFZjDknINdleQ2tppMjbj5shCrheduM/5{+}lckQjO{+}FPfHP/wBatou5z1I2ZFH98VsWLbX5rJGwDPzZz7Vp2uCAc/N7Upq5pR0eh3XhPxVc{+}H51CszW/wDEgPI{+}le7eEfH6XpieOQHpgk1802MhYAZI46iug0jU7jSblZoGGCfmTs1ePXoKWqPfoVnHQ{+}//AAB4vSZYjv5OM4Ne{+}eGfEMbIoMnP1r86vAHxU8i5VWYoxPzRuefw9RX0v4O{+}JsU8EcizdcZUnpXz9WjKDue9SqKcbM{+}urHU0cJkgMB09aui9jIz{+}ArxPQviDC0a5mGcZznP511Vv4thuItwdcH3rONRpEyoXZ6Abpc8Aj2qvcXaouM4xXGjxXAFz5oB9CaxtW8f2ltEd1wCR79KHO440bG34h1kQq5yPYd68D{+}KHjWHTbaZnkAIG7k8Cn{+}PfjBZ2VvLNLcpEi5G5mxn{+}pNfIHxR{+}Kk/i64khti0dluIyTzJ9fQe1VRpSqu5dWpGlHzML4leOpfE9{+}6o7NbqxIwfvn1P9Kx/CPhObWr2M7SRnJOKraRo8ur3qxxoSTjmvpX4X/Dg21vE/lZfjnFepWqRoQtE8{+}lCVaXNIr{+}FvA4hgiVYiGxzxV7xPoDRWLgLjA5BFe0af4U{+}z2{+}9UwQO39K5nxfoq/ZGG0nIIrxfatyPXcEo2R8/HwsNX8O3luVGXjZcEfl/OvnvxDoqaLPZ3dumYJMMy9lfoR{+}eK{+}ydC01IrC93rhUJ/LFeK6NZw3/irUrBLOO{+}kT7TsgkA2iNjndkkAEZ45HOAOtexhK7jJ9jxsbQU4rucd4NnS2sry8nBClVjJHVh1wD6ksv613PgyM{+}ItXLNmPeQo2dEwdxPsMDr6cd68w1zWTFcNCCqjzWeTACqrE84A4xyAMdhXWeFfEI0223szRsUDN1yoPReO5/oK7asW05dzz6Ukmodj3eeysyqRLMtxdw8JKyggkcKx{+}nXHcke9O1aGOC2azgCPBAg3gt944z{+}POT{+}Ga8us/G76dGLl18{+}VhsjiA5Y8ADFWL3xiNP8AtVzeOZdsRlmdD8u4rwq9uuB{+}NciUnY6pOK1M3VT/AGj5k7ghEXzPlGM4l3Lj2wlV/E8KW1nahCMiXbnPzdAfw6VEt016IQwCzF1Z4gMbQQ{+}VHtzj8aPHtxttbMRklwVL8c5IAP8AI1q1siIvqV2vpFvrDynVQCpBXg8sxP8AP9BXsV75DeDNHUSItxPbH5DkkFl4A/E/pXg2r6nFZRac7MAAoUtjBwCefyIr13wZq1r4m1bRrKJg5lJd3P8AAuTkY6fdTP8A{+}utYK1PmZE5XnY5z4waIbdobSND9oeMBkA5Ygbf5lj{+}VYv7MelRQftAaNLdQqyQvLMExuBKxNt/Ns/0r0nx/MNX1HU9TiDTL5Lm33gDGOEbjoCScfQVh/s3w2ug/GfQ766cPBBfWtsxxgszDj8M5rajK1M5cQuaVz9VfhvoD6D4SZJwGv7m4lnmPJ3SMxJx7U7w3YLoPjTU7YbimqxfbSzEkNMpCyH2yDFx7V02nWYgijyAscCBVI6ZwAf5frWN4hiNnHb6yMq2nXPnSAd4GG2QH2Cnd9UFdtrJWPI5uZvzOmXOMenFKwwKXOGPOQeeKVulWZkZBApKfQelAEZIHakp3eigBnagn2NPoI4oAjPPrRTgCaKAJ8UZ5xS0CgA70vFHPrml6mgAxmm45px9qQDmgAApR0oo6UAKBmg8UL9Kd25oAavNJJwjH2pwGD7UMMqw9RigD4A/aB{+}HH2rXfEVlbtFE0im6RmGCFLnjk85ZnHTovvivzz8YaXc2slxFNC8N5p820qB0jJJXHt1/MCv1g/ac0K40u9i8QwIrk2nlMHB2owkVovxLlh/wMV8LftNeEo7DxLc6jZoC0qs7ryBLDjPfuCP0rhX7urbuekv3lP0PCPCV3HdLJG74LKI5M9wTt/kf0rHvrd9MvZA{+}SYZPmHqp4JFOH/EtvEu4CXgk4bI4Knp/n2re1q0XVNNhu0H73JXfnOR2B/Hj9K7L2lfozBaq3YmSFnnt3TEomj2deG6EHNaHh2{+}kt5Z9Ob/WA{+}ZASe6np/M/gKzPClwt3o7WrKRdWr7kXvgHkfhmrmrRtaTJewDLI{+}8Y9PT{+}Vcb0fIzsTdlNHtPw11iLUtMv9Ec4uY1FzalmwGAGMfoqn/ZU102rWZ8T6CY2QG6RQHHRm9D6g/Lj2Irw/TNcOmahZ6pafJJA4lwvAKHhhjvwa9qS7ia7hvopB9kmUP8h6Z5H4f/XrzqkWpXR3wkpRsfO/iSwli8TXEs3LygjaRgZyent/IjFYOsrGLSzZyyyPu4xzweeOte4/FvwQ1wP7UtmLx45jCgnJH88/5614lJbTIxe6VlZU2hGAwu48f/rr16NRTSZ5dWHJKxl6lZ7bQYLNgKTkdsVzsNszPcsOiISSfriut1Zh9lKxqRlF6/rWDYoyaVqUrHG4rED6/Nk/yrqjLQ4qi95GK4JKgdav2J5Ge1WNA0dtZ1G1tEwr3EgjEjHhMnr{+}FOFr9nuWVTuUMVBIxnB4NaSfQKSe5q2sDKoIAINakKkIFP6UmkwieMDrxVya0aLDc4HrXnyld2PVhF2uJAxDL2PUY6iuz8OePdU0Yqok89ByAx5H41xMYKj5iMdsVZhfyWBJODxj1rnlFS3OqM3F3R7xoXxue2wZfORuhK4YV18Hx/jMajzpRnsBj{+}tfNVvdgA4HStK3vCq57iuKWHg9bHXDEz2bPf7348SuhET3Lj04Fclrnxl1W6R1gXyif43Ysce1ecm6LJhW/E1DLJhcnvURoxXQ0eIm9mO1fV73WJ2lu7p5mPdyT{+}VZdtYm5nVdvB71aYksq9eOmK6zwX4Zk1W9QOmQWBIrZtU0Zq82dt8IPh6bydZvJ{+}XPcV9YeFvCkVjbR4QAgcmua{+}GPhJLOyjVEz8vU8e9ez6ZpqpbggEnHp/KvCrVOdnqQioI5{+}609YogFGDnJ215/4tsfMbYFBAz098V7BqlsqQj5QMdzXn2t2jXEzbVPHQVybM7ItSR5CukMLO/jHBbOABntxXyOmp3vhfxBqtwr/vZo5Y3D8gjdnlc4OCBjPQgHqBj7wk0pldgVwTkjivkH48fDxtAnuNSjuEMdwzg2/R0Oclh6jBOfT3r1cFUXPyvqeZjabcOZdDwKS9ee9d5fn2sWwe7VvWWoPFC7M{+}f42A7n/ADiuWc7bkq3XOD9fXP61dvJ5fssSQjBbCt6j/P9K{+}slG9kfJxnZtm9a6tPcXQkJy{+}SoyflQeg9z39q19d18zNZafDKjKnzTuSfmcduPwH5VxyXn2S1hVfv4OCOvXk1LoarLJOrNkycl{+}{+}7HQe5/wrL2a3L9o37p6T4auGubyWTI3ME5YZCDIJ/ln8Kv{+}MSsuhpKCwWMxr6tjJBJPr0/Ouc0WdbexlfO2SdSijnPHv8AQGuk8RsLvwnPtYM0Z7nJPfHtiuCS9656cH7tjgvF2qxxyRRSDMaohCKcDPIP5gGu{+}{+}GF5cLbxcqHkjYKka/MF4BI59f5V5J4lCG4RzJ/CrEHk857/j{+}ldv8AC/WUgFxJIuXEIh467Sx4{+}gAP510zVqByRlzV9T2rWdYtU057WSARwySxxM8ZDEogcgEZ4z1P1{+}lRfChI9Rj17WDASdI1Oy1t0SPjy1EgK9{+}AdhyeBn1ryq01i51fXls4ywJmAUI3I6gn3{+}6v4Z9a9t/Z30dNcg{+}LWgQyws13oE6R5bl3jCN8vXoFck47cUU4WikxVpps/V6zmF1ZQtGwZJF3Djse9OlgS4NxDIgeF0Csp6EHII/KuW{+}DOs/8JN8KfCeqO6SyXWmW7O6dC4QK{+}P8AgQNdcExLKcfwqOfqa7VqjxtmYvha6kl0hIZ233VlI9nKx6sUJUMf95drf8CrarGtrc2PiTUAABBeRpcD/rogCN{+}a7PyNbZHFCB7kZ4PHSkPQ0pPNIeRTEIBRilOR0o3ZXrQA2ilpKAG9CaKU80UAWMDFAHNKOlFABQaKDQAUUuKMUAGM0KMmgZpTmgBcYoo70oHNACHg04AUmMnmnUAeNftD{+}G/7Z8E3CpGGaNSqgjjIKMhPsGUGvi39ofw4mv8AhXRdQH79HjaBpIkA2DYrKTjoMuw5OOmK/RrxZo41jSLi1OcSo6k9cZBwfwP8q{+}P/AIhaJGngYoYYrV3SWCSROiyQNswV903Y9SorhxEXujvw0raM/MHZFbTXenXTKkIbajHqp/z/ADNJbajN4fvJrK9/49JhskX{+}6SMB1P4fpWx8T9HFhq9ywAI80rlTnJ5z{+}v8AMVlLCms6RGJQZbi2jKEZwTGD/wCy8V1wanFMmScZNIiilk0DV1kyHGcB1PyyDp19x39fpXaOEubdDwyPgI69DnlSfTIyPYgCvNYJ2ERtrjdJGPusOSvuP8K6jwlq6lBp1ywZXB8l88HPO38{+}R71nVg7XKozXws1bJvsTyR8si8bc4ypOAf1wfr7V3fg3V5I7F7HeTNCDJbAn/WJ3j{+}vp{+}PrXIyr5dyXlVmdQA4/vqRgj8f6U2OU2kuFkcSRYZWBxuHUMMd8fyrilHmR2xlyux7d4d8TWuq2ktpIFlfy/MQ5wSpA/MdAfTg9mrx7x1oHk3LXduW{+}zmXayMDuixxg8dvpWzJqfnG11WybyZgcyiNej55YAeo6jvW1qYXxFYPd2aItyYws9srfLICOq{+}nsfw6isabdKVzWcVUjY8WYxz2KTGZG{+}dk2jjAznjPSs2OKOaxhtIWO6admI4PHTt{+}fToa6m98PvEZRHC7xFid4HzA{+}hHHI/WuevY7nS5nmg/wBH2qY/LU/PtZSp469OM160JKWx5dSLSuzoLC4vPhxPYa1piwXXLQtOo3pE/B2gAgg7Wzk4zzjOOOZ{+}wOJhkHk967D4M6JJrB8QzXdstzZSwrbO0{+}f3JJMgkyCDwIiDgMcPkKcYplzZiSYkIq4bO0c8dqmU{+}Vl0oXjcpaBBibZwB3IFdc2iLLEDtyf/AK1YFrD9luY3xxnBGK9K0OJb2LGBjqD7151Wet0enQhdWPOrrSZrUklQMdsZBrLcMG2kYJ7CvYdQ8OmWIkoNuCK4nVPDrQuzKvTnkfyqYVk9zaVJo563Vu44rRt1{+}bcenoKatjIh27cgdCRxWla6Wznpz1yOKqUkRGLuQ/aN2GGRjikUvKduCC1bUWgs5UAE98AVu6X4MlndB5fB74rJzSRuqcmzG0XQJb{+}4iCqRnnOK{+}hvhj4GW28p5Ihnrms7wJ4C{+}zOrNHjAH3u9e{+}eE/DyWyINm0D0FeXXr30R6VKly6s6jw1pQt4o1CgLwea7q3ASBeOnQmsbTIdqbcbSOnFbDMAgUcjJ6cV59zeWpl6qPNDLnOTXLXVnukY4BJPpXUXqliAefXPeqUtnuUkjGOc4rN7mkdDjrqz2SM2D069cV85/tJaB9p0F7mOEubYvNIsfJKAc4454P5Zr6ruLQFwpHBzXkPxp8OR3mnTJiXHkStmBXLnCE/wc{+}vBIB6ZGa2oT5KkWwrLmptH5o6pZSWdy{+}QGVWyG7Yz1q2I0uVjmiYHcNpTOPwz25rS8QWJtJPLaTzWCBtvPykjJB47HINc5b35tHLRRYBPKg5H/wBavuYSc4po{+}InFQlZ7F26jDTKGcCNAQjOuT75//XWpaRQiPGxkEa5Z93ByPx9c1mLr9tuDTQMrZ646Vt6S9pdgG3cDncy52k/h3{+}maJtpbExim7pmtpE5urxNhKxoPusPmJ9sZ5J6/WtLUtYeHRktDHhpJG8xienHP86i0y2VnEcexSCGYgbSR37de2BUvikwXkBkij8uJHxk8Hpz{+}tcTs5WO6Dajqed{+}IQVmYkEs0Q25OSNp/{+}vW54RuzpmiXszYIclRkcggdM/U1ma{+}VubC3mTlhlW{+}h4/mBTYX{+}zeHY03FVdznHQZ//AFV125qaj5nHflquXkdb8Ktajh8RXOp3JAFvGSinnnBwP5V7z{+}yn4li0r4l6REqKJL{+}0uYpWY7VkDDbtbPXIGO3X618p6VNJZxXlv0bYxfI9FbB/PFelfDnxL/wj{+}r{+}H9WQ7WsLkJI2eAj4IP4EEVclbYhSurM/W79knVhJ8MbzQmY{+}f4c1i90xlY5IQSmSPP/AJAPwr2fAZ5mB54X8h/wDXr41/Z/8AGM3g/wCOHi2yLifTvFGn2{+}uQbT8u9f3b47ZIZfyr6x0rxFb3VmcOWfnIPXP0qo6o5JwfMyTXJDb3ulTgjH2nyWz/AHXRgP8Ax4LWtuyK5zxPexyWET527Lm2YE8/8tlrbju0ZRk4OKfUhrQmPNFJvX1pcg0yQpMD0paKAEwKQjBp3ekOO9ADfxNFLRQBMOAKWk20uKADBNLntSfhzS4IoAUUtAooAKKKOpoABT6aOpp1ABTCxPTp60{+}kxQAghDLsJ45zXzp8QvBk80Xi3TbdHa4huhqdtuIwDIADz6ZWVfbcPWvo9DzXnnxEs20vxZpOsrkWtzDJpt2du5QGIaJj/uyKAP8AfrKoro1pyaZ{+}Q/7Svg9tA8W30K8QTAXEQVcBlKgjGe4Ax{+}FeH2F3Lpd4ksBYSxt5iEjr2IPsRxivtz9tvwVZaJ4oEEEE8pWT5HySuHLHacddpLAd8Yr4q1SNZbjYi{+}VJjpnpx3/IVnQ0XKdU/eakTara295uu7MCHeSUj67G7rx{+}ntj8ci9nWxkhkjXdb3Ch2jB5jfOGx6cg1csL4wbWZdySDbIh6HB4I961L/Q7fWNOMttJGL5HANschnTH3hxg4wB17jr26k{+}jMpRfxRNnQtWTX7aMSSh7pVChuhcDgEj16fX8qluyXBAIV05Qjt7fQ15qstxot3G8ZMbqMFT35PBFdpBrsPiKNQWFtfYyrfwy/X0Of1/OueVJxd1sb06ymrS3NPS9Z{+}wXbnloScSQnnYT/EB6GuttbzymW6tH{+}XqQp4XPf3U{+}v9RXll/eOsg81RFcJlQ{+}OGx2Iq1pHiaWwnRt{+}wZywJ/M/T1rGdLmV0dEKvK7M9XldL5ZJ4UzME/exHuPw{+}8P9ociuW1HSI7k/J/owTJ8oMWYnBwc85Htn86u2Wqq6LPbzLGx{+}b5DwP8A61bca2utYkYpDeEfeH3X/wA{+}2a4lJ02djgqhj6HffZnhspbOKeOO1RYlC7VneN95E2BnBDSgkHcQV5443dc0sahq0tygtQJJGKm0i8qLGeNq/wAK4IwPpUkLnT3RZosNghZOMNkY4I6{+}n9a2by7W4KPbRvcSeWu9pWAO7ADYxwBxx3wBUSrNoqFGzOGn0hlY5yWz{+}tdV4TUtdRoRtKqSc8DNOl0u4SD7RPHsLDIXPGO1aWgad9nX7S/zhiBjPU{+}1c86l0dNOHLI7Ky00SxbNm4{+}4xg1mat4ZSdTlACBx9a6HTf8AXRI4BYjGAeV{+}o7V1o0hbi3Bdd2eTjt71x89md1kzw6fwgyycJx64q9YeEhwNvI6{+}1es3HhoblULkHrkdantNBVc/uyQeBgYqvbEqkjjNG8IRyEELyOORXc6L4SRCpEY{+}XjI9a2tM0kRhTt244JrrNNsFX5cgtxnIrCVVm8YJEmhaEkW07MkcHFd7pdmI4/u4I/L{+}VZGmWojcHcPpXRWxYBSBgdCf/rVySd2dCjZGjbOseMY554FWXm4AIGOwz1qquQoBBxnrTJJwHUdugFIixM3zy/e3fTvU5g/dbsDnoAKrQR7nzzz1ArTRCIwD1xjFVEiTsYd1CBhug6ZrwT9pLUfs3hzUIIEE80kQjMaAhhESDKVODk7QR8oyNw56ivovVIkt7SWVwFRULFjxgDk14B49nvBJfnylaxtbIz30ruquvmg4iATIO5Gwck4wM9SKdKL9onYVSa9mfDPinwFqdx/aV/DbyQW2nObZ0YHfu6tngdD9K80ubW2a5kZZGgYkgpJ0P/1q{+}6PGWhHStSubm{+}trbw3p{+}riFmkkMk1pAqx9Aq7mdnx3AbJzk8187fEX4N3GmFdatrNp7W5AlhC/dKlQcn86{+}spV0rXeh87VoOXTVHic1s0YAK7kOcEdKLdZbMq5Yknn5O1dNDp5t4LgXEL7TCzRn0YAj07YpllBE13CmxSqgHnv0P{+}Feh7TSxwOm0zqPC9q9/BtEpSSXBkJJ4HHUf56movHlrd6Sn2doWht2XcOPzz{+}X616jo3wd1a90a38Q6EjXtuFKyiM/KpH16dB16VzfxZh1vVrCxtdU0ldKKBk80QtE746A5{+}VsdMj{+}lc8YpvmNJyaXKzxbSr5Zd9vJgJKSAx7dMVsLYBmhgIOyM{+}oyT2xWfa6H5d6iEqGGSUJ44xnn{+}VbSqthOJ52ZWKhtipk4I6//AF66mktjli21qZMDROLl/KLBR1P1yF/HH61p6cJtO0{+}SCQ8TAbXByBIpDAH8f5mpQwjtGeC3eC2JMhZiQ7dhz{+}tbHgTTYfEel3ujsxjlQia2nznYwyBn1XJ59s1Ll7txpe9ZH2F8F/GtvfeC/hp47mYodEvh4e1U9hBMPLRm/wBlSUPPvX37baKix5T0GPWvys/ZU1mJfEuufDbxCxh0rxXbNYNEWKtFdqSUIPQHlwPXI9K/Rz9mjxrc{+}LPh1Hp2rXKy{+}JvDkp0fVfVpYxhZcekibXz7n0pRCZ1mv2zeRbwNyZJ42wP9lgQfzxWiFlHGST1qzPYtPcq8gBCnILdQP/14/Kri25d/T61aWpk7aFBJZonHJPtVmK9ZTlj{+}Bqc2pCg8mhrJWA4GcUyeVCrqSnggg1KLtSm4EfSq5scDpj1xUbW7IuyPIGejDiglx7FxLnOM4zVjcGXNZKLKD2wO1SpdGEjIGB1wKCLGhj2J{+}lFZgv1X7zc{+}5ooCzNkcUZoAoyPxoJFB5p1NHWnUAFLjPNNJxTloACueaXGKWigBAMUtFFABRRRQAqnBqrrukQa7pFxZzKGDr8vbDDkHPbkVZp{+}7AoGj5y/aK{+}DcPxO8B3ojjC61p9sSJNpV96glNp7fMATjqARX5IfEHw4z{+}RqEMUkFy7SLcRFQqxyhjvQcnOOK/fO806G6WcMABJGUbjqDX5S/tmfBS68AfETWrm3hH2a/kOoWxjHyybsBwFzwwYkkeh{+}lY25GmjqhLmTTPjXTYvtCzwlcSoC6KxwDjqv17j6GrcZJADExshOxs4Kn/wDXxS3U6yXEd5G/lyjKtFx{+}v19aW{+}gW9VrmEMsjclQvHv79f51ruWmWJ5LbXiI76Jo7ngecv8XHHTr9awr3RrjSWJB8yAnhgMEe47ZqZ51ugrcxv2LcDPtVyzvriWMxH96ucFB6e47003FeQnFPUyTqkksflXREsfTe3X61E0nksNrdOQrVs3tjC3zLF5LdMYK59fasdrXHCZ442sOD7UJxewveTsy1Y6u1q{+}YgyZI3ID8prstK8TuGDxzFcjDcZOfcd/rXCxRQGRcxqG/uu20H2z0rTbWI7SI20FiFlPd{+}RnHauerSUtkdNKu4bnp1n43shBtluFkyQGVCF59SCcV2mlRQ6np9rcWl/bM0zMoiV8uoHcjoAc8c54NfMMkz3N5hE813wgwOWJ9MV7z4P8Ea/wCFPFY0e8mjkis3CO8UZKqcsCu713Iy88kggdK4a2GhTjvqdtDFTnK1tDu5tEvAFjklyNvU8jrW5oHhy7vwA0wZFHy{+}Uo/yK62x8NtLZC4MexSPuv19jWp4Z0hrW5KKBktx6V4bkrWPdtYxLGwbT3VVhKrkAny{+}frnvXoWgWYljKhcZHTbgZqe70Xem4x89xjHNaPhmzZZgG9SMHiuaUuxvFXHT6OyrlgMY6YquLAr80Yzu6gGvSRoIliDMmEwOlZOoaGsSkIBx{+}JqbjS1OZtLYIVzk{+}vtW5aKgO0DAbp61WNm8bYI6c59KmswxfDHPpUN3NIo6KzXAVg3PtWrFIEQE4PrzWRakYDbeo{+}lakBViAV28Z4rOxqXo7h3Y7c/405iZGUkc9OT0qa2tvMUFRkAVYS0yACe9Ul0MXoOs4c4Ax6nNasDDoR2/KoLW1CMAAx78VpxWzMSFTBPc8mtlE5JyRy/jmR20hLSEhZ7txGCWx8v3m7HsD2/LrXzb43s7IanbQXV3cLby3E0TzTZaF5cRPG7kbss4WROCpA5Oe/uvxUvJ9TjOlaZIpmgVp7qQKGO1cnajg/KcqAc44YeleG6rYqPDV9eW8smqRy2a/wBoJcDb5flzBXjQkdkV{+}f8Aa962/htJ/MmKdRG3Nq1/rs8M9xq9pBqunXhsV0DRis/nxzArGzy4ZU42/MrZGWxzjHo{+}rfCa0l8LWenzwRSGOEIy4AUNjnFYHgLwFYaHB4fSKxSEu0sz/ZbtYIyPtKsoYY3SBc8D6V7hqcIWPIUg9ga1qzuouOiMorlk0z4O{+}LHwDt4r4S2tv5EEKbTGkXyP1ySeoPXmvHvA3w1eP4iWlhJb7kcnyy3Ibjp7V{+}gHjjQVvi7hFckEAnqK4XwV8Ho7bXkuLmBBLu3oyD7jfjXqUZSmkzkrKMGz3P4RfBfQPDfg{+}0SKwjEnlp5gZeQ4wGIHbpXmn7THwl0Txn4T1fUZIcRaejLb3EK5kZlBLNkn5gTtXPP3eK{+}ltDid9Jjj{+}ZCECyBTx06564rivj3bwxfDvUYtgQSRiBEXGSzEKgA9ctmvW5bRujw1Juep{+}RHiTwxovh/VUkk1BHWMCSWIoUkB7qFPfpyOOtcRPFc6/wCJnuNhET/NGg5AXtz65r7Q{+}P8A8NdJlWBmtIvtKJ5fmAcjHv3NeEaT4Li0{+}4RLeMkn5ck8n/OaFNJeZbg5PyMvSNKsryLyb2CQI0ZQk84Pr0qCx8Ea14W1xJLG2e5hkLLEyf8ALRO657HFe{+}{+}FfhC13b/NE3mMMgkdPWuw0TwPqGmzC2yDuIMbFenPbNQr7F2W589/ELTLvT5NH8daXA9vPaTxRX8RUq0Vwpykhx0OQFP4etfYHgj402Wl6zpXxf0B2v8AQtUhi07xTp1t8z27D5lmZM53xFmGf4kBx2rpo/hBZ{+}L/AAzqFrf2qTR3sXlTrjG/jg/UdQexFfGj{+}AvFnwP{+}KtxoelanLpc84KxXLFVgu4zkqX3/ACYI{+}U7uA2a0Whk9T9edOvLfX7K31CymjubW6RZYZoTuSRCMqQe4IrUS1Veq5OOtfnf8B/2mfG/wPZ/Dvirwbfy{+}HjM0kT2sTMtopPPljP3M5bapOM8Y4FfZnwy{+}PWlfFKUxaNFJMQB5kjo0Sr68OAeOe1aqxzShJbHozWgbiovsu3P8Xbp0rW2YfJOSo5oEQfA4GfSquYqTMz7GccgGmtbgZzwc1rCPjGM0ksIOPl6D1707j5mYE1viUKMfnVSa0bkg/L04rZuoWwpxye{+}KZLbMdqj8aVhprqczNZEv6fWituayLyE7CfoKKLD{+}ZZU0Z4NFKOlIxG455p9FFABQOtGaU880AOoxSA5paAAjFKBmkpcfrQAmMGgHFIWC9SPxqnearBaDLMPegC6TUU93HbJl3Cj61xus/ECG1jYQqWOOMCvGPHvxL164jmjt38iPHBUfMKTZooNnuHib4jaR4ftJZbq8jjVB3YZ96{+}Hf2uPinpHxLsha2MbSXlsfNt5zxtYZ4B9xj8cVyPj7xFqs5keeaa5fJBZnOBXimu63Pbv5sknfg7uTyDWDcmzeCUTwPxFB5OpTr1DncCBgHj07Gm2F68W1Ap3JwDnqDXQeLbNLu5aeJfmDHO0c1lGzBjjLrgOuN4HQ9s10XutSk{+}pWlZJUeWIFg5IeE4BDeoqjHZSMcsSMcq3Un61pWtnPDMS8fmGQYdc/eHqPetqXTBFZj5Vki6/MOVHoT/WpcuUpR5jlJtTvI12{+}YzJnAbqagu7{+}TykRkQPyfM28t9f8au3NmzSnyxhACSx5FZZKuQDjA4q1Z9DOV1pcrTTvdYDct6gVu{+}E7gWOoGa6i{+}0xiCVIoic/vWQqjY77SQ31UVnrCLgLFAP3jdTnoKtQ6lf6M6waa5s5FbbJdR8OzHqN/UL7Dr3zRKzXKhRTj7zOv8P8Ah1YNbsJP7NvLWfTWTUbl508shAQYguf4pHKqPqCOK{+}hfhZqV1rXge{+}mvUF14i015FW5Kl33u7SqxyCu7llBbkADb3ryT4f8AgnWPGUluLXU7m2EOHl1OcFzcSjO1VDdUTJxnqSTxwK9n{+}HWmXPg/xTcwajFEbme0K3E0W3y3xIuyXafuqd{+}CRkhgAOCK8LE1dbJ6o9zD0Wld9T2bQbu213wzDdQgDKfMmDgMByOQM4ORnGPlq14Yt4/PfeOjbeO1cZDOvg3xa9scR6Xq/wA0LsQqxzdcE{+}{+}MDJ712GlXaCcOjghueO9eHNfaWzPcgvsvoejR6dHdQDb8zHjkVlvpjaZfhlXCE9K1NBdpgOcD/wCtXQT6YJELHnisHqdEXys3dCiW5sVP3uOQeAahvNKRgw2YPrWl4HjG427A5PAyetdDdabGZdu0/XFUldHPOpyTszyy{+}0FuqD2yDWfFpGwknKHdwMda9Q1DREUFVTCnv71gXmkAAsSMHr7UnFpam0Kqlsc1DAUOQMLV8IxEe32zge1WUsRFIM5Unpmp7ez3FMHDKeMDis7G/NbU39FtP9EX5OT6Vfi01m2tjg45FT6THHBaL5hyCM4NSy6iUYEYUKcdK2SUVdnBJym3YSLTm8wkZOfU9KzvGmtjwnoTzRoJLqTKRKjDcT0wBnrycfSrUerASHlQMZ5OM{+}1eW{+}KvEC{+}J7m61SQtLpVuTBZRjBExyd0v8PY7V/PnNbxnGMXL{+}v6Rj7OUp8r26leB7jTNGuri7uDDdarIbRJ7gkCIeW0kkuD06L2xnPrWYJFufhb4j167KanG8D28STKEDKWw7HZjqxY8Y4x0p/jpodP0Lw2rXSwG1dppEeEurhyu8YxgcbupHAbHIALLC1tz8LNTsXGYLmOQorNtynPPtXLzaq52xhuxfBHitNQ0zwtLOtmky/abXYLdn4jnCrjrt4GdxP9417BqN/EY8segx1r5/8HzTweGtCMsC7bYyLcGG8EarMHiBJCg{+}YS3OOmeSepr0m71mfUAVhQgZ5kLbR{+}FbtvlivL9SOROcmjbgxqt{+}beBSXI/1in7v416X4f8AA0FtptuGjLyKCS7gZJrzHwtZ3CyxgzSSHeHcq{+}R9MV9GaNYNHZQ7sqdoyK{+}lwSThZnzOPk4SSRgmSPSIvuStjkJDCzc/lXH{+}KtPn1uWO71JRaabbN5sdsSC0rDozDtg4Prx6Zz6fqlkJMgjH06VwfizSfMtpF3MVI6E16NjzIyvqfFnxptTqerzCPIhBOOM4FedeBfB32/xFF5se{+}PfgqB/WvoX4heFiZ5AFG3JJJqh8JvBL/wBsPPJGwGMKT0xWDjqdaloep/Dv4eQiBcpwR8oPFeif8KksJ9jvAC6D1PXj/CtLwbpyW0cYKlSa7rA2EL0x1roS01OOU2nocfovguHT4DGEzzwTXmvxh{+}D{+}g{+}LfJOs6bFdRQPvV2QErxgr67TnkD2NfQkEQ8o5XrzzWZrGix6lbyI/8qLIhVGndnmHgbw7pej6FZ6bas11BbJiI3LGRlX{+}6SeSOOM16D4as7O3hVUtkifOdygCsWz8Ff2dJIYs7T0UHit2zsZbadF6qo60IbldG6xBVhzycdaeybAwHTGM4otoSAu4DOM1OE3jk9TTOcjWEhyRjjkj0odWKHoC3TFSRRkZw2QTiknYq4IHCjtQBUmj3cMvTuKiW3O3J4PvVwnzCqdSeTn0omjAODgYPQCgDOeXyzjaCOo70VZaMZyV69BjoKKdx3Zj08dKbtpehpCFooowKAFBxQRk8UmKcDmgAzjrQKZI6xgliMe9YOqeKIrYMsTbm9qTdhpN7G7NdxwruchRWLfeKIYDtjG41ytzqVxqJ5Y4zzUsFhI3LZI96V2zRRS1LNzrV7eONuVXNVmtZLvPmSFiK07axKoTjnpj0qzbaeVVgeWPUU0u5aOVutGEq8gKK43xH4XWWGQLHtjbqQOTXrc{+}mmVNvQnoKztT0PfCyeuaLIZ8Z/EPwOXhmIUnJPHY18zePvCN3byvtj2queMdM1{+}lms{+}BItQtpxtOV7EV4V41{+}FaS5LQjHbC/WiyKTPz8udLmhdiYwSeoPeqLWIVGBjA5xgHpX1vqPwCa{+}unZFKc9QM1Un/ZjkVA8cTs4/X/PFMdj5e0bSnuLlFaM7dwwT0Fdd4n8MCxsoWgCB2TdIXGQnp{+}desf8ACidT0m8VzaMoBz0Nc98QvDM6XluhDQAIAynAz74rmnK2rOulFvRHzf4iTbJ5UaDys/O6LjcenbpXKyW7K{+}1fmHUMO9fTVj4Hg1I{+}UIhIG4LqvU8/nXQWnwFtLyIE2qtxwwXFc7x0YaNHSsvnPW58m2W/TbyCd1OwN82BnjvXungT4a3PiaRNS0{+}W11O3cAlomBIOOjKOQfYiu8n/AGcocFfJwuM4IzWVcfs{+}rosMt7ZS3enyxoW32kzRHge1cdfFRrfA{+}VnZRwc6PxK6O60AwaFazuitJLaMYZoMFAsg/h5Xk/TI5HNYuja/f6x4k168bUog9lp8jm1IMZTJUCNMkHdlQwXnlN5{+}5VPw9od74ii0/Sb6{+}1f7Olu05VNRkAulZyVaUKRubkjJycYGcAAdzoPg2z0u/Flb2KQE2U4hjhT52PylvbHGWY88ADlq86nZTs3dnfUUpU02tjWvSvjzwXDG7FJnjDK6lgyOOhBPPB4ye4rU{+}H2qS6spivmMWp27eXcRjHzMAPnALEkNnPQDnA6VT8L6LNpGoyaQdqRyMZYF{+}6TuGT7nG36AY9a7K18OXHh7VIdaijd4lAjvo0H34e5x3K9R{+}NcqkneHc6nCSSkuh6R4fjZNo6YHcV3NlArpjBJbvWJpGnxQwoy5fKhlYqQGBGQefUYNdFZRBXTGVXqQf881zP3XZnRo1eJoaPD9lvI8KSOufT8a6d5AJCXXOT155rDtoxJKMHPv04{+}lX7qVtoJUEj9a0T5Vc4Jx55G9awJdR7T8wAxzWDrekxRqzKBt54pbfUiFPBX2HaqWo37KjDacHOea2lVTiZ06U4yumc9cKke9gc4J57AVShmDTgnhV5Ip2pXBYbvlQYxt7VhyX0jS4HIJ9c8VwuSPWUOh2Y1RrjaFPy47U9/MWMnqPas7w6pnCk8Y6mtzViLS3VEhFxNK2yOIhtkhBGVLL9w7ckH2Na01KqzCrKNLocxrTXF5G2m2ryRvMMXEsUjKyRnho2XHVuxB6E{+}ormzbRXWpPZwRAW1oFXCBCAw5K9CRxs7jj1r0tdIh0yxmmuJC3loXlmmfkgDuT6AY57AVxOkSm30ae/uN{+}zDzuW6oOWI6Ae3TsOtVWWqitlsKjNNNv1Z5D4wF3r/im8SC8Ki3lSFrZBlsEFWJ9Fzke{+}a6/wAXaZdXHhDU7WIsZ2s5I1UDuUOKteFWGu{+}M9RhW1mtbUn7SgKGIXC4XY74ciQg{+}ZgkDA2jsa9D1XTIbTRrm5kUiKONiQBntWfI27I1VWMY3Z4Z8KvDlxp8ccEsdzDDdyu8cYCPsVli3b2/hG5SMLyckZGCK9403wvHtDbS2APcVxXglIpfEui209rPbXCWvAlZd8Rd2LIwAGSdgGRwNp4yePZJwliqZIK9zjoa9H2d1G/Y82dflcrdSvpdlaRSRKYsOWAynNey2spS2jU4wFABrynQojf61ZKgG3eCSox/WvWEiZAATx24r6LBq0D5nGyvNIilG4EnkemKwdYtRcKVAGB6jrXTiMAFTzUEtmsgOBj3Nd55y3PGde8CR6tKzMm7nOcfzqx4b8DLYKxEQQA4GFxXqTaYibtwAI/Wli035QFGR1oNfaWVilpGniKNc88VtLAFUjOBiljtDDjgYqdUy3T/69Bk3cVEGOO3amTAyZGeM1MyjacDv1pgG05PHvQSyuYjhvkOKkS33OcqO2PepmO9MZOM05FwOhJ9M9aA2EYhcdsCl25VcdO9CjfuJGRjvT2YAYVeMfnQK5EqlHXjIJ7nvTJQXHI5JzgVIzgLu/i7H2pCdqnJBwOn8qAsRFjGu/G7jAxUbIeOSQepPWpXdeByT1xTC2yItxycAj07UANkBdyc/Wil83A6qP60UCMIjNA/WlHSmkc0AOpRxTVORSjrQAvTOaq3uoR2MZZ2AxRf3qWUDSOQMe9eaa74jkvZyindk8AUm7FpXNnV/Es145jiO1PY1QstNluW3OWwTnB70mlacWG5/v9ea6S3gEajGQcYzRbW5o9NiC10lQQpxnqRWlFarlQVzk9PaiNG80BRkhfyq/ANzDjJUHFMW5XWMlsKuAanhUAqgBz16VMsPzKAcYGc1Oi7NpHBAz0oBuxUNvtZeCRg9aqTWnz/dzmtpE2ur5PTv0pHVJI1kHIzj3oFzGEulJJbS/Lhj3ri/EfheO4jZfKHK8MB9a9RijWWNicE9qyrvTxLGwIGRx0oLTPFYfBxNz8kecYJBrsdO8GQ7CPKVxwdpHArpk0EKyuBkn2ratbHylwqgKy8YNBXMcNL8P7G6f5oFA57ZxXzv{+}0Z8Coje2F/bw4QqwbLBV6569utfZkVh5WXIGMZwe9edftD{+}HTq/gpfIUho5Q{+}VAJHHvWFaN4M2w9RxqJHw7ofg0aRdi3liwpPD4yPwr13QPC9rLGjgKB228ZNc3d2cuI1QSbkUB{+}ACp9x6c1taHrn2NfKkbBBxtr4{+}rK02j7amrwTOmfwbbSkEqAD0FNn{+}GFvqVtJEyAoQQ2R2xWvpOridFYkKVroNP1IbsklhjJwKzXKDckfNPwi8AWd7r{+}oz/AGeCOdVjt0EKbMxKCEyuwAtgDcwJ9D6nv/EnguHRdUsHkjA80Oi71UpnIPIxnOM429D8xztrU{+}Hs6p40trBNKubewsobpmvZhgOzT8ADGCmQ{+}Dk8gggYrtPi5crFpel3Nu7pLHdhSI3xlSrZz6gjjH{+}1W1L{+}JdmE52hyr{+}tTzzXPCqadZWOtx7g1tJtkCqxDRtwQQvLYO0hc4yBniu20jSEubfmPORk7xnP1rq4tCttQ0OfTpvm86MoQvA{+}Ydj{+}Nc14BMsWmyWU{+}1Z9PkazcKoCnbjBAA6Yx{+}Vcji002dftFJNCaHYjRmbSpGysWZLUY5MW75l4XqpYckng{+}3PRWsBmkGQFTI9ttO1Ww8{+}CKWPC3EEnmRPnoe4/EZ{+}nB7Va065S9ihmhVYxLyUzyh7r1OcHIPuOuQa2nHmXN16nPGfLePQ1VtViiyH4POcdfwqKa5IJ5LE9AB0{+}ta8ESyQYJ5x2HBrNv7UAEAAjsQOnrxUTi0tCISUnY524uZhJlMqc/wj0qCWZ2VmYc1dmtsEsRge1VEtv3mT8ykdx1rkaZ6ScUtDLv7aS5jBj{+}UD{+}HHWsS2sne5wdwGepHeu2uIf3HTBxjpiseOECdsnH49afINVNTodCijsrCSRsEINxJ4YnsB7mptCQ3yz6pOAXuci3DAqRFnKEqTw2CBx2Ue9YUCnXLxdPRithCN1yrLw/IKgHOQSR/3zn1FdjbzR2yTSSTQA24WSeN3AZYyeXGeuAGOBXox/dx5ep5NaWrm2YPj64a38PizWSWGa/nS2RlJ3c8sRjk/KDVPWNEli8Ialb2iK87Wzxjedu7K{+}pOBn1J4rf1XVfCetzI51WdJNMjkuzDLbtGsqIpLOu5RvwAB8pIyR9a{+}X/H/jzxL4k1WJmS5SykkZbPSLa3NykoX7{+}YhxIV4yzfKDxx0Oc5LmTIoSc4OMV63PbfhlCdU26pcafJZzxWaWR{+}UiFtrHbsJGW4JJPTkYz1rpfGVvPf{+}HriztAhlmwig9M7hxXyDo3xj8b{+}Evi9p8l/cXUtvdyxQXFrJb/Z0khJCfNEOA4HII5z3xkV9YfEDxRN4b8Pi9sIjLes6iACPfltpIO3vyBWkNUrjmmpcrRgeEpjd{+}OZbwRSw7UWENLglyAWbKjhclt49nxngivSrtzKxXHHoK86{+}GeL/Wbu/C4S5J2spJDCMLGQRuIBVgw6AkdiMGvTJYspg/LjqR711q99Tkk0nobHw9tCNVkIA2Khz2Oe1emHkD2rhvh4ARPKA2CduSfSu7jAA68CvpKCtTR83iHzVGIi5OTQIjuPGB2p4cfSgyAdDW5zMjnt1IHAzSxphh7CnsyuQM4p6YHQcUAyKYjIwe9BXC596eyqxGfyqJ{+}w5I96ADJC9TyetBjyRUYLHA4FWI48EEjI/lQDDOAvueooBz370EBVGcdeQO9KCHz35oJDrwOhPNAULHwe/pQY8A9yTmjBRTx1oArSnjApYwGQ7xnPGfSpAVJ5HbrTPu4AwQRigaIZAcFlUnPTdxRLIqqMHgdvSrDgeWHI{+}6OprOnkymSRk88UAWFkDDJAP1FFQROSnJB/GigRlHpxS9KKM4oAKjmmESknjFOcnBrnPE{+}rCyt2BYBcdKBpXOc8a{+}JVkYWytjJx161kaTpcrp5r{+}oGTWZGr6pdSzv8Adz0I{+}7Xa6dZ{+}dEfL6DB9s0G3kaNjaDY79hj8KvlGcnaDsIzU1tb7zsHygpnA74q9HCHkXHKle1AhtnAEkJ4yVqzHbjnHp{+}dLGnKt0yMA9RT1yWj4{+}ooFcjVNpQsMDHB/pUyuDt/vY5FSNEHVQMHnFCQBio2leO9AmJIFZE75H51HLGwtQOuG5qxtUMv{+}yOlTXEYeMKF69{+}lAitbRmPhj1HSkmhQsFx97g4qaJVaQAnpxUiqEfJHPrQHUqG1AVcYODU0aAAexxTpE4JHHOadwR8pHrQN6itEdvB56Vg{+}M7dbzQ7uJh5hEZIjx1xXQRsGJO4cGmTxRSo5bGfpSkrpoqGjTPjq70e3W9uAiN5ZbcBu4{+}lYGpaQ8D71TaPQHmvUPE{+}nrY6xdRoqRhZG5XvWFc2UchGQfUAd6{+}LxCtUaPusPK9NHIWF1LbuV64Az7V0llqbpGeN64zUTaGI5GmAwpGcCpLC12lgo3Ejn0rjOo4XRNWu7f4htf{+}Uk1gyywBnOHgnMmSV/2TGgGD6{+}5r0jxwwv/AA60YQzEOj{+}WrmNmwQcB/wCH615xf6XqV34o/sr54bOO6FxbsmxA00qlVDMWGU{+}UkjnGOOpz6B4ntN3hmYkK{+}xFb5n2hgMHg8Y{+}taq6lFmKs{+}aJ2GiahI9lA0vySCNVYMwODjkZ7/XvXPzzrovxFWeSTZb6xCIiCOsqkAc9uCB{+}JqTwlczXWiRmTyhJz8kEnmInJGFbPIHTNVPHlpIdAkuYWZbqzImjYZyMZBIxz0JPrwMUVPiaCCukzt9244xlc5PNY9/J/YWowXv8Ay7Tfu39VY9Mnsp56fxAf3jUnhTXF1vRLS7VQrOvzgHgN0Iz3GQefatC9to9Rs57W4RZIpUKMp96cZcrTIkubQ63TblZLJZEOFZdw4qvfZlYke{+}RntXBeBNfuNK1G78Pak7STod8M7YAlRj8pyWJZm{+}cnAAXb24ruJ5Sp55Hfmt5WtpsYRTjLUoXWCjBe3JzUEUXlYf04xirM8nlfMeSRiqryEMSRgA8DvXM4nWm0h10fkwO/rXF{+}INTFhMqRfNLM6xxptzvc5ITPUZAPPQck9K62SZCrM7YHJ2lwuT6ZPeuR8NaefEGuTeIDh7dsxWRZdjGLJIdh0zzj6D3q4xSXMxX1sdl4W0VdI02KEfvJG{+}eRj/Ef8P8ACppfGGm6Zc6vFrctvNLawpJDELXbMke4FQZD94M5wB6ge9aFofKQcc5HT/PvXlP7Sss1rZ{+}G2tiITcTuZJlXL4j27R7jLk49QK1S0uedXfPa5yfjn9oyOe336wsNlYNNugikgEgRcHDgdWI5yelcVL8ZbSHULuS31y6tLi0iwjm0KRhCwJxs3ZBbJwR/EKr/AAg/Z21T40Wd/wCIvGt08Gm3hYWtnCoR/J6JluvAA4BHSuT8cfAH4neA7mfTNE0XT/E{+}k5P2e8kZ0lCk5AcKy57flQ6bb52zphThrFS1N27{+}IC{+}PfiDo/ibW4IzZ6QggtoNnlte3eWaIFfUttJA6KhPevqLXbbHh/Tp7uOaU2qfaZI4TgybIywXjGMkAH2Jr460T9nHx/Dpx8U{+}LY1a{+}06SGWytbZziCN5FjdVjHGcvGQepw2T0r6k8Qa9e31joCxSy{+}a0DPcLC{+}FcKnz7vVcZ/HFPSFRRve5DV3ZdCx8Hgwa8depzJKMEAO7E8cnsBn39K9MklVwEX5ievt9a82{+}EsMun6NcmYjzgVRwG3DcBkkHJ4yT3P612lvqsbzhG3Ek8LjIz/n{+}VddNc07M56isrnp/hBEstNATkO26urhfcOOlZWiaasdnAAuBtB5rYVVQYA/Kvp4q0Uj5OT5pNi4Ap2A3Q/Wjdg7SM{+}9CkE8d6okjDFGIp6ytt4HFNI4J9{+}lKCdoHbrigB7rgZpincB/OpjjFQg7myePYUCF2gEccd81JnGBnkDNREnd17U4DdgjpQFh23eAMdf0qPlWCgYHXHpUw4x60gByGyOn50CGoxAHY5/KnF1XqceoNGeQOOe/r71G6k9RkjgH3oBDC4AJB6etPGCASOfWoHYE7fUVIu4KOcgdvagZFdBvIYKc5NY7z5kCZ6HFaN/PvYKTjHOKwrd994Ox3d6Ckrm0iAqDkZPPNFXFcIMBOKKDMwOe3NLg0Y49DQeB1oAhmJ2kA15d8S77bIkJYqGznB5xXpN5IUwf5V4r441aC/190JB8sbQP/r0FxVy7oExaM7uEZBgGu50p1jEewYUjBrgNLkeRFVTgL/D7V22k58kE7gM5AoLa7nTWAEOJM5bOCB1Aq8mVUENj5sZHWqFqWUscfMeQTV{+}NGVT/ABdOaAslqWlKlBzznqKeEG7B4APr0pIkBRuODggmph0Y7ffFBA3cAcZ79fX2qRcnjIBBPJFMeH/Z9/TNPJwQMYFArjfLLFSeCO3tUysGGMHI7moyfmGDz0pzZKcc59KAEJAkJI4zUsmF5U5Wq6gsAT27GrAAManFAxDtlC4z05pgVVAB65waQRlFLcjP5U1X3N3I65oAcLcnA6AnOM0NFgHnIIqcBjgAg96jlzGp/iPagDxL4iaV9n12QouWkAZhnr{+}H51yJtAC3GPQH1r1v4oaYZLNLtIczA7N2ePxrzBSxOw5Hsa{+}Vx8OWpfufX4Cpz0kZrwFN{+}eVPGP8ACkt7LyZNu0nNajxgfgen9aQgTYCrtIOeh5rzLHq3PJ/ilY3OoX26zuWsJtN8q9E8aFmIBdT095FH416Ett9u8GlZXG8wFXccYIz/ACrnPiRE2nt/a6zrarFbSLkRLJlgpYAhvlx8pzkHjtW/4NMV34LiijeN41R4tyPvXAJ79{+}vWrstDJPXQqfDK4iu9DxBJBIIXKu0AwpJAb8{+}fU565Oa6m9hWWB1PzBlwVPpXFfCGIxLqtqzAlJAVQQ7GAJI5P8XQDPtjmu/ubb5SBgH6VVVe{+}7CpytE4T4cyy6RqeraO6YFuwdGAzkHv{+}Ix9SGPXNd8WVj94njpXnnjOA{+}H9Z07xBEmPKdYbgKoYlDx378kdR1HpXodo8V3axTRNvikQMjr0IIyD{+}VZJGjetzm/HenzGyXVbA/wDEx0/LoF6yRkYePn{+}8uRn1rc8I{+}KovFGiW98pDiVATgEAHGGHIHQ5H4VZbGzHBzkc{+}leVTWkvw58bG8hJXQ9RlHmKBkQyHqRyAM1pBpaMHDmWm/wDWh6/PIxbjgDsOarNuIJGRS2dzHdWqzRSK6t0KEMPfkcHnI4rJ8a66NE0lIYIhdaleMEtrckrv5wSCDkYyD6YHvSSu7MzT1VjL1i{+}PiTUV0C2bMLoJNQdCGXysgrGQRlXLKf8AgJ967jTrKOCOJY0AjUbQMcAdqwPCnh/{+}ybI{+}fIbq9uGM11cSHLSSn7xJ9PSusjGyMjk/Tn8qLtlTtH3YliLBUYGDnNeK/HrWo/EN9H4fgV01XSF{+}2q20jfBIMMV9drJHn/f{+}texK4ijd3fAUE5z0FfLHxZ8K{+}I7y0f4iaE0n9qxXb3EcW0MHgxtKEdSrJgFT71blpZGMaSqOzfp6n0h8Mte0zUfA2my6aUEKx{+}XJEnBikHDIR2wfzGDW5JMsxDMNx4wM8mviDwB8S7bXru3/AOEZ8Qr4R16eRY5tLvmGwy8A8P8ALIpPQE5Fdd4q/aP8Y6PC1qp06CWFVjnufsjIyygYkXazEDa25fwz3p6tWRk4Om2prU99{+}IPiuPS1sNJjlRbrUZlyrchERlfJxyBlV/WuYuBNo{+}jNNC9v5xU2IlmIASJo23v6qpKpyvODjvivNfgxp2ueO9VuPE2qTXGyRPL{+}1NkBxkFvL9c4AyOAM8k8D0jxUhWe4hQx2lu9tDZRyOhIeSSRm8sfKwJKwkc8VlSi5Vky7KMH5m/4QsVtvDNsI2BHJ6bcDOAMdsY6dvwrtvCWmS6hrMWV/chgTzn8K5nRd{+}m6TbQzAOQgBwck/UjrXpHwztnur55mUiJBgH1J9q9rBwUqnMcOMqONPQ9QtU2wgDqMVL9xCO9KpCoNvApjHPfk19CfKoY2QM9c0{+}InJJ/CnLgLjvSp93FAC4DEA9qbsIJIFKxwSBT8gLjGBmgTGBuR6Un3TtHHv6UrEA8cYpqY5Lcn{+}dAWHIArEEdqcxXJJPIHbvTQSW9AaSRh2wMCgRIhVgTzk9P8KDtAHp/nilQAA9uOlNK5I9KBDZe2OMf5xSs3yjJ596exxgkVWuGBQ9aBkIcGRiMFelBmCMVbIUjt6UkMQbd6noTTHDebtGDg/eNAFK/laLzSCNm3HPU8isrTRmdSzA4Oc1Jrk37uZcndlRx9ai0yEs4L8oD0oNFojpon3oMLkDgEminW6lohhtvtiigyMMHNIRyaXocdqbIcAkGgDL1yUW1hNK33VUk5r5uub5NV1CeYuRGHIPUYr3jx5dvB4fvGBwojYmvnHSYGiWaaL51c/Mr9D7ig1gdno{+}peQUDNzjZz39K9L8Mv9rhiywLbSDk{+}leDT6nNbQSYBLRkED2r134eagNSt7eRXXBfDAHpxVWLPSoINwUk7iRgVpWql02HgY/GqFgsikDPAOOelayIY41PcHGKkhtWJY14BB4x/kU7BIAzye1MUbe/U4pCwUKV/KgzJskDa2DkdaaMHpyTzQrbwex7j1pqAq{+}OueOlAIZcDZt4ozsfkHHr7Us0m{+}QbjTZPnGQe350FEjsNhAJIHPSnQs0seBwq1TZnV9vY9xVi3k2nb2I6igCdX3jaR7fhmgREHcRlW4pq/KTg08uyYUgkYzxQA6AY6njpT88/hUY2nIycdaGHlnjJGetAHOeO7RdQ0GeMnA4YkdsV43PbeUThsg9{+}v4V7rqsH2uzkjZSVIIrw6/j8i7kDIFCsQMPkn8O1ePmELpM93LZfZIeCm0c4PPekUDOAAfpTDIWBHA54wO1PB3McZzivmGz6dI4r4vRBNGtZvKS4jExWSCRsI4ZWUgn6Metafwxs4LXw/BZIqB4IwZGgkWSFmbLZjKk5XBAGTk4/AafidrddGklvIo54oWWQh1BBIYYBBrkfg1catDd61Y6vEYbiKdjE{+}4FWjZ32YA4XhTwMD2rRHPN9jR8FEad4r1iwElwVlzIqMuYlIP8J7HB5GPTk8Ad7lmwMZwMcda4K3jaz{+}Jz/u5mDwtukik{+}QA4wHQ59PvYGOmTnj0TysDB7elbTW3oQnuZOtaOuqafPbMVXep2yEfdPY/gefwrE{+}HurfaNNuNOlHl3Fi5jKsV{+}7k46dshh0Fday7VQE9Bn61wur2h8LeNLbVQzLZX37mdEVdu84AYk4I7dPSsZdzaNtUdi4DIeMc{+}maxvEWiW2u6VcWd0A8UoIJP8Pv8AyrcVPmyzZGM81g{+}Jvs5hgt7u4a3tJnIlkQqp8sKS3LegBPGTx0qoQdSaiuonVVOLm{+}hwfw3{+}J9np0OuafqGqWOqWmhtvkvI9QR0jUvtO9txC8kZXrn6113g63m8RTr4ovsP9qjDWiqVKbCBmQY4ywUcj2rT8F/CX4ZaXo3iCx0zQtCubOUeRqMbvHcqWJUKjZ4DlkBHGTuY9Rirmn{+}H7HwTLaaFpSQ2{+}kwxeXb26A/uVyWVVzyQFIHPp04r1sXg/Z01KLvbc8zC4znqSi1a{+}xsRMFZRjBPerakgVFDbJwNwOPSrCqQDk8/WvHR6LOW{+}Imom00NbSJcz6hItogORndw347c9eOlatvpMFtosFgMNHGgRsqPm9cgVh30n9uePoLfyy0Omxb3crlfNfBC88AgbTxnr2rsIoX3L0PbrQ1d3KbUUkeBfGD4LeGLTSb3xDp/h8XGvKVhtI7ZDuaZ84YYIwQEc57cHqBW94f/Zr8B6Lrl3qH9kpfagJ2aWW9czBZs/PgMf72etb/wASvGcvh3V9Jt7cB/JkF68QiZ2lXBi2kgEAbpE565YYB5x2WhxXCaNDLeqq39xmaZMdHclmH4EmhXbsTzySu2Ri1S2Ty40CBRgBRwK8x8RO114p09baQtCk0yypt484BGK4PT93JGwcHpuHcivV7v8AdRuCwfIP515/qM0kmqw2a3qu0121zJCigiKN9giJbrnbEx2/7YJHFdFJJO5jNtpHQzxStHkj7oAVxzx0AAr1n4TW7/2M05G4s3BORivMpZoki3KxA{+}6oHpXt/gu3Wz8O2iKQVKBj9TzXs4BXbZ5GYyaikdAmSvNIUIAYdenFLv6elIz5PBxXtHgAgKHHX1IpzqTj09qWNPlPORTRlcjoaABHGMHtTpX{+}Wq/O/qKeWJAGPpQFhgLMxGMmoNSivZdKvl0{+}dLe/eB1tpZF3KkpU7CR3AODirkJAJJ9akPJJA{+}ooEfPmhfGbxX41v9Nt7BItNi8UFW0aWS33tbC2jY3/AJoPpKqov{+}/S{+}FvjL4o8Xano1vHBDYReJZIo9Mka33GA2qBtTEgPU7w8ae65r6B2qzcfezwaYyBM7jnnH06cUxHgfhP40eKfEGraLaSQQxf21NBp9vtg4jubUxNqoJJ9GuY0HZrVjzmrHxA{+}IUWl{+}MtV1vR/FVjbIPC0V1YRtsmj1KZZ7kpCmT824jbtj{+}c8YIwc{+}6RMFJJ6/wAqNm6QYQYHI46UCseUXnxB1WwsPGms6pqL2ljp{+}pRaTa2tvbwgwNIttiSSSRguVeYgsxVAuSQcCuZtPivf6rofh62k8XWmla1d6peWyyzNaiKezguNpmckYZtmxF8ogM8gIG0Er9ASKBGVZAVYYIIyCKhcLIoGOBQOx86eKPjLf{+}FPBuoSWd0LTWLW61{+}7iidYRBdLb6hcRxxHzW3uTtGViG7nOV{+}UHvvH/iiN/EHhGKw8ULpdqdck0/UPJaEqZfsskiQSFwcMTswvBO8d8V6PKqM6Lt5ByM9qhmQBU2orDdkkjrg//WoA8c8D{+}NdX8X6leQXwji/smNNP1ALHtEupKzCbb3CBVjdfUTj0r1XTIlDAgcngVi3wDXhJ5{+}cc5rcs32xgKef1pGj2NNYMjnOf97FFEcZKD5tvrkUUGRiNUbd6KKAOB{+}KhK{+}F73BIyvUdeteDeGZ3ktpGY5zKVI9R70UU0aRL2t2yRWtwq5Aj6fpXT/CG9kiKwqFEYYMBj6f40UUyz3aykYtMemGHArYVice/NFFSZMeTiMnvkCnNGBx{+}NFFBIkKBj1PrTj1/GiigY2UBsEgHJxUVud5wenIoooKJ5IVwG7jiowdrpj1xRRQBalXDZHWoyxIXJz8tFFAEoG1Cw4OKCxIUnqRRRQBFccpj1HXvXhHi2If21dnJwrkBT0FFFefjf4R6uX/xDIjOCOO1TRudiEcZ54oor5GW59gtjN8ZRpL4W1JWUEGBs{+}/y5rhPh9qlzaeJ5LIyGdLnT4tSlknJZ2lYuMZ7KOcAAdTnPGCiqRjPodLq9uJvHWmXxZxNHAJFCtwMtsIx6YY16AHLQjsRgAj8DRRW89l6GEeoA4A7njk1h{+}ONOh1Lw1dQzbtq4dWB5Vs9RRRU9Gax3LOlSvLpFgztvkaFdzkAFjjqccV598bdOi1LR4rWUssbxTZMZ2kHZnOaKKcHZpi3k16niUPwT8P8AgX9m7xnYabNf7ryS0uZLt5lE26PeV5VVBHzN1B619Ix2g0XxVpelQSSNYpZrPHHKdxjyCmwHrtAUYBJx2oor0swlJ06d33PPy{+}MVUqadjurcfdxxk44{+}hq/CoPB5BOOfpRRXlI72cR8PMTXHiC5dQ1w{+}oSq0hzkhQNoP0zgV20JJRCepJooq6exFX4zzzxJKL744{+}GrGaNHt7LQby{+}QY5aUysvzHuAAMD1Ar0C6uHaPqASwGQOeRRRRH7RjDoY{+}pMWDR5IVsA4NeYG3Fr8RtTjSSQot40ADOSAqRgDjoDjAz7D3yUVcdn6G/2kej2EC3OqW0T5KF14HpnpXvViAkSqowqrwB0FFFe9gPgZ89mPxo0YjkA{+}opQN0mD0oor1TxyeMc47Diq0zHIOetFFAkJGNwJJzinkZH1oooBj84K{+}9PxndyaKKBMSNByMdASKcyjdjFFFAIiKhCQPXFCkgj6ZoooBiyOSSM8VCrHgduaKKBjFUNICeuP61Wu2OdoJAJzxRRTQjlrljJfoCeAe30rfsQCVGOMUUUi5bGrCA0YJA/CiiigzP/2Q==')");



        //    driver.Quit();


        //}

        //////Tirar foto
        ////driver.SwitchTo();
        ////    driver.FindElement(By.Id("btnCamera")).Click();
        ////    Thread.Sleep(200);
        ////    driver.FindElement(By.CssSelector("#pnlCapturePreview > div > div.section-content.buttons > div > div.imagebutton.accept.dark")).Click();
        ////    Thread.Sleep(20000);

        ////    //Bot�o Avan�ar
        ////    driver.FindElement(By.Id("avancarBiometriaFacial")).Click();

        ////    //Cenario 7 -- Formaliza��o 

        ////    //Validar P�gina Outros

        ////    wait(By.Id("avancarBiometriaFacial"));


        ////    Assert.IsTrue(IsElementPresent(By.Id("panelAssinaturaEletronica")));

        ////    driver.Quit();


        ////}


        //[Test]
        //public void CT008_Cadastro_de_Clientes_Pensonistas()
        //{




        //    //M�todo de Print
        //    ScreenShot print = new ScreenShot("VendaCartoes", "CT001_Cadastro_de_Clientes");

        //    //M�todo simulador de teclado
        //    InputSimulator simulator = new InputSimulator();

        //    //Execu��o

        //    //Acessar P�gina do CCM
        //    driver.Navigate().GoToUrl(baseURL);


        //    wait(By.XPath("//button[@type='submit']"));
        //    wait(By.CssSelector("img.png_bg"));
        //    print.PrintScreen();

        //    try
        //    {
        //        Assert.AreEqual("Login", driver.FindElement(By.CssSelector("strong")).Text);
        //    }
        //    catch (AssertionException e)
        //    {
        //        verificationErrors.Append(e.Message);
        //    }
        //    Assert.IsTrue(IsElementPresent(By.Id("login_username")));
        //    Assert.IsTrue(IsElementPresent(By.Id("login_password")));
        //    Assert.IsTrue(IsElementPresent(By.XPath("//button[@type='submit']")));



        //    //Validar Tela de Login
        //    Assert.AreEqual("PSF Security", driver.Title);
        //    Assert.IsTrue(IsElementPresent(By.CssSelector("img.png_bg")));
        //    try
        //    {
        //        Assert.AreEqual("Portal Lojas", driver.FindElement(By.Id("tituloPagina")).Text);
        //    }
        //    catch (AssertionException e)
        //    {
        //        verificationErrors.Append(e.Message);
        //    }

        //    print.PrintScreen();

        //    //Realizar login
        //    Login();


        //    wait(By.Id("btnSearch"));
        //    print.PrintScreen();


        //    ////Alterar Filial
        //    Thread.Sleep(2000);
        //    SendKeys.SendWait("{ESC}");
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("labelFilial")).Click();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("changeFilial")).SendKeys("54 - L 054 JD / SP");
        //    Thread.Sleep(2000);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);
        //    Thread.Sleep(2000);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);
        //    Thread.Sleep(5000);


        //    //Validar P�gina Inicial
        //    Assert.AreEqual("PSF Security", driver.Title);

        //    try
        //    {
        //        Assert.AreEqual("Buscar por:\r\nCPF\r\nCART�O", driver.FindElement(By.Id("tipoSearch")).Text);
        //    }
        //    catch (AssertionException e)
        //    {
        //        verificationErrors.Append(e.Message);
        //    }
        //    Assert.IsTrue(IsElementPresent(By.Id("cpfSearch")));
        //    Assert.IsTrue(IsElementPresent(By.Id("btnSearch")));
        //    try
        //    {
        //        Assert.AreEqual("Concessao do Cartao", driver.FindElement(By.Id("menuCollapse3162")).Text);
        //    }
        //    catch (AssertionException e)
        //    {
        //        verificationErrors.Append(e.Message);
        //    }
        //    try
        //    {
        //        Assert.AreEqual("Atendimento", driver.FindElement(By.Id("menuCollapse3164")).Text);
        //    }
        //    catch (AssertionException e)
        //    {
        //        verificationErrors.Append(e.Message);
        //    }


        //    //Cenario 1 - Cadastro de Clientes

        //    //Consultar Cliente
        //    driver.FindElement(By.Id("cpfSearch")).Clear();
        //    driver.FindElement(By.Id("cpfSearch")).SendKeys("8947536300");
        //    Thread.Sleep(2000);
        //    print.PrintScreen();
        //    driver.FindElement(By.Id("btnSearch")).Click();
        //    Thread.Sleep(2000);



        //    wait(By.Id("iniciarCadastro"));


        //    //Validar P�gina Cadastro de Clientes


        //    Assert.IsTrue(IsElementPresent(By.Id("formPreCadastro")));

        //    driver.FindElement(By.Id("dataNascCad")).Click();
        //    driver.FindElement(By.Id("dataNascCad")).SendKeys("10/01/2001");
        //    driver.FindElement(By.Id("iniciarCadastro")).Click();


        //    //Validar Mensagem de Retorno CPF

        //    Thread.Sleep(8000);


        //    wait(By.Id("avancarDadosPessoais"));

        //    //Cenario 2  -- Dados pessoais


        //    //Validar P�gina Cadastro de Clientes


        //    Assert.IsTrue(IsElementPresent(By.Id("formDadosPessoais")));


        //    //Nome da M�e
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("nomeMaeDadosPessoais")).Click();
        //    driver.FindElement(By.Id("nomeMaeDadosPessoais")).Clear();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("nomeMaeDadosPessoais")).SendKeys("Maria Aparecida");


        //    //Estado Civil
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("estadoCivil")).Click();
        //    Thread.Sleep(2000);
        //    //new SelectElement(driver.FindElement(By.Id("estadoCivil"))).SelectByValue(Cadastro.EstadoCivil);
        //    driver.FindElement(By.Id("estadoCivil")).SendKeys("CASADO");// PEGAR DO BANCO
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);



        //    //Sexo       
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("sexoDadosPessoais")).Click();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("sexoDadosPessoais")).SendKeys("MASCULINO"); // PEGAR DO BANCO
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);


        //    //Deseja Receber a Fatura por email?         
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("faturaEmailOutros")).Click();
        //    driver.FindElement(By.Id("faturaEmailOutros")).SendKeys("N�o"); // PEGAR DO BANCO //


        //    //Email
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("cliEmail")).Click();
        //    driver.FindElement(By.Id("cliEmail")).Clear();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("cliEmail")).SendKeys("starline@starline.com.br");//PEGAR NO BANCO

        //    Thread.Sleep(2000);

        //    print.PrintScreen();

        //    //Bot�o avan�ar
        //    driver.FindElement(By.Id("avancarDadosPessoais")).Click();

        //    //Cenario 3 -- Dados Residenciais 

        //    //Validar P�gina Dados Residenciais

        //    wait(By.Id("avancarDadosResidenciais"));


        //    Assert.IsTrue(IsElementPresent(By.Id("formDadosResidenciais")));

        //    //Cep
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("cepDadosResidenciais")).Click();
        //    driver.FindElement(By.Id("cepDadosResidenciais")).Clear();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("cepDadosResidenciais")).SendKeys("06033050");
        //    Thread.Sleep(2000);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);


        //    //Numero
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("numeroDadosResidenciais")).Click();
        //    driver.FindElement(By.Id("numeroDadosResidenciais")).Clear();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("numeroDadosResidenciais")).SendKeys("1010");

        //    //Tipo Residencia
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("tipoResidencia")).Click();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("tipoResidencia")).SendKeys("ALUGADA"); // PEGAR DO BANCO
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);


        //    //Telefone Residencia
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("telefoneResidencialDadosResidenciais")).Click();
        //    driver.FindElement(By.Id("telefoneResidencialDadosResidenciais")).Clear();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("telefoneResidencialDadosResidenciais")).SendKeys("1135926538");
        //    Thread.Sleep(2000);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);



        //    print.PrintScreen();

        //    //Bot�o avan�ar
        //    //Thread.Sleep(2000);
        //    //driver.FindElement(By.Id("avancarDadosResidenciais"));

        //    //Cenario 4 -- Dados Comerciais 

        //    //Validar P�gina Dados Comerciais

        //    wait(By.Id("avancarDadosComerciais"));


        //    Assert.IsTrue(IsElementPresent(By.Id("formDadosComerciais")));


        //    //Classe Profissisional
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("classeProfDadosComerciais")).Click();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("classeProfDadosComerciais")).SendKeys("Pensonistas"); // PEGAR DO BANCO
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);

        //    //Atividade
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("atividadeDadosComerciais")).Click();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("atividadeDadosComerciais")).SendKeys("Pensonistas"); // PEGAR DO BANCO
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);


        //    //Renda Mensal
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("rendaMensalDadosComerciais")).Click();
        //    driver.FindElement(By.Id("rendaMensalDadosComerciais")).Clear();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("rendaMensalDadosComerciais")).SendKeys("10000");


        //    ////Tempo de Empresa (ANO)
        //    //Thread.Sleep(2000);
        //    //driver.FindElement(By.Id("tempoEmpresaAnoDadosComerciais")).Click();
        //    //driver.FindElement(By.Id("tempoEmpresaAnoDadosComerciais")).Clear();
        //    //Thread.Sleep(2000);
        //    //driver.FindElement(By.Id("tempoEmpresaAnoDadosComerciais")).SendKeys("10");



        //    ////Tempo de Empresa (MES)
        //    //Thread.Sleep(2000);
        //    //driver.FindElement(By.Id("tempoEmpresaMes")).Click();
        //    //driver.FindElement(By.Id("tempoEmpresaMes")).Clear();
        //    //Thread.Sleep(2000);
        //    //driver.FindElement(By.Id("tempoEmpresaMes")).SendKeys("10");


        //    //CEP
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("cepDadosComerciais")).Click();
        //    driver.FindElement(By.Id("cepDadosComerciais")).Clear();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("cepDadosComerciais")).SendKeys("06033050");

        //    //Numero
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("numeroDadosComerciais")).Click();
        //    driver.FindElement(By.Id("numeroDadosComerciais")).Clear();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("numeroDadosComerciais")).SendKeys("365");

        //    //Telefone Comercial
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("telefoneComercialDadosComercial")).Click();
        //    driver.FindElement(By.Id("telefoneComercialDadosComercial")).Clear();
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("telefoneComercialDadosComercial")).SendKeys("1136811452");
        //    Thread.Sleep(2000);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);

        //    //Bot�o avan�ar
        //    //Thread.Sleep(2000);
        //    //driver.FindElement(By.Id("avancarDadosComerciais"));

        //    //Cenario 5 -- Outros 

        //    //Validar P�gina Outros

        //    wait(By.Id("avancarOutros"));


        //    Assert.IsTrue(IsElementPresent(By.Id("formOutros")));

        //    //Vencimento Melhor Dia
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("vencimentoOutros")).Click();
        //    driver.FindElement(By.Id("vencimentoOutros")).SendKeys("10");
        //    Thread.Sleep(2000);

        //    //Confirma��o de Vencimento
        //    SendKeys.SendWait("{ESC}");
        //    Thread.Sleep(2000);
        //    SendKeys.SendWait("{ESC}");
        //    Thread.Sleep(2000);
        //    //Cobrar tarifa de overlimit


        //    wait(By.CssSelector("#tarifaOverlimitOutros"));

        //    driver.FindElement(By.CssSelector("#tarifaOverlimitOutros")).Click();
        //    driver.FindElement(By.CssSelector("#tarifaOverlimitOutros")).SendKeys("N�o");

        //    //Conta B�nus?
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("tipoCartaoOutros")).Click();
        //    driver.FindElement(By.Id("tipoCartaoOutros")).SendKeys("N�o");


        //    //Campanha
        //    Thread.Sleep(2000);
        //    driver.FindElement(By.Id("campanhaOutros")).Click();
        //    driver.FindElement(By.Id("campanhaOutros")).SendKeys("931 - Desc Primeira Compra Marisa Itaucard 10%");
        //    Thread.Sleep(2000);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);


        //    //Bot�o avan�ar
        //    //Thread.Sleep(2000);
        //    driver.FindElement(By.Id("avancarOutros")).Click();

        //    //Cenario 6 -- Camera 

        //    //Validar P�gina Outros

        //    wait(By.Id("avancarBiometriaFacial"));






        //    //   Assert.IsTrue(IsElementPresent(By.Id("fieldSetBiometria")));

        //    //Ligar Camera
        //    Thread.Sleep(2000);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.TAB);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);
        //    SendKeys.SendWait("+^(i)");
        //    System.Threading.Thread.Sleep(2000);
        //    SendKeys.SendWait("+^(p)");
        //    System.Threading.Thread.Sleep(2000);
        //    SendKeys.SendWait("(base 64)");
        //    Thread.Sleep(2000);
        //    simulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);



        //    driver.FindElement(By.Id("campanhaOutros")).Click();



        //    driver.Quit();


        //}

        //////Tirar foto
        ////driver.SwitchTo();
        ////    driver.FindElement(By.Id("btnCamera")).Click();
        ////    Thread.Sleep(200);
        ////    driver.FindElement(By.CssSelector("#pnlCapturePreview > div > div.section-content.buttons > div > div.imagebutton.accept.dark")).Click();
        ////    Thread.Sleep(20000);

        ////    //Bot�o Avan�ar
        ////    driver.FindElement(By.Id("avancarBiometriaFacial")).Click();

        ////    //Cenario 7 -- Formaliza��o 

        ////    //Validar P�gina Outros

        ////    wait(By.Id("avancarBiometriaFacial"));


        ////    Assert.IsTrue(IsElementPresent(By.Id("panelAssinaturaEletronica")));

        ////    driver.Quit();


        ////}


        private bool IsElementPresent(bool v)
        {
            throw new NotImplementedException();
        }

        private bool IsElementPresent(By by)
        {
            try
            {
                driver.FindElement(by);
                return true;
            }
            catch (NoSuchElementException)
            {
                return false;
            }
        }

        public bool verify(string elementName)
        {
            try
            {
                driver = null;
                return true;
            }
            catch
            {
                return false;
            }
        }

        private bool isElementNotPresent(By by)
        {
            try
            {
                driver.FindElement(by);
                return true;
            }
            catch (NoSuchElementException)
            {
                return false;
            }
        }



        private bool IsAlertPresent()
        {
            try
            {
                driver.SwitchTo().Alert();
                return true;
            }
            catch (NoAlertPresentException)
            {
                return false;
            }
        }

        private string CloseAlertAndGetItsText()
        {
            try
            {
                IAlert alert = driver.SwitchTo().Alert();
                string alertText = alert.Text;
                if (acceptNextAlert)
                {
                    alert.Accept();
                }
                else
                {
                    alert.Dismiss();
                }
                return alertText;
            }
            finally
            {
                acceptNextAlert = true;
            }
        }
    }
}

