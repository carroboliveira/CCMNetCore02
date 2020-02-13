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
        public string TipoBrowser = null;

        [SetUp]
        public void SetupTest(string[] pTipoBrowser)
        {
            TipoBrowser = pTipoBrowser[0];
            processTest = new ProcessTest();
            InternetExplorerDriverService IEService = null;
            ChromeDriverService ChService = null;
            if (TipoBrowser == "IE")
            {
                var options = new InternetExplorerOptions();
            }
            else
            {
                var options = new ChromeOptions();
            }

            var os = Environment.OSVersion;
            OSLinux = os.Platform == PlatformID.Unix;
            if (OSLinux)
            {
                PathChromeDriver = "/usr/bin/";
                if (TipoBrowser=="IE")
                {
                    IEService = InternetExplorerDriverService.CreateDefaultService(PathChromeDriver, "IEDriverServer");
                }
                else
                {
                    ChService = ChromeDriverService.CreateDefaultService(PathChromeDriver, "chromedriver.exe");
                }
            }
            else
            {
                PathChromeDriver = @"C:\Projetos\CCMNetCore02\CCMNetCore02\bin\Debug\netcoreapp2.1";
                if (TipoBrowser == "IE")
                {
                    IEService = InternetExplorerDriverService.CreateDefaultService(PathChromeDriver, "IEDriverServer.exe");
                }
                else
                {
                    ChService = ChromeDriverService.CreateDefaultService(PathChromeDriver, "chromedriver.exe");
                }

                //IWebDriver driver = new InternetExplorerDriver(@"C:\Projetos\CCMNetCore02\CCMNetCore02\bin\Debug\netcoreapp2.1\IEDriverServer.exe");
                //driver.Navigate().GoToUrl("http://www.google.com");
            }


            //processTest.CustomerName = "marisa"; //CustomerName evitar Acentuacao e espa�o
            //processTest.SuiteName = "ccm"; //CustomerName evitar Acentuacao e espa�o
            //processTest.ScenarioName = "cadastroaposentado";
            //processTest.ReportID = 271;

            if (TipoBrowser == "IE")
            {
                var options = new InternetExplorerOptions();
                driver = new InternetExplorerDriver(IEService, options);
            }
            else
            {
                var options = new ChromeOptions();
                //options.AddArgument("--headless");   
                //options.AddArgument("--no-sandbox");
                //options.AddArgument("--disable-dev-shm-usage");
                //options.AddArgument("--touch-events=enabled");
                //options.AddArgument("start-maximized");
                //options.EnableMobileEmulation("iPhone 5/SE");
                driver = new ChromeDriver(ChService, options);
            }

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

