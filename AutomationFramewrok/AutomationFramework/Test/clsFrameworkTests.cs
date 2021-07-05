using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutomationFramework.Test
{
    [TestFixture]
    class clsFrameworkTests : clsWebBrowser
    {
        public bool blStop;
        public clsReportResult clsRR = new clsReportResult();
        public clsWebElements clsWE = new clsWebElements();

        [OneTimeSetUp]
        public void BeforeClass()
        {
            blStop = clsReportResult.fnExtentSetup();
            if (!blStop)
                AfterClass();
        }

        [SetUp]
        public void SetupTest()
        {
            clsReportResult.objTest = clsReportResult.objExtent.CreateTest(TestContext.CurrentContext.Test.Name);
            fnOpenBrowser("Chrome");
        }

        [Test]
        public void Test_DBExecution()
        {
            clsDB objDBOR = new clsDB();
            objDBOR.fnOpenConnection(objDBOR.GetConnectionString("lltcsed1dvq-scan", "1521", "viaonei", "oferia", "P@ssw0rd!0"));
            DataTable datatable = new DataTable();
            string strQuery = "SELECT * FROM viaone.claim WHERE file_num = '402103288B80001'";
            datatable = objDBOR.fnDataSet(strQuery);
            foreach (DataRow row in datatable.Rows)
            {
                Console.WriteLine("Key: {0}, value: {1}", Convert.ToString(row["KEY"]), Convert.ToString(row["VALUE"]));
            }

        }


        [Test]
        public void Test_fnPageLoadPass()
        {
            fnNavigateToUrl("https://www.amazon.com.mx/");
            bool blResult = fnPageLoad(fnGetWe("//input[@id='twotabsearchtextbox']"), "Amazon", true, true);
            if (blResult)
                clsReportResult.fnLog("PageLoadResult", "PageLoad Pass Return Values", "Info", false);
        }

        [Test]
        public void Test_fnPageLoadFail()
        {
            fnNavigateToUrl("https://www.google.com/");
            fnPageLoad(fnGetWe("//input[@id='twotabsearchtextbox1']"), "Amazon", true, true);
        }

        [Test]
        public void Test_fnSendKeysPass()
        {
            fnNavigateToUrl("https://www.amazon.com.mx/");
            fnSendKeys(fnGetWe("//input[@id='twotabsearchtextbox']"), "Search TextBox", "Monitor Samsung");
        }

        [Test]
        public void Test_fnSendKeysFail()
        {
            fnNavigateToUrl("https://www.amazon.com.mx/");
            fnSendKeys(fnGetWe("//*[@title='Buscar123']"), "Search TextBox", "Selenium with C#");
        }

        [Test]
        public void Test_fnClickPass()
        {
            fnNavigateToUrl("https://www.amazon.com.mx/");
            fnClick(fnGetWe("//a[@href='/gp/goldbox?ref_=nav_cs_gb']"), "Promociones Section");
        }

        [Test]
        public void Test_fnClickFail()
        {
            fnNavigateToUrl("https://www.amazon.com.mx/");
            fnClick(fnGetWe("//a[@href='/gp/goldbox?ref_=nav_cs_gb123']"), "Promociones Section");
        }

        [Test]
        public void Test_fnClearPass()
        {
            fnNavigateToUrl("https://www.amazon.com.mx/");
            fnSendKeys(fnGetWe("//input[@id='twotabsearchtextbox']"), "Search TextBox", "Monitor Samsung");
            fnClear(fnGetWe("//input[@id='twotabsearchtextbox']"), "Clear Search TextBox");
        }

        [Test]
        public void Test_fnClearFail()
        {
            fnNavigateToUrl("https://www.google.com/");
            fnSendKeys(fnGetWe("//*[@title='Buscar']"), "Search TextBox", "Selenium with C#");
            fnClear(fnGetWe("//*[@title='Buscar1']"), "Search TextBox");
        }

        [Test]
        public void Test_fnContainsTextPass()
        {
            fnNavigateToUrl("https://www.amazon.com.mx/");
            string textContains = fnGetWe("(//div[@class='a-cardui-header'])[1]").Text;
            fnContainsText("Contains Envio Instructions", textContains, "Gratis en tu", true, true);
        }

        [Test]
        public void Test_fnContainsTextFail()
        {
            fnNavigateToUrl("https://www.amazon.com.mx/");
            string textContains = fnGetWe("(//div[@class='a-cardui-header'])[1]").Text;
            fnContainsText("Contains Envio Instructions", textContains, "Free en tu", true, true);
        }

        [Test]
        public void Test_fnElementExistPass()
        {
            fnNavigateToUrl("https://www.amazon.com.mx/");
            fnElementExist("Mejoras para el Hogar", "(//div[@class='a-cardui-header'])[2]", true, true);
        }

        [Test]
        public void Test_fnElementExistFail()
        {
            fnNavigateToUrl("https://www.amazon.com.mx/");
            fnElementExist("Mejoras para el Hogar", "(//div[@class='a-cardui-header123'])[2]", true, true);
        }

        [Test]
        public void Test_fnElementNotExistPass()
        {
            fnNavigateToUrl("https://www.amazon.com.mx/");
            fnElementNotExist("Promo Not Exist", "(//div[@class='a-cardui-header123'])[3]", true, false);
            fnNavigateToUrl("https://www.google.com/");
        }

        [Test]
        public void Test_fnElementNotExistFail()
        {
            fnNavigateToUrl("https://www.amazon.com.mx/");
            fnElementNotExist("Favoritos de Amazon Estados Unidos NOT Exist", "(//div[@class='a-cardui-header'])[3]", true, false);
            fnNavigateToUrl("https://www.google.com/");
        }

        [Test]
        public void Test_fnDoubleClickPass()
        {
            fnNavigateToUrl("https://www.amazon.com.mx/");
            fnDoubleClick(fnGetWe("//a[@id='nav-cart']"), "Carrito");
        }

        [Test]
        public void Test_fnDoubleClickFail()
        {
            fnNavigateToUrl("https://www.amazon.com.mx/");
            fnDoubleClick(fnGetWe("//a[@id='nav-cart1']"), "Carrito");
        }

        [Test]
        public void Test_GetAttributePass()
        {
            fnNavigateToUrl("https://www.amazon.com.mx/");
            string resultGA = fnGetAttribute(fnGetWe("//a[@id='nav-cart']"), "URL redirection", "href");
            Console.WriteLine(resultGA);
        }

        [Test]
        public void Test_GetAttributeFail()
        {
            fnNavigateToUrl("https://www.amazon.com.mx/");
            string resultGA = fnGetAttribute(fnGetWe("//a[@id='nav-cart123']"), "URL redirection", "href");
            Console.WriteLine(resultGA);
        }

        [Test]
        public void Test_ScrollToPass()
        {
            fnNavigateToUrl("https://www.amazon.com.mx/");
            fnPageLoad(fnGetWe("//a[@href = '/ref=nav_logo']"), "Amazon", true, true);
            fnScrollTo(fnGetWe("(//div[@class='navFooterLinkCol navAccessibility'])[4]"), "Metodos de Pago Section", true, true);
        }

        [Test]
        public void Test_ScrollToFail()
        {
            fnNavigateToUrl("https://www.amazon.com.mx/");
            fnPageLoad(fnGetWe("//a[@href = '/ref=nav_logo']"), "Google", true, true);
            fnScrollTo(fnGetWe("(//div[@class='navFooterLinkCol navAccessibility'])[5]"), "Metodos de Pago Section", true, true);
        }

        [Test]
        public void Test_fnVerifyTextPass()
        {
            fnNavigateToUrl("https://www.amazon.com.mx/");
            string verifyText = fnGetWe("(//div[@class='a-cardui-header'])[1]").Text;
            fnVerifyText("Verify Envio Instructions", verifyText, "Envío Gratis en tu primer pedido elegible", true, true);
        }

        [Test]
        public void Test_fnVerifyTextFail()
        {
            fnNavigateToUrl("https://www.amazon.com.mx/");
            string verifyText = fnGetWe("(//div[@class='a-cardui-header'])[1]").Text;
            fnVerifyText("Verify Envio Instructions", verifyText, "Envío Gratis en tu primer pedido", true, true);
        }

        [Test]
        public void Test_fnSelectListPass()
        {
            fnNavigateToUrl("https://www.nylaarp.com/Life-Insurance/Term/Protection");
            fnSelectList("Coverage 25000", fnGetWe("//select[@id='CoverageAmount']"), "25000", "value", true, true);
        }

        [Test]
        public void Test_fnSelectListFail()
        {
            fnNavigateToUrl("https://www.nylaarp.com/Life-Insurance/Term/Protection");
            fnClick(fnGetWe("//select[@id='CoverageAmount']"), "Coverage DropDown");
            fnSelectList("Coverage 8500", fnGetWe("//select[@id='CoverageAmount']"), "8500", "value", true, true);
        }

        [Test]
        public void Test_fnSelectedItemPass()
        {
            fnNavigateToUrl("https://www.nylaarp.com/Life-Insurance/Term/Protection");
            fnSelectList("Coverage 25000", fnGetWe("//select[@id='CoverageAmount']"), "25000", "value", true, true);
            fnVerifySelectedItem("Coverage 25000", fnGetWe("//select[@id='CoverageAmount']"), "25000", true, true);
        }

        [Test]
        public void Test_fnSelectedItemFail()
        {
            fnNavigateToUrl("https://www.nylaarp.com/Life-Insurance/Term/Protection");
            fnSelectList("Coverage 25000", fnGetWe("//select[@id='CoverageAmount']"), "25000", "value", true, true);
            fnVerifySelectedItem("Coverage 25000", fnGetWe("//select[@id='CoverageAmount']"), "8500", true);
            fnSelectList("Coverage 10000", fnGetWe("//select[@id='CoverageAmount']"), "10000", "value", true, true);
        }

        [Test]
        public void Test_fnCheckBoxPass()
        {
            fnNavigateToUrl("https://www.amazon.com.mx/");
            fnSendKeys(fnGetWe("//input[@id='twotabsearchtextbox']"), "Search TextBox", "Monitor Samsung");
            fnClick(fnGetWe("//input[@type='submit']"), "Buscar Producto", false);
            fnSelectCheckBox("Select CheckBox Marca: Samsung", fnGetWe("//div[@id='brandsRefinements']/ul/li[contains(., 'SAMSUNG')]//i"));

        }

        [Test]
        public void Test_fnCheckBoxFail()
        {
            fnNavigateToUrl("https://www.amazon.com.mx/");
            fnSendKeys(fnGetWe("//input[@id='twotabsearchtextbox']"), "Search TextBox", "Monitor Samsung");
            fnClick(fnGetWe("//input[@type='submit']"), "Buscar Producto", false);
            fnSelectCheckBox("Select CheckBox Marca: Samsung", fnGetWe("//div[@id='brandsRefinements']/ul/li[contains(., 'SAMSUNG')]//input"), true, true);
            fnNavigateToUrl("https://www.google.com/");
        }

        [Test]
        public void Test_fnSelectRadioBtnPass()
        {
            fnNavigateToUrl("https://www.keynotesupport.com/internet/web-contact-form-example-radio-buttons.shtml");
            fnScrollTo(fnGetWe("(//p[contains(. , 'SELECT A SOFTWARE PRODUCT')])[1]"), "Radio Buttons section");
            fnSelectRadioBtn("Select Radio Button Software Product: QuickBooks Pro", fnGetWe("//span[contains(., 'QuickBooks Pro')]/input[@type='radio']"));
        }

        [Test]
        public void Test_fnSelectRadioBtnFail()
        {
            fnNavigateToUrl("https://www.keynotesupport.com/internet/web-contact-form-example-radio-buttons.shtml");
            fnScrollTo(fnGetWe("(//p[contains(. , 'SELECT A SOFTWARE PRODUCT')])[1]"), "Radio Buttons section");
            fnSelectRadioBtn("Select Radio Button Software Product: QuickBooks Pro", fnGetWe("//span[contains(., 'Test')]/input[@type='radio']"));
        }

        [Test]
        public void Test_fnVerifyListDropDownPass()
        {
            fnNavigateToUrl("https://www.nylaarp.com/Life-Insurance/Term/Protection");
            fnClick(fnGetWe("//select[@id='CoverageAmount']"), "Coverage section");
            string[] strListTest = new string[4] { "$100,000", "$50,000", "$25,000", "$10,000" };
            fnVerifyList("Verify Coverage List", fnGetWe("//select[@id='CoverageAmount']"), strListTest, "option", true, false);
        }

        [Test]
        public void Test_fnVerifyListDropDownFail()
        {
            fnNavigateToUrl("https://www.nylaarp.com/Life-Insurance/Term/Protection");
            fnClick(fnGetWe("//select[@id='CoverageAmount']"), "Coverage section");
            string[] strListTest = new string[4] { "$100,000", "$30,000", "$25,000", "$10,000" };
            fnVerifyList("Verify Coverage List", fnGetWe("//select[@id='CoverageAmount']"), strListTest, "option", true, false);
        }

        [Test]
        public void Test_fnVerifyListCheckBoxPass()
        {
            fnNavigateToUrl("https://www.amazon.com.mx/");
            fnSendKeys(fnGetWe("//input[@id='twotabsearchtextbox']"), "Search TextBox", "Monitor 4k");
            fnClick(fnGetWe("//input[@type='submit']"), "Buscar Producto", false);
            string[] strListTest = new string[4] { "SAMSUNG", "LG", "AOC", "BenQ" };
            fnScrollTo(fnGetWe("//div[@id='brandsRefinements']/ul"), "Brand Section");
            fnVerifyList("Verify Marca List", fnGetWe("//div[@id='brandsRefinements']/ul"), strListTest, "li", true, false);
        }

        [Test]
        public void Test_fnVerifyListCheckBoxFail()
        {
            fnNavigateToUrl("https://www.amazon.com.mx/");
            fnSendKeys(fnGetWe("//input[@id='twotabsearchtextbox']"), "Search TextBox", "Monitor 4k");
            fnClick(fnGetWe("//input[@type='submit']"), "Buscar Producto", false);
            string[] strListTest = new string[4] { "SAMSUNG", "TestBrand", "AOC", "BenQ" };
            fnScrollTo(fnGetWe("//div[@id='brandsRefinements']/ul"), "Brand Section");
            fnVerifyList("Verify Marca List", fnGetWe("//div[@id='brandsRefinements']/ul"), strListTest, "li", true, false);
        }

        [TearDown]
        public void CloseTest()
        {
            fnCloseBrowser();
            clsReportResult.fnExtentClose();
        }

        [OneTimeTearDown]
        public void AfterClass()
        {
            try
            {
                clsReportResult.objExtent.Flush();

            }
            catch (Exception objException)
            {
            }
        }
    }
}
