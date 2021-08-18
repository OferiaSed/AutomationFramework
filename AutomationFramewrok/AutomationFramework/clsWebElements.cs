using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.Extensions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace AutomationFramework
{
    public class clsWebElements
    {
        public static DefaultWait<IWebDriver> objFluentWait;
        public static WebDriverWait objExplicitWait;
        public static string strAction = "";


        public IWebElement fnGetWe(string pstrLocator)
        {
            try
            {
                IWebElement pobjElement = clsWebBrowser.objDriver.FindElement(By.XPath(pstrLocator));
                return pobjElement;
            }
            catch (Exception pobjException)
            {
                fnExceptionHandling(pobjException, "WebElement: " + pstrLocator + " doesn't exist", true);
                return null;
            }
        }

        public object fnGetFluentWait(IWebElement pobjWebElement, string pstrAction, string pstrTextEnter = "")
        {
            objFluentWait = new DefaultWait<IWebDriver>(clsWebBrowser.objDriver);
            objFluentWait.Timeout = TimeSpan.FromSeconds(10);

            objFluentWait.PollingInterval = TimeSpan.FromMilliseconds(250);
            objFluentWait.IgnoreExceptionTypes(typeof(WebDriverTimeoutException), typeof(SuccessException));

            switch (pstrAction)
            {
                case "Displayed":
                    objFluentWait.Until(x => pobjWebElement.Displayed);
                    break;
                case "Click":
                    objFluentWait.Until(x => pobjWebElement).Click();
                    break;
                case "SendKeys":
                    objFluentWait.Until(x => pobjWebElement).SendKeys(pstrTextEnter);
                    break;
                case "Clear":
                    objFluentWait.Until(x => pobjWebElement).Clear();
                    break;
            }
            return objFluentWait;
        }

        public bool fnGetExplicitWait(string pstrLocator, string pstrOption)
        {
            bool pblStatus = false;
            IWebElement objWebElement;
            try
            {
                objExplicitWait = new WebDriverWait(clsWebBrowser.objDriver, TimeSpan.FromSeconds(10));
                objExplicitWait.IgnoreExceptionTypes(typeof(WebDriverTimeoutException));

                switch (pstrOption)
                {
                    case "ElementExists":
                        objWebElement = objExplicitWait.Until(ExpectedConditions.ElementExists(By.XPath(pstrLocator)));
                        pblStatus = true;
                        break;
                    case "Elementvisible":
                        objWebElement = objExplicitWait.Until(ExpectedConditions.ElementIsVisible(By.XPath(pstrLocator)));
                        pblStatus = true;
                        break;
                }
                return pblStatus;

            }
            catch (Exception pobjException)
            {
                fnExceptionHandling(pobjException);
                return pblStatus;
            }
        }

        public bool fnPageLoad(IWebElement pobjWebElement, string pstrPage, bool pblScreenShot = true, bool pblHardStop = false, string pstrHardStopMsg = "PageLoad Failed and HardStop defined")
        {
            clsReportResult clsRR = new clsReportResult();

            bool blResult = false;
            try
            {
                clsReportResult.fnLog("PageLoad", "Step - PageLoad in Page: " + pstrPage, "Info", false);

                strAction = "Displayed";
                IJavaScriptExecutor objJS = (IJavaScriptExecutor)clsWebBrowser.objDriver;
                clsWebBrowser.wait.Until(wd => objJS.ExecuteScript("return document.readyState").ToString() == "complete");

                fnGetFluentWait(pobjWebElement, strAction);
                clsReportResult.fnLog("PageLoadPass", "The Page is loaded for the Page: " + pstrPage, "Pass", pblScreenShot);
                blResult = true;
            }
            catch (Exception pobjException)
            {
                clsReportResult.fnLog("PageLoadFail", "The Page is not loaded for the Page: " + pstrPage, "Fail", true, pblHardStop, pstrHardStopMsg);
                fnExceptionHandling(pobjException);
            }
            return blResult;
        }

        public bool fnSendKeys(IWebElement pobjWebElement, string pstrField, string pstrTextEnter, bool pblScreenShot = true, bool pblHardStop = false, string pstrHardStopMsg = "SendKeys Failed and HardStop defined")
        {
            clsReportResult clsRR = new clsReportResult();

            bool blResult = false;

            try
            {
                clsReportResult.fnLog("SendKeys", "Step - Sendkeys: " + pstrTextEnter + " to field: " + pstrField, "Info", false);
                strAction = "SendKeys";
                fnGetFluentWait(pobjWebElement, strAction, pstrTextEnter);
                clsReportResult.fnLog("SendKeysPass", "The SendKeys for: " + pstrField + " with value: " + pstrTextEnter + " was done successfully.", "Pass", pblScreenShot, pblHardStop);
                blResult = true;
            }
            catch (Exception pobjException)
            {
                clsReportResult.fnLog("SendKeysFail", "The SendKeys for: " + pstrField + " with value: " + pstrTextEnter + " has failed.", "Fail", true, pblHardStop, pstrHardStopMsg);
                fnExceptionHandling(pobjException);
            }
            return blResult;
        }

        public bool fnClick(IWebElement pobjWebElement, string pstrElement, bool pblScreenShot = true, bool pblHardStop = false, string pstrHardStopMsg = "Click Failed and HardStop defined")
        {
            bool blResult = false;
            try
            {
                clsReportResult.fnLog("Click", "Step - Click on " + pstrElement, "Info", false);

                strAction = "Click";
                fnGetFluentWait(pobjWebElement, strAction);
                //clsReportResult.fnLog("ClickPass", "The click to the element is correctly for: " + pstrElement, "Pass", pblScreenShot);
                clsReportResult.fnLog("ClickPass", "Click on " + pstrElement + " was done successfully.", "Pass", pblScreenShot);
                blResult = true;
            }
            catch (Exception pobjException)
            {
                //clsReportResult.fnLog("ClickFail", "The click to the element is not working for: " + pstrElement, "Fail", true, pblHardStop, pstrHardStopMsg);
                clsReportResult.fnLog("ClickFail", "Click on " + pstrElement + " was done successfully.", "Fail", true, pblHardStop, pstrHardStopMsg);
                fnExceptionHandling(pobjException);
            }
            return blResult;
        }

        public bool fnDoubleClick(IWebElement pobjWebElement, string pstrElement, bool pblScreenShot = true, bool pblHardStop = false, string pstrHardStopMsg = "DoubleClick Failed and HardStop defined")
        {
            bool blResult = false;
            try
            {
                //clsReportResult.fnLog("DoubleClick", "Step - Double Clic on " + pstrElement, "Info", false);

                strAction = "Displayed";
                fnGetFluentWait(pobjWebElement, strAction);
                Actions actions = new Actions(clsWebBrowser.objDriver);
                actions.DoubleClick(pobjWebElement).Perform();
                //clsReportResult.fnLog("DoubleClickPass", "The Double click to the element is correctly for: " + pstrElement, "Pass", pblScreenShot);
                clsReportResult.fnLog("DoubleClickPass", "Double Click on " + pstrElement + " was done successfully.", "Pass", pblScreenShot);
                blResult = true;
            }
            catch (Exception pobjException)
            {
                //clsReportResult.fnLog("DoubleClickFail", "The Double click to the element is not working for: " + pstrElement, "Fail", true, pblHardStop, pstrHardStopMsg);
                clsReportResult.fnLog("DoubleClickPass", "Double Click on " + pstrElement, "Info", false);
                clsReportResult.fnLog("DoubleClickPass", "Couldn't Double Click on " + pstrElement, "Fail", true, pblHardStop, pstrHardStopMsg);
                fnExceptionHandling(pobjException);
            }
            return blResult;

        }

        public string fnGetAttribute(IWebElement pobjWebElement, string pstrElement, string pstrAttName, bool pblScreenShot = true, bool pblHardStop = false, string pstrHardStopMsg = "GetAttribute Failed and HardStop defined")
        {
            string strAttributeContent = "";
            try
            {
                clsReportResult.fnLog("GetAttribute", "Step - Get Attribue " + pstrAttName + " from " + pstrElement, "Info", false);

                strAttributeContent = pobjWebElement.GetAttribute(pstrAttName);
                //clsReportResult.fnLog("GetAttributePass", "The Attribute " + pstrAttName + " from the WebElement: " + pstrElement + "  is captured successfully", "Pass", pblScreenShot);
                clsReportResult.fnLog("GetAttributePass", "Get Attribute " + pstrAttName + " from Element: " + pstrElement + "  was done successfully.", "Pass", pblScreenShot);

            }
            catch (Exception pobjException)
            {
                //clsReportResult.fnLog("GetAttributeFail", "The Attribute " + pstrAttName + " from the WebElement: " + pstrElement + "  is not captured successfully", "Fail", pblScreenShot, pblHardStop, pstrHardStopMsg);
                clsReportResult.fnLog("GetAttributeFail", "Get Attribute " + pstrAttName + " from Element: " + pstrElement + ".", "Info", false);
                //clsReportResult.fnLog("GetAttributeFail", "Couldn't Capture Get Attribute " + pstrAttName + " from Element: " + pstrElement + ".", "Fail", pblScreenShot, pblHardStop, pstrHardStopMsg);
                fnExceptionHandling(pobjException);
            }
            return strAttributeContent;
        }

        public bool fnClear(IWebElement pobjWebElement, string pstrField, bool pblScreenShot = true, bool pblHardStop = false, string pstrHardStopMsg = "Clear Failed and HardStop defined")
        {
            bool blResult = false;
            try
            {
                //clsReportResult.fnLog("Clear", "Step - Clear to field: " + pstrField, "Info", false);

                strAction = "Clear";
                fnGetFluentWait(pobjWebElement, strAction);
                Thread.Sleep(3000);
                string testTXT = pobjWebElement.Text;
                if (pobjWebElement.Text.Equals("") || pobjWebElement.Text.Equals(null))
                {
                    //clsReportResult.fnLog("ClearPass", "The text is cleared for the field: " + pstrField, "Pass", pblScreenShot);
                    clsReportResult.fnLog("ClearPass", "Clear to field" + pstrField + " was done successfully.", "Pass", pblScreenShot);
                    blResult = true;
                }
                else
                    throw new ArgumentException("ClearFail");
            }
            catch (Exception pobjException)
            {
                fnExceptionHandling(pobjException);
            }
            return blResult;
        }

        public bool fnContainsText(string pstrStepName, string pstrParentString, string pstrSubString, bool pblScreenShot = true, bool pblHardStop = false, string pstrHardStopMsg = "Contains Failed and HardStop defined")
        {
            bool blResult = false;
            try
            {
                clsReportResult.fnLog("Contains", "Step - Contains: " + pstrSubString + " in the string: " + pstrParentString, "Info", false);

                if (pstrParentString.Contains(pstrSubString))
                {
                    clsReportResult.fnLog("ContainsTextPass", pstrStepName, "Pass", pblScreenShot);
                    blResult = true;
                }
                else
                    throw new ArgumentException("ContainsTextFail");
            }
            catch (Exception pobjException)
            {
                fnExceptionHandling(pobjException);
            }
            return blResult;
        }

        public bool fnElementExist(string pstrStepName, string pstrLocator, bool pblScreenShot = true, bool pblHardStop = false, string pstrHardStopMsg = "Element Exist Failed and HardStop defined")
        {
            bool blResult = false;
            try
            {
                bool blElementExist = fnGetExplicitWait(pstrLocator, "ElementExists");

                if (blElementExist)
                {
                    clsReportResult.fnLog("ElementExistPass", "Element " + pstrStepName + " exist in the page", "Pass", pblScreenShot);
                    blResult = true;
                }
                else
                    throw new ArgumentException("ElementExistFail");
            }
            catch (Exception pobjException)
            {
                fnExceptionHandling(pobjException);
            }
            return blResult;
        }

        public bool fnElementNotExist(string pstrStepName, string pstrLocator, bool pblScreenShot = true, bool pblHardStop = false, string pstrHardStopMsg = "Element Not Exist Failed and HardStop defined")
        {
            bool blResult = false;
            try
            {

                bool blElementExist = fnGetExplicitWait(pstrLocator, "ElementExists");

                if (!blElementExist)
                {
                    clsReportResult.fnLog("ElementNotExistPass", "Element " + pstrStepName + " not exist in the page", "Pass", pblScreenShot);
                    blResult = true;
                }
                else
                    throw new ArgumentException("ElementNotExistFail");

            }
            catch (Exception pobjException)
            {
                fnExceptionHandling(pobjException);
            }
            return blResult;
        }

        public bool fnElementExistNoReport(string pstrStepName, string pstrLocator, bool pblScreenShot = true, bool pblHardStop = false, string pstrHardStopMsg = "Element Exist Failed and HardStop defined")
        {
            bool blResult = false;
            try
            {
                bool blElementExist = fnGetExplicitWait(pstrLocator, "ElementExists");

                if (blElementExist)
                {
                    //clsReportResult.fnLog("ElementExistPass", "Element " + pstrStepName + " exist in the page", "Pass", pblScreenShot);
                    blResult = true;
                }
                else
                    blResult = false;
                //throw new ArgumentException("ElementExistFail");
            }
            catch (Exception pobjException)
            {
                blResult = false;
                //fnExceptionHandling(pobjException);
            }
            return blResult;
        }


        public bool fnVerifyText(string pstrStepName, string pstrExpectedString, string pstrActualString, bool pblScreenShot = true, bool pblHardStop = false, string pstrHardStopMsg = "Verify Text Failed and HardStop defined")
        {
            bool blResult = false;
            try
            {
                clsReportResult.fnLog("VerifyText", "Step - Verify Text, Expected: " + pstrExpectedString + ", Actual: " + pstrActualString, "Info", false);

                if (pstrExpectedString.Equals(pstrActualString))
                {
                    clsReportResult.fnLog("VerifyTextPass", pstrStepName, "Pass", pblScreenShot);
                    blResult = true;
                }
                else
                    throw new ArgumentException("VerifyTextFail");
            }
            catch (Exception pobjException)
            {
                fnExceptionHandling(pobjException);
            }
            return blResult;
        }

        public bool fnSelectList(string pstrStepName, IWebElement pobjWebElement, string pstrValue, string pstrValueType = "value", bool pblScreenShot = true, bool pblHardStop = false, string pstrHardStopMsg = "Select List Failed and HardStop defined")
        {
            bool blResult = false;
            try
            {
                clsReportResult.fnLog("SelectDropdown", "Step - Select Dropdown: " + pstrStepName + " With Value: " + pstrValue, "Info", false);

                var selectElement = new SelectElement(pobjWebElement);
                switch (pstrValueType.ToUpper())
                {
                    case "VALUE":
                        selectElement.SelectByValue(pstrValue);
                        break;
                    case "TEXT":
                        selectElement.SelectByText(pstrValue);
                        break;
                    case "INDEX":
                        selectElement.SelectByIndex(Int32.Parse(pstrValue));
                        break;
                }
                clsReportResult.fnLog("SelectListPass", pstrStepName, "Pass", pblScreenShot);
                blResult = true;

            }
            catch (Exception pobjException)
            {
                clsReportResult.fnLog("SelectListFail", pstrStepName, "Fail", true, pblHardStop, pstrHardStopMsg);
                fnExceptionHandling(pobjException);
            }
            return blResult;
        }

        public bool fnVerifySelectedItem(string pstrStepName, IWebElement pobjWebElement, string pstrValue, bool pblScreenShot = true, bool pblHardStop = false, string pstrHardStopMsg = "Selected List Failed and HardStop defined")
        {
            bool blResult = false;
            try
            {
                clsReportResult.fnLog("VerifySelectedDropdown", "Step - Selected Dropdown: " + pstrStepName + " With Value: " + pstrValue, "Info", false);
                string pstrSelectItem = new SelectElement(pobjWebElement).SelectedOption.GetAttribute("value");
                if (pstrValue == pstrSelectItem)
                {
                    clsReportResult.fnLog("SelectedItemPass", pstrStepName, "Pass", pblScreenShot);
                    blResult = true;
                }
                else
                    throw new ArgumentException("VerifySelectedItemFail");
            }
            catch (Exception pobjException)
            {
                fnExceptionHandling(pobjException, pstrStepName, pblHardStop, pstrHardStopMsg);
            }
            return blResult;
        }

        public bool fnSelectCheckBox(string pstrStepName, IWebElement pobjWebElement, bool pblScreenShot = true, bool pblHardStop = false, string pstrHardStopMsg = "Select CheckBox Failed and HardStop defined")
        {
            bool blResult = false;
            try
            {
                clsReportResult.fnLog("SelectCheckBox", "Step - " + pstrStepName, "Info", false);
                strAction = "Click";
                fnGetFluentWait(pobjWebElement, strAction);
                Thread.Sleep(5000);
                clsReportResult.fnLog("SelectCheckBoxPass", pstrStepName, "Pass", pblScreenShot);
                blResult = true;
            }
            catch (Exception pobjException)
            {
                clsReportResult.fnLog("SelectCheckBoxFail", pstrStepName, "Fail", true, pblHardStop, pstrHardStopMsg);
                fnExceptionHandling(pobjException);
            }
            return blResult;

        }

        public bool fnSelectRadioBtn(string pstrStepName, IWebElement pobjWebElement, bool pblScreenShot = true, bool pblHardStop = false, string pstrHardStopMsg = "Select Radio Button Failed and HardStop defined")
        {
            bool blResult = false;
            try
            {
                clsReportResult.fnLog("SelectRadioBtn", "Step - " + pstrStepName, "Info", false);
                strAction = "Click";
                fnGetFluentWait(pobjWebElement, strAction);
                Thread.Sleep(5000);

                if (pobjWebElement.Selected)
                {
                    clsReportResult.fnLog("SelectRadioBtnPass", pstrStepName, "Pass", pblScreenShot);
                    blResult = true;
                }
                else
                    throw new ArgumentException("SelectRadioBtnFail");

            }
            catch (Exception pobjException)
            {
                fnExceptionHandling(pobjException, pstrStepName, pblHardStop, pstrHardStopMsg);
            }
            return blResult;
        }

        public bool fnVerifyList(string pstrStepName, IWebElement pobjWebElement, string[] pstrListValues, string pstrListType = "option", bool pblScreenShot = true, bool pblHardStop = false, string pstrHardStopMsg = "Verify List Failed and HardStop defined")
        {
            bool blResult = false;
            try
            {
                bool blElementsFound = true;

                clsReportResult.fnLog("VerifyList", "Step - " + pstrStepName, "Info", false);
                IList<IWebElement> objListWE = pobjWebElement.FindElements(By.TagName(pstrListType));

                for (int i = 0; i < pstrListValues.Count(); i++)
                {
                    bool blElementFound = false;
                    foreach (IWebElement objWE in objListWE)
                    {
                        if (pstrListValues[i] == objWE.Text)
                        {
                            blElementFound = true;
                        }
                    }

                    if (blElementFound)
                        clsReportResult.fnLog("VerifyListItems", "Element from the List: " + pstrListValues[i] + " is Displayed", "Info", false);
                    else
                    {
                        clsReportResult.fnLog("VerifyListItems", "Element from the List: " + pstrListValues[i] + " is not Displayed", "Info", false);
                        blElementsFound = false;
                    }
                }

                if (blElementsFound)
                {
                    clsReportResult.fnLog("VerifyListPass", "All Elements from the List are Displayed", "Pass", pblScreenShot, pblHardStop);
                    blResult = true;
                }
                else
                    throw new ArgumentException("VerifyListFail");

            }
            catch (Exception pobjException)
            {
                clsReportResult.fnLog("VerifyListPass", "Some Elements from the List are not Displayed", "Pass", pblScreenShot, pblHardStop);
                fnExceptionHandling(pobjException, pstrStepName, pblHardStop, pstrHardStopMsg);
            }

            return blResult;
        }

        public static void fnExceptionHandling(Exception pobjException, string pstrStepName = "", bool pblHardStop = false, string pstrHardStopMsg = "Failed Step and HardStop defined")
        {
            clsReportResult clsRR = new clsReportResult();

            //string pstrExceptionName = pobjException.GetType().Name;
            switch (pobjException.Message.ToString())
            {
                case "SendKeysFail":
                    clsReportResult.fnLog("SendKeysFail", "SendKeys action Fail", "Fail", true, pblHardStop, pstrHardStopMsg);
                    break;
                case "ClearFail":
                    clsReportResult.fnLog("ClearFail", "Clear action Fail", "Fail", true, pblHardStop, pstrHardStopMsg);
                    break;
                case "ElementExistFail":
                    clsReportResult.fnLog("ElementExistFail", "Element exist verification failed", "Fail", true, pblHardStop, pstrHardStopMsg);
                    break;
                case "ElementNotExistFail":
                    clsReportResult.fnLog("ElementNotExistFail", "Element not exist verification failed", "Fail", true, pblHardStop, pstrHardStopMsg);
                    break;
                case "ContainsTextFail":
                    clsReportResult.fnLog("ContainsTextFail", "Contains text verification failed", "Fail", true, pblHardStop, pstrHardStopMsg);
                    break;
                case "VerifyTextFail":
                    clsReportResult.fnLog("VerifyTextFail", "Verify text verification failed", "Fail", true, pblHardStop, pstrHardStopMsg);
                    break;
                case "VerifySelectedItemFail":
                    clsReportResult.fnLog("SelectedItemFail", "Coverage selected verification failed", "Fail", true, pblHardStop, pstrHardStopMsg);
                    break;
                case "SelectRadioBtnFail":
                    clsReportResult.fnLog("SelectRadioBtnFail", "Select Radio Button verification failed", "Fail", true, pblHardStop, pstrHardStopMsg);
                    break;
                case "Timed out after 10 seconds":
                    if (pobjException.InnerException.ToString().Contains("no such element: Unable to locate element"))
                        clsReportResult.fnLog("NoSuchElement", "WebElement doesn't exist or is incorrect", "Info", false);
                    break;
                default:
                    clsReportResult.fnLog("Exception", "Exception: " + pobjException.Message.ToString(), "Fail", true, pblHardStop, pstrHardStopMsg);
                    Console.WriteLine("Exception: " + pobjException.GetType().Name);
                    break;
            }
        }

        public void fnScrollTo(IWebElement pobjWebElement, string pstrField, bool pblScreenShot = true, bool pblHardStop = false, string pstrHardStopMsg = "Scroll To Failed and HardStop defined")
        {
            clsReportResult clsRR = new clsReportResult();

            try
            {
                clsReportResult.fnLog("ScrollTo", "Step - Scroll to element: " + pstrField, "Info", false);
                clsWebBrowser.objDriver.ExecuteJavaScript("arguments[0].scrollIntoView(true);", pobjWebElement);
                Thread.Sleep(TimeSpan.FromSeconds(2));
                clsReportResult.fnLog("ScrollToPass", "Scrolled to element: " + pstrField, "Pass", pblScreenShot);
            }
            catch (Exception pobjException)
            {
                clsReportResult.fnLog("ScrollToFailed", "Failed Scroll to element: " + pstrField, "Fail", true, pblHardStop, pstrHardStopMsg);
                fnExceptionHandling(pobjException);
            }
        }
    }
}
