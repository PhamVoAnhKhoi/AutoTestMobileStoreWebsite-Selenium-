using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using System;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;
namespace UnitTestProject1
{
    [TestClass]
    public class Product
    {
        Excel.Application dataApp;
        Excel.Workbook dataWorkbook;
        Excel.Worksheet dataWorksheet;
        Excel.Range xlRange;


        IWebDriver driver;

        [TestInitialize]
        public void Test1OpenWeb()
        {
            driver = new ChromeDriver();
            driver.Navigate().GoToUrl("https://localhost:44366");
            dataApp = new Excel.Application();
            dataWorkbook = dataApp.Workbooks.Open(@"D:\\Learn\\DBCLPM\\DoAn_DBCLPM\\DoAn_DBCLPM\\Final_TestCase - Copy.xlsx");
            dataWorksheet = (Excel.Worksheet)dataWorkbook.Sheets[1];
            xlRange = dataWorksheet.UsedRange;

            Thread.Sleep(1000);
        }

        [TestMethod]
        public void ID_55_DGSP1()
        {
            IWebElement loginIcon = driver.FindElement(By.CssSelector(".bi.bi-person"));
            loginIcon.Click();
            //DN đúng
            IWebElement emailInput = driver.FindElement(By.CssSelector("input[placeholder='Tài khoản']"));
            emailInput.SendKeys(xlRange.Cells[6][200].value.ToString);
            //MK đúng
            IWebElement passwordInput = driver.FindElement(By.CssSelector("input[placeholder='Mật khẩu']"));
            passwordInput.SendKeys(xlRange.Cells[6][201].value.ToString);

            IWebElement loginButton = driver.FindElement(By.CssSelector("input[value='Đăng nhập']"));
            loginButton.Click();

            IWebElement PhoneButton = driver.FindElement(By.CssSelector(".nav-link.active"));
            PhoneButton.Click();

            IWebElement RandomProduct = driver.FindElement(By.ClassName("card-img-top"));
            RandomProduct.Click();

            IWebElement reviewTextArea = driver.FindElement(By.ClassName("comment-text-area"));
            reviewTextArea.SendKeys(xlRange.Cells[6][202].value.ToString);

            IWebElement submitReviewButton = driver.FindElement(By.CssSelector("button[class='btn submit-comment']"));
            submitReviewButton.Click();
            string Suc = " success";
            string Failed = " failed";
            if (driver.Url != "https://localhost:44366")
            {
                xlRange[11][200].value = Suc;

            }
            else
            {
                xlRange[11][200].value = Failed;

            }
            Thread.Sleep(1500);

        }

        [TestMethod]
        public void ID_56_DGSP2()
        {
            IWebElement loginIcon = driver.FindElement(By.CssSelector(".bi.bi-person"));
            loginIcon.Click();
            //DN đúng
            IWebElement emailInput = driver.FindElement(By.CssSelector("input[placeholder='Tài khoản']"));
            emailInput.SendKeys(xlRange.Cells[6][204].value.ToString);
            //MK đúng
            IWebElement passwordInput = driver.FindElement(By.CssSelector("input[placeholder='Mật khẩu']"));
            passwordInput.SendKeys(xlRange.Cells[6][205].value.ToString);

            IWebElement loginButton = driver.FindElement(By.CssSelector("input[value='Đăng nhập']"));
            loginButton.Click();


            IWebElement PhoneButton = driver.FindElement(By.CssSelector(".nav-link.active"));
            PhoneButton.Click();

            IWebElement RandomProduct = driver.FindElement(By.ClassName("card-img-top"));
            RandomProduct.Click();

            IWebElement submitStarButton = driver.FindElement(By.CssSelector("span[id='5-star'] i[class='bi bi-star-fill']"));
            submitStarButton.Click();

            IWebElement reviewTextArea = driver.FindElement(By.ClassName("comment-text-area"));
            reviewTextArea.SendKeys(xlRange.Cells[6][206].value.ToString);

            IWebElement submitReviewButton = driver.FindElement(By.CssSelector("button[class='btn submit-comment']"));
            submitReviewButton.Click();
            string Suc = " success";
            string Failed = " failed";
            if (driver.Url != "https://localhost:44366")
            {
                xlRange[11][204].value = Suc;

            }
            else
            {
                xlRange[11][204].value = Failed;

            }
            Thread.Sleep(1500);

        }


        [TestMethod]
        public void ID_45_ThemSP1()
        {

            IWebElement loginIcon = driver.FindElement(By.CssSelector(".bi.bi-person"));
            loginIcon.Click();
            //DN đúng
            IWebElement emailInput = driver.FindElement(By.CssSelector("input[placeholder='Tài khoản']"));
            emailInput.SendKeys(xlRange.Cells[6][144].value.ToString());
            //MK đúng
            IWebElement passwordInput = driver.FindElement(By.CssSelector("input[placeholder='Mật khẩu']"));
            passwordInput.SendKeys(xlRange.Cells[6][145].value.ToString());


            IWebElement loginButton = driver.FindElement(By.CssSelector("input[value='Đăng nhập']"));
            loginButton.Click();

            IWebElement submitQLSPButton = driver.FindElement(By.CssSelector("body > div:nth-child(2) > div:nth-child(2) > ul:nth-child(1) > li:nth-child(4) > a:nth-child(1) > span:nth-child(2)"));
            submitQLSPButton.Click();

            IWebElement submitTSPButton = driver.FindElement(By.CssSelector("a[href='/Admin/AdminProducts/Create']"));
            submitTSPButton.Click();


            IWebElement vietTenArea = driver.FindElement(By.CssSelector("input[placeholder='Tên sản phẩm']"));
            vietTenArea.SendKeys(xlRange.Cells[6][146].value.ToString());

            IWebElement vietSoLuongArea = driver.FindElement(By.CssSelector("input[value='1']"));
            vietSoLuongArea.Clear();



            IWebElement GIAArea = driver.FindElement(By.CssSelector("#fname"));


            IWebElement THArea = driver.FindElement(By.CssSelector("#CategoryID"));
            THArea.Click();

            IWebElement CHIPArea = driver.FindElement(By.CssSelector("#Chipset"));

            IWebElement RAMArea = driver.FindElement(By.CssSelector("#Ram"));

            IWebElement MMArea = driver.FindElement(By.CssSelector("#Memory"));

            IWebElement SCArea = driver.FindElement(By.CssSelector("#ScreenSize"));

            IWebElement OSArea = driver.FindElement(By.CssSelector("#OS"));

            IWebElement CAMERAArea = driver.FindElement(By.CssSelector("#Camera"));

            IWebElement PINArea = driver.FindElement(By.CssSelector("#Pin"));

            IWebElement ResolutionArea = driver.FindElement(By.CssSelector("#Resolution"));



            string Suc = " success";
            string Failed = " failed";



            vietSoLuongArea.SendKeys(xlRange.Cells[6][156].value.ToString());
            GIAArea.SendKeys(xlRange.Cells[6][147].value.ToString());
            OSArea.SendKeys(xlRange.Cells[6][152].value.ToString());
            CHIPArea.SendKeys(xlRange.Cells[6][148].value.ToString());
            RAMArea.SendKeys(xlRange.Cells[6][149].value.ToString());
            CAMERAArea.SendKeys(xlRange.Cells[6][153].value.ToString());
            PINArea.SendKeys(xlRange.Cells[6][154].value.ToString());
            MMArea.SendKeys(xlRange.Cells[6][150].value.ToString());
            SCArea.SendKeys(xlRange.Cells[6][155].value.ToString());
            ResolutionArea.SendKeys(xlRange.Cells[6][158].value.ToString());
            IWebElement SubmitArea = driver.FindElement(By.CssSelector("button[type='submit']"));
            SubmitArea.Click();

            if (driver.Url != "https://localhost:44366")
            {
                xlRange[11][144].value = Suc;

            }
            else
            {
                xlRange[11][144].value = Failed;

            }
            Thread.Sleep(1500);


        }
        [TestMethod]
        public void ID_46_ThemSP2()
        {

            IWebElement loginIcon = driver.FindElement(By.CssSelector(".bi.bi-person"));
            loginIcon.Click();
            //DN đúng
            IWebElement emailInput = driver.FindElement(By.Name("UserEmail"));
            emailInput.SendKeys(xlRange.Cells[6][144].value.ToString());
            //MK đúng
            IWebElement passwordInput = driver.FindElement(By.Name("UserPassword"));
            passwordInput.SendKeys(xlRange.Cells[6][145].value.ToString());


            IWebElement loginButton = driver.FindElement(By.CssSelector("input[value='Đăng nhập']"));
            loginButton.Click();

            IWebElement submitQLSPButton = driver.FindElement(By.CssSelector("body > div:nth-child(2) > div:nth-child(2) > ul:nth-child(1) > li:nth-child(4) > a:nth-child(1) > span:nth-child(2)"));
            submitQLSPButton.Click();

            IWebElement submitTSPButton = driver.FindElement(By.CssSelector("a[href='/Admin/AdminProducts/Create']"));
            submitTSPButton.Click();

            IWebElement quaylaiTSPButton = driver.FindElement(By.CssSelector("div[class='main-admin'] div a"));
            quaylaiTSPButton.Click();

        }

        [TestMethod]
        public void ID_47_XoaSP()
        {
            IWebElement loginIcon = driver.FindElement(By.CssSelector(".bi.bi-person"));
            loginIcon.Click();
            //DN đúng
            IWebElement emailInput = driver.FindElement(By.Name("UserEmail"));
            emailInput.SendKeys(xlRange.Cells[6][172].value.ToString());
            //MK đúng
            IWebElement passwordInput = driver.FindElement(By.Name("UserPassword"));
            passwordInput.SendKeys(xlRange.Cells[6][173].value.ToString());


            IWebElement loginButton = driver.FindElement(By.CssSelector("input[value='Đăng nhập']"));
            loginButton.Click();

            IWebElement submitQLSPButton = driver.FindElement(By.CssSelector("body > div:nth-child(2) > div:nth-child(2) > ul:nth-child(1) > li:nth-child(4) > a:nth-child(1) > span:nth-child(2)"));
            submitQLSPButton.Click();

            IWebElement submitXoaSPButton = driver.FindElement(By.CssSelector("a[href='/Admin/AdminProducts/Delete/1032']"));
            submitXoaSPButton.Click();

            IWebElement submitXoaSPButton2 = driver.FindElement(By.CssSelector("input[value='Xóa']"));
            submitXoaSPButton2.Click();


            string Suc = " success";
            string Failed = " failed";
            if (driver.Url != "https://localhost:44366")
            {
                xlRange[11][172].value = Suc;

            }
            else
            {
                xlRange[11][172].value = Failed;

            }
            Thread.Sleep(1500);
        }
        public void ID_47_XoaSP2()
        {
            IWebElement loginIcon = driver.FindElement(By.CssSelector(".bi.bi-person"));
            loginIcon.Click();
            //DN đúng
            IWebElement emailInput = driver.FindElement(By.Name("UserEmail"));
            emailInput.SendKeys(xlRange.Cells[6][172].value.ToString());
            //MK đúng
            IWebElement passwordInput = driver.FindElement(By.Name("UserPassword"));
            passwordInput.SendKeys(xlRange.Cells[6][172].value.ToString());


            IWebElement loginButton = driver.FindElement(By.CssSelector("input[value='Đăng nhập']"));
            loginButton.Click();

            IWebElement submitQLSPButton = driver.FindElement(By.CssSelector("body > div:nth-child(2) > div:nth-child(2) > ul:nth-child(1) > li:nth-child(4) > a:nth-child(1) > span:nth-child(2)"));
            submitQLSPButton.Click();

            IWebElement submitXoaSPButton = driver.FindElement(By.CssSelector("a[href='/Admin/AdminProducts/Delete/1030']"));
            submitXoaSPButton.Click();

            IWebElement tbXoaSPButton = driver.FindElement(By.CssSelector("div[class='form-actions no-color'] a"));
            tbXoaSPButton.Click();


        }

        [TestMethod]
        public void ID_48_CSSP1()
        {
            IWebElement loginIcon = driver.FindElement(By.CssSelector(".bi.bi-person"));
            loginIcon.Click();
            //DN đúng
            IWebElement emailInput = driver.FindElement(By.Name("UserEmail"));
            emailInput.SendKeys(xlRange.Cells[6][177].value.ToString());
            //MK đúng
            IWebElement passwordInput = driver.FindElement(By.Name("UserPassword"));
            passwordInput.SendKeys(xlRange.Cells[6][178].value.ToString());
            IWebElement loginButton = driver.FindElement(By.CssSelector("input[value='Đăng nhập']"));
            loginButton.Click();
            IWebElement submitQLSPButton = driver.FindElement(By.CssSelector("body > div:nth-child(2) > div:nth-child(2) > ul:nth-child(1) > li:nth-child(4) > a:nth-child(1) > span:nth-child(2)"));
            submitQLSPButton.Click();

            IWebElement submitCSButton = driver.FindElement(By.CssSelector("a[href='/Admin/AdminProducts/Edit/10']"));
            submitCSButton.Click();

            IWebElement reviewTenArea = driver.FindElement(By.CssSelector("input[value='Samsung Galaxy A53 128GB ']"));
            reviewTenArea.Clear();

            reviewTenArea.SendKeys(xlRange.Cells[6][179].value.ToString());

            IWebElement submitLuuButton = driver.FindElement(By.CssSelector("button[type='submit']"));
            submitLuuButton.Click();
            string Suc = " success";
            string Failed = " failed";
            if (driver.Url != "https://localhost:44366")
            {
                xlRange[11][177].value = Suc;

            }
            else
            {
                xlRange[11][177].value = Failed;

            }
            Thread.Sleep(1500);
        }

        [TestMethod]
        public void ID_49_CSSP2()
        {
            IWebElement loginIcon = driver.FindElement(By.CssSelector(".bi.bi-person"));
            loginIcon.Click();
            //DN đúng
            IWebElement emailInput = driver.FindElement(By.Name("UserEmail"));
            emailInput.SendKeys(xlRange.Cells[6][183].value.ToString());
            //MK đúng
            IWebElement passwordInput = driver.FindElement(By.Name("UserPassword"));
            passwordInput.SendKeys(xlRange.Cells[6][184].value.ToString());

            IWebElement loginButton = driver.FindElement(By.CssSelector("input[value='Đăng nhập']"));
            loginButton.Click();

            IWebElement submitQLSPButton = driver.FindElement(By.CssSelector("body > div:nth-child(2) > div:nth-child(2) > ul:nth-child(1) > li:nth-child(4) > a:nth-child(1) > span:nth-child(2)"));
            submitQLSPButton.Click();

            IWebElement submitCSButton = driver.FindElement(By.CssSelector("a[href='/Admin/AdminProducts/Edit/1']"));
            submitCSButton.Click();

            IWebElement reviewGiaArea = driver.FindElement(By.CssSelector("input[value='16']"));
            reviewGiaArea.Clear();

            reviewGiaArea.SendKeys(xlRange.Cells[6][185].value.ToString());

            IWebElement submitLuuButton = driver.FindElement(By.CssSelector("button[type='submit']"));
            submitLuuButton.Click();
            string Suc = " success";
            string Failed = " failed";
            if (driver.Url != "https://localhost:44366")
            {
                xlRange[11][183].value = Suc;

            }
            else
            {
                xlRange[11][183].value = Failed;

            }
            Thread.Sleep(1500);

        }

        [TestMethod]
        public void ID_49_CSSP3()
        {
            IWebElement loginIcon = driver.FindElement(By.CssSelector(".bi.bi-person"));
            loginIcon.Click();
            //DN đúng
            IWebElement emailInput = driver.FindElement(By.Name("UserEmail"));
            emailInput.SendKeys(xlRange.Cells[6][183].value.ToString());
            //MK đúng
            IWebElement passwordInput = driver.FindElement(By.Name("UserPassword"));
            passwordInput.SendKeys(xlRange.Cells[6][184].value.ToString());

            IWebElement loginButton = driver.FindElement(By.CssSelector("input[value='Đăng nhập']"));
            loginButton.Click();


            IWebElement submitQLSPButton = driver.FindElement(By.CssSelector("body > div:nth-child(2) > div:nth-child(2) > ul:nth-child(1) > li:nth-child(4) > a:nth-child(1) > span:nth-child(2)"));
            submitQLSPButton.Click();

            IWebElement submitCSButton = driver.FindElement(By.CssSelector("a[href='/Admin/AdminProducts/Edit/1']"));
            submitCSButton.Click();

            IWebElement tbCSButton = driver.FindElement(By.CssSelector("body div[class='main-content'] div[class='main-admin'] div a:nth-child(1)"));
            tbCSButton.Click();
        }
        [TestMethod]
        public void ID_50_TKSP1()
        {



            IWebElement TkBt = driver.FindElement(By.CssSelector("input[placeholder='Tìm kiếm']"));
            TkBt.SendKeys(xlRange.Cells[6][188].value.ToString());

            IWebElement submitCSButton = driver.FindElement(By.CssSelector(".bi.bi-search"));
            submitCSButton.Click();
            string Suc = " success";
            string Failed = " failed";
            if (driver.Url != "https://localhost:44366")
            {
                xlRange[11][188].value = Suc;

            }
            else
            {
                xlRange[11][188].value = Failed;

            }
            Thread.Sleep(1500);
        }

        [TestMethod]
        public void ID_51_TKSP2()
        {


            IWebElement TkBt = driver.FindElement(By.CssSelector("input[placeholder='Tìm kiếm']"));
            TkBt.SendKeys(xlRange.Cells[6][194].value.ToString());

            IWebElement submitCSButton = driver.FindElement(By.CssSelector(".bi.bi-search"));
            submitCSButton.Click();
            string Suc = " success";
            string Failed = " failed";
            if (driver.Url != "https://localhost:44366")
            {
                xlRange[11][194].value = Suc;

            }
            else
            {
                xlRange[11][194].value = Failed;

            }
            Thread.Sleep(1500);
        }
        [TestMethod]
        public void ID_52_PLSP1()
        {


            IWebElement MENUCSButton = driver.FindElement(By.CssSelector(".bi.bi-list"));
            MENUCSButton.Click();
            IWebElement OppoButton = driver.FindElement(By.CssSelector("div[class='narbar'] a:nth-child(2)"));
            OppoButton.Click();
            Thread.Sleep(1000);
            string Suc = " success";
            string Failed = " failed";
            if (driver.Url != "https://localhost:44366")
            {
                xlRange[11][196].value = Suc;

            }
            else
            {
                xlRange[11][196].value = Failed;

            }
            Thread.Sleep(1500);
        }

        [TestMethod]
        public void ID_53_PLSP2()
        {
            IWebElement MENUdtCSButton = driver.FindElement(By.CssSelector(".nav-link.active"));
            MENUdtCSButton.Click();
            IWebElement T4RButton = driver.FindElement(By.CssSelector("body > div:nth-child(3) > div:nth-child(1) > div:nth-child(3) > div:nth-child(1) > div:nth-child(1) > button:nth-child(2)"));
            T4RButton.Click();
            Thread.Sleep(1000);
            string Suc = " success";
            string Failed = " failed";
            if (driver.Url != "https://localhost:44366")
            {
                xlRange[11][197].value = Suc;

            }
            else
            {
                xlRange[11][197].value = Failed;

            }
            Thread.Sleep(1500);
        }
        [TestMethod]
        public void ID_54_PLSP3()
        {
            IWebElement MENUIOSCSButton = driver.FindElement(By.CssSelector("div[class='narbar'] li:nth-child(2) a:nth-child(1)"));
            MENUIOSCSButton.Click();
            IWebElement T8_11RButton = driver.FindElement(By.CssSelector("body > div:nth-child(3) > div:nth-child(1) > div:nth-child(3) > div:nth-child(1) > div:nth-child(1) > button:nth-child(4)"));
            T8_11RButton.Click();
            Thread.Sleep(1000);
            string Suc = " success";
            string Failed = " failed";
            if (driver.Url != "https://localhost:44366")
            {
                xlRange[11][198].value = Suc;

            }
            else
            {
                xlRange[11][198].value = Failed;

            }
            Thread.Sleep(1500);
        }

        [TestMethod]
        public void ID_57_TTH1()
        {
            IWebElement loginIcon = driver.FindElement(By.CssSelector(".bi.bi-person"));
            loginIcon.Click();
            //DN đúng
            IWebElement emailInput = driver.FindElement(By.Name("UserEmail"));
            emailInput.SendKeys(xlRange.Cells[6][210].value.ToString());
            //MK đúng
            IWebElement passwordInput = driver.FindElement(By.Name("UserPassword"));
            passwordInput.SendKeys(xlRange.Cells[6][211].value.ToString());

            IWebElement loginButton = driver.FindElement(By.CssSelector("input[value='Đăng nhập']"));
            loginButton.Click();
            IWebElement THButton = driver.FindElement(By.CssSelector("body > div:nth-child(2) > div:nth-child(2) > ul:nth-child(1) > li:nth-child(6) > a:nth-child(1) > span:nth-child(2)"));
            THButton.Click();
            IWebElement ThemTHButton = driver.FindElement(By.CssSelector("a[href='/Admin/AdminCategories/Create']"));
            ThemTHButton.Click();
            IWebElement vietTenArea = driver.FindElement(By.CssSelector("#CategoryName"));
            vietTenArea.SendKeys(xlRange.Cells[6][212].value.ToString());
            IWebElement smButton = driver.FindElement(By.CssSelector("button[type='submit']"));
            smButton.Click();
            string Suc = " success";
            string Failed = " failed";
            if (driver.Url != "https://localhost:44366")
            {
                xlRange[11][210].value = Suc;

            }
            else
            {
                xlRange[11][210].value = Failed;

            }
            Thread.Sleep(1500);
        }
        [TestMethod]
        public void ID_57_TTH2()
        {
            IWebElement loginIcon = driver.FindElement(By.CssSelector(".bi.bi-person"));
            loginIcon.Click();
            //DN đúng
            IWebElement emailInput = driver.FindElement(By.Name("UserEmail"));
            emailInput.SendKeys(xlRange.Cells[6][210].value.ToString());
            //MK đúng
            IWebElement passwordInput = driver.FindElement(By.Name("UserPassword"));
            passwordInput.SendKeys(xlRange.Cells[6][211].value.ToString());

            IWebElement loginButton = driver.FindElement(By.CssSelector("input[value='Đăng nhập']"));
            loginButton.Click();
            IWebElement THButton = driver.FindElement(By.CssSelector("body > div:nth-child(2) > div:nth-child(2) > ul:nth-child(1) > li:nth-child(6) > a:nth-child(1) > span:nth-child(2)"));
            THButton.Click();
            IWebElement ThemTHButton = driver.FindElement(By.CssSelector("a[href='/Admin/AdminCategories/Create']"));
            ThemTHButton.Click();
            IWebElement BackTHButton = driver.FindElement(By.CssSelector("div[class='main-admin'] a"));
            BackTHButton.Click();
        }
        [TestMethod]
        public void ID_58_XoaTH1()
        {
            IWebElement loginIcon = driver.FindElement(By.CssSelector(".bi.bi-person"));
            loginIcon.Click();
            //DN đúng
            IWebElement emailInput = driver.FindElement(By.Name("UserEmail"));
            emailInput.SendKeys(xlRange.Cells[6][215].value.ToString());
            //MK đúng
            IWebElement passwordInput = driver.FindElement(By.Name("UserPassword"));
            passwordInput.SendKeys(xlRange.Cells[6][216].value.ToString());

            IWebElement loginButton = driver.FindElement(By.CssSelector("input[value='Đăng nhập']"));
            loginButton.Click();
            IWebElement THButton = driver.FindElement(By.CssSelector("body > div:nth-child(2) > div:nth-child(2) > ul:nth-child(1) > li:nth-child(6) > a:nth-child(1) > span:nth-child(2)"));
            THButton.Click();
            IWebElement DTHButton = driver.FindElement(By.CssSelector("a[href='/Admin/AdminCategories/Delete/12']"));
            DTHButton.Click();
            IWebElement DxTHButton = driver.FindElement(By.CssSelector("input[value='Delete']"));
            DxTHButton.Click();
            string Suc = " success";
            string Failed = " failed";
            if (driver.Url != "https://localhost:44366")
            {
                xlRange[11][215].value = Suc;

            }
            else
            {
                xlRange[11][216].value = Failed;

            }
            Thread.Sleep(1500);
        }
        [TestMethod]
        public void ID_58_XoaTH2()
        {
            IWebElement loginIcon = driver.FindElement(By.CssSelector(".bi.bi-person"));
            loginIcon.Click();
            //DN đúng
            IWebElement emailInput = driver.FindElement(By.Name("UserEmail"));
            emailInput.SendKeys(xlRange.Cells[6][215].value.ToString());
            //MK đúng
            IWebElement passwordInput = driver.FindElement(By.Name("UserPassword"));
            passwordInput.SendKeys(xlRange.Cells[6][216].value.ToString());

            IWebElement loginButton = driver.FindElement(By.CssSelector("input[value='Đăng nhập']"));
            loginButton.Click();
            IWebElement THButton = driver.FindElement(By.CssSelector("body > div:nth-child(2) > div:nth-child(2) > ul:nth-child(1) > li:nth-child(6) > a:nth-child(1) > span:nth-child(2)"));
            THButton.Click();
            IWebElement DTHButton = driver.FindElement(By.CssSelector("a[href='/Admin/AdminCategories/Delete/12']"));
            DTHButton.Click();
            IWebElement BTHButton = driver.FindElement(By.CssSelector("div[class='form-actions no-color'] a"));
            BTHButton.Click();
        }
        [TestMethod]
        public void ID_59_CSTH1()
        {
            IWebElement loginIcon = driver.FindElement(By.CssSelector(".bi.bi-person"));
            loginIcon.Click();
            //DN đúng
            IWebElement emailInput = driver.FindElement(By.Name("UserEmail"));
            emailInput.SendKeys(xlRange.Cells[6][220].value.ToString());
            //MK đúng
            IWebElement passwordInput = driver.FindElement(By.Name("UserPassword"));
            passwordInput.SendKeys(xlRange.Cells[6][221].value.ToString());

            IWebElement loginButton = driver.FindElement(By.CssSelector("input[value='Đăng nhập']"));
            loginButton.Click();
            IWebElement THButton = driver.FindElement(By.CssSelector("body > div:nth-child(2) > div:nth-child(2) > ul:nth-child(1) > li:nth-child(6) > a:nth-child(1) > span:nth-child(2)"));
            THButton.Click();
            IWebElement CSButton = driver.FindElement(By.CssSelector("a[href='/Admin/AdminCategories/Edit/3']"));
            CSButton.Click();
            IWebElement nameth = driver.FindElement(By.CssSelector("input[value='ấdasd']"));
            nameth.Clear();
            nameth.SendKeys(xlRange.Cells[6][222].value.ToString());
            IWebElement CSxButton = driver.FindElement(By.CssSelector("button[type='submit']"));
            CSxButton.Click();
            string Suc = " success";
            string Failed = " failed";
            if (driver.Url != "https://localhost:44366")
            {
                xlRange[11][220].value = Suc;

            }
            else
            {
                xlRange[11][220].value = Failed;

            }
            Thread.Sleep(1500);
        }
        [TestMethod]
        public void ID_59_CSTH2()
        {
            IWebElement loginIcon = driver.FindElement(By.CssSelector(".bi.bi-person"));
            loginIcon.Click();
            //DN đúng
            IWebElement emailInput = driver.FindElement(By.Name("UserEmail"));
            emailInput.SendKeys(xlRange.Cells[6][220].value.ToString());
            //MK đúng
            IWebElement passwordInput = driver.FindElement(By.Name("UserPassword"));
            passwordInput.SendKeys(xlRange.Cells[6][221].value.ToString());

            IWebElement loginButton = driver.FindElement(By.CssSelector("input[value='Đăng nhập']"));
            loginButton.Click();
            IWebElement THButton = driver.FindElement(By.CssSelector("body > div:nth-child(2) > div:nth-child(2) > ul:nth-child(1) > li:nth-child(6) > a:nth-child(1) > span:nth-child(2)"));
            THButton.Click();
            IWebElement CSButton = driver.FindElement(By.CssSelector("a[href='/Admin/AdminCategories/Edit/3']"));
            CSButton.Click();
            IWebElement bCSButton = driver.FindElement(By.CssSelector("body div[class='main-content'] div[class='main-admin'] div a:nth-child(1)"));
            bCSButton.Click();

        }
        [TestCleanup]
        public void Test2CloseWeb()
        {
            dataWorkbook.Save();
            dataWorkbook.Close();
            dataApp.Quit();
            // Close the browser
            driver.Quit();
        }
    }
}
