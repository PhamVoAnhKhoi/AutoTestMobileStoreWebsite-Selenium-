using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Threading;
using static System.Net.WebRequestMethods;
using Excel = Microsoft.Office.Interop.Excel;

namespace UnitTestProject1
{
    [TestClass]
    public class LoginTests
    {
        private IWebDriver driver;
        private IWebElement edtAccount;
        private IWebElement edtPassword;
        private IWebElement navBtnLogin;
        private IWebElement btnLogin;
        private IWebElement hrefSignUp;
        private IWebElement btnLogOut;
        bool isDisplayed;

        Excel.Application dataApp;
        Excel.Workbook dataWorkbook;
        Excel.Worksheet dataWorksheet;
        Excel.Range xlRange;

        [TestInitialize]
        public void OpenWebsite()
        {
            driver = new ChromeDriver();
            driver.Url = "https://localhost:44366/Users/Login";
            driver.Navigate();
            dataApp = new Excel.Application();
            dataWorkbook = dataApp.Workbooks.Open(@"D:\\Learn\\DBCLPM\\DoAn_DBCLPM\\DoAn_DBCLPM\\Final_TestCase - Copy.xlsx");
            dataWorksheet = dataWorkbook.Sheets[1];
            xlRange = dataWorksheet.UsedRange;

        }

        //Test biểu tượng đăng nhập trên thanh nav_bar
        [TestMethod]
        public void ID_22_IconTKNavBar()
        {
            //Test navBtnLogin
            navBtnLogin = driver.FindElement(By.Id("login"));
            // Kiểm tra xem nút có hiển thị trên trang không
            isDisplayed = navBtnLogin.Displayed;
            Console.WriteLine("navBtnLogin hiển thị trên trang: " + isDisplayed);
            // Nhấp vào nút
            navBtnLogin.Click();
            if (driver.Url == "https://localhost:44366/Users/Login")
            {
                Console.WriteLine("Chuyển hướng đến trang đúng.");
            }
            else
            {
                Console.WriteLine("Chuyển hướng không thành công hoặc đến trang không mong muốn.");
            }
            Thread.Sleep(5000);
        }
        //Test chuyển hướng đến trang SignUp
        [TestMethod]
        public void ID_23_ChuyenHuongTrangDangKy() 
        {
            //Test hrefSignUp
            hrefSignUp = driver.FindElement(By.XPath("/html/body/div[3]/div/div/div[2]/form/div[3]/a"));
            // Kiểm tra xem nút có hiển thị trên trang không
            isDisplayed = hrefSignUp.Displayed;
            Console.WriteLine("hrefSignUp hiển thị trên trang: " + isDisplayed);
            // Nhấp vào nút
            hrefSignUp.Click();
            if (driver.Url == "https://localhost:44366/Users/SignUp")
            {
                Console.WriteLine("Chuyển hướng đến trang đúng.");
            }
            else
            {
                Console.WriteLine("Chuyển hướng không thành công hoặc đến trang không mong muốn.");
            }
            Thread.Sleep(7000);
        }
        //Chuẩn bị các khai báo Login
        public void prepareLogin()
        {
            //UserAccount
            edtAccount = driver.FindElement(By.Name("UserEmail"));

            //UserPassword
            edtPassword = driver.FindElement(By.Name("UserPassword"));

            //Test btnLogin
            btnLogin = driver.FindElement(By.XPath("/html/body/div[3]/div/div/div[2]/form/input"));
        }

         //Test GUI đăng nhập
        [TestMethod]
        public void ID_24_GUIDangNhap()
        {
            prepareLogin();

            //UserAccount
            Assert.AreEqual(true, edtAccount.Enabled);

            //UserPassword
            Assert.AreEqual(true, edtPassword.Enabled);
            
            // Kiểm tra xem nút có hiển thị trên trang không
            isDisplayed = btnLogin.Displayed;
            if (true)
            {

            }
            Console.WriteLine("btnLogin hiển thị trên trang: " + isDisplayed);

            Thread.Sleep(3000);
        }

        // Đăng nhập thành công
        [TestMethod]
        public void ID_18_DangNhapThanhCong()
        {
            prepareLogin();   
            // Lấy data từ Excel để truyền vào 
            edtAccount.SendKeys(xlRange.Cells[6][52].value.ToString());
            edtPassword.SendKeys(xlRange.Cells[6][53].value.ToString());

            // Nhấp vào nút
            btnLogin.Click();

            // Kiểm tra đăng nhập thành công bằng cách kiểm tra URL
            string sucOrFail = driver.Url != "https://localhost:44366/Users/Login" ? "Pass" : "Fail";

            dataWorksheet.Cells[12][52] = sucOrFail;

            Thread.Sleep(5000);
        }

        //Đăng nhập không thành công sai tài khoản 
        [TestMethod]
        public void ID_19_DangNhapTkKhongTonTai()
        {
            prepareLogin();

            // Lấy data từ Excel để truyền vào 
            edtAccount.SendKeys(xlRange.Cells[6][54].value.ToString());
            edtPassword.SendKeys(xlRange.Cells[6][55].value.ToString());

            // Nhấp vào nút
            btnLogin.Click();

            IWebElement errorMessage = null;
            try
            {
                errorMessage = driver.FindElement(By.XPath("/html/body/div[3]/div/div/div[2]/form/small"));
            }
            catch (NoSuchElementException)
            {
                // Không có thông báo lỗi
            }


            // Kiểm tra đăng nhập thành công bằng cách kiểm tra URL
            string sucOrFail = errorMessage != null ? "Pass" : "Fail";

            dataWorksheet.Cells[12][54] = sucOrFail;

            Thread.Sleep(5000);
        }

        //Đăng nhập không thành công sai tài khoản 
        [TestMethod]
        public void ID_20_DangNhapSaiTK()
        {
            prepareLogin();

            // Lấy data từ Excel để truyền vào 
            edtAccount.SendKeys(xlRange.Cells[6][56].value.ToString());
            edtPassword.SendKeys(xlRange.Cells[6][57].value.ToString());

            // Nhấp vào nút
            btnLogin.Click();

            IWebElement errorMessage = null;
            try
            {
                errorMessage = driver.FindElement(By.XPath("/html/body/div[3]/div/div/div[2]/form/small"));
            }
            catch (NoSuchElementException)
            {
                // Không có thông báo lỗi
            }


            // Kiểm tra đăng nhập thành công bằng cách kiểm tra URL
            string sucOrFail = errorMessage != null ? "Pass" : "Fail";

            dataWorksheet.Cells[12][56] = sucOrFail;

            Thread.Sleep(5000);
        }

        //Đăng nhập không thành công sai mật khẩu
        [TestMethod]
        public void ID_21_DangNhapSaiMK()
        {
            prepareLogin();

            // Lấy data từ Excel để truyền vào 
            edtAccount.SendKeys(xlRange.Cells[6][58].value.ToString());
            edtPassword.SendKeys(xlRange.Cells[6][59].value.ToString());

            // Nhấp vào nút
            btnLogin.Click();

            IWebElement errorMessage = null;
            try
            {
                errorMessage = driver.FindElement(By.XPath("/html/body/div[3]/div/div/div[2]/form/small"));
            }
            catch (NoSuchElementException)
            {
                // Không có thông báo lỗi
            }


            // Kiểm tra đăng nhập thành công bằng cách kiểm tra URL
            string sucOrFail = errorMessage != null ? "Pass" : "Fail";

            dataWorksheet.Cells[12][58] = sucOrFail;

            Thread.Sleep(5000);
        }


        //Test LogOut
        [TestMethod]
        public void ID_33_DangXuat()
        {
            string sucOrFail;

            prepareLogin();
            // Lấy data từ Excel để truyền vào 
            edtAccount.SendKeys(xlRange.Cells[6][52].value.ToString());
            edtPassword.SendKeys(xlRange.Cells[6][53].value.ToString());

            // Nhấp vào nút
            btnLogin.Click();

            IWebElement btnIconUser = driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[3]/div[3]/a"));
            btnIconUser.Click();

            btnLogOut = driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[3]/div[3]/div/a[4]"));
            btnLogOut.Click();
            if (driver.Url == "")
            {
                sucOrFail = "Pass";
                Console.WriteLine("Log out success");
                dataWorksheet.Cells[12][97] = sucOrFail;
            }
            else
            {
                sucOrFail = "Fail";
                Console.WriteLine("Log out fail");
                dataWorksheet.Cells[12][97] = sucOrFail;
            }
           
            Thread.Sleep(10000);
        }
        [TestCleanup]
        public void Teardown()
        {
            dataWorkbook.Save();
            dataWorkbook.Close();
            dataApp.Quit();
            // Close the browser
            driver.Quit();
        }
    }

    /// <summary>
    /// Test SignUp
    /// </summary>
    
}
