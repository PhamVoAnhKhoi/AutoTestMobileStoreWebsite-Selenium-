using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;

namespace UnitTestProject1
{
    
    /// <summary>
    /// Summary description for Regist
    /// </summary>
    [TestClass]
    public class Regist
    {

        Excel.Application dataApp;
        Excel.Workbook dataWorkbook;
        Excel.Worksheet dataWorksheet;
        Excel.Range xlRange;

        private IWebDriver driver;
        private IWebElement edtName, edtEmail, edtPhone, edtPassword, edtRePassword, chkBoxSignUp,
                            btnSignUp;
        [TestInitialize]
        public void OpenWebsite()
        {
            driver = new ChromeDriver();
            driver.Url = "https://localhost:44366/Users/SignUp";
            driver.Navigate();
            dataApp = new Excel.Application();
            dataWorkbook = dataApp.Workbooks.Open(@"D:\\Learn\\DBCLPM\\DoAn_DBCLPM\\DoAn_DBCLPM\\Final_TestCase.xlsx");
            dataWorksheet = dataWorkbook.Sheets[1];
            xlRange = dataWorksheet.UsedRange;
        }

        public void prepareSignUp()
        {
            edtName = driver.FindElement(By.Name("UserName"));
            edtEmail = driver.FindElement(By.Name("UserEmail"));
            edtPhone = driver.FindElement(By.Name("PhoneNumber"));
            edtPassword = driver.FindElement(By.Name("UserPassword"));
            edtRePassword = driver.FindElement(By.Name("RePassword"));
            btnSignUp = driver.FindElement(By.XPath("/html/body/div[3]/div/div/div[2]/form/div/div[6]/div[2]/button"));
            chkBoxSignUp = driver.FindElement(By.XPath("/html/body/div[3]/div/div/div[2]/form/div/div[6]/div[1]"));
        }
        [TestMethod]
        //Gui Đăng ký
        public void ID_25_GUIDangKy()
        {
            prepareSignUp();

            Assert.AreEqual(true, edtName.Enabled);
         
            Assert.AreEqual(true, edtEmail.Enabled);
           
            Assert.AreEqual(true, edtPhone.Enabled);
          
            Assert.AreEqual(true, edtPassword.Enabled);
           
            Assert.AreEqual(true, edtRePassword.Enabled);


            Thread.Sleep(5000);
        }

        //Đăng ký thành công 
        [TestMethod]
        public void ID_26_DangKyThanhCong()
        {
            prepareSignUp();

            string sucOrFail;
            // Điền thông tin vào các trường
                edtName.SendKeys(xlRange.Cells[6][65].value.ToString());
                edtEmail.SendKeys(xlRange.Cells[6][66].value.ToString());
                edtPhone.SendKeys(xlRange.Cells[6][67].value.ToString());
                edtPassword.SendKeys(xlRange.Cells[6][68].value.ToString());
                edtRePassword.SendKeys(xlRange.Cells[6][69].value.ToString());
           
            // Chọn checkbox đồng ý đăng ký
            
            chkBoxSignUp.Click();
            btnSignUp.Click();
            if (driver.Url == "https://localhost:44366/Users/Login")
            {
                sucOrFail = "Passed";
                dataWorksheet.Cells[12][65] = sucOrFail;
            }
            else
            {
                sucOrFail = "Failed";
                dataWorksheet.Cells[12][65] = sucOrFail;
            }
            Thread.Sleep(5000);
        }

        //"Đăng ký thất bại do đã có tài khoản từ trước
        [TestMethod]
        public void ID_27_DangKyTrungTK()
        {
            prepareSignUp();
            string sucOrFail;
            // Điền thông tin vào các trường
            edtName.SendKeys(xlRange.Cells[6][70].value.ToString());
            edtEmail.SendKeys(xlRange.Cells[6][71].value.ToString());
            edtPhone.SendKeys(xlRange.Cells[6][72].value.ToString());
            edtPassword.SendKeys(xlRange.Cells[6][73].value.ToString());
            edtRePassword.SendKeys(xlRange.Cells[6][74].value.ToString());
            chkBoxSignUp.Click();
            btnSignUp.Click();
            if (driver.Url != "https://localhost:44366/Users/Login")
            {
                sucOrFail = "Passed";
                dataWorksheet.Cells[12][70] = sucOrFail;
            }
            else
            {
                sucOrFail = "Failed";
                dataWorksheet.Cells[12][70] = sucOrFail;
            }
            Thread.Sleep(5000);
        }

        //Đăng ký thất bại do sai yêu cầu email
        [TestMethod]
        public void ID_28_DangKySaiEmail()
        {
            prepareSignUp();
            string sucOrFail;

            // Điền thông tin vào các trường
            edtName.SendKeys(xlRange.Cells[6][75].value.ToString());
            edtEmail.SendKeys(xlRange.Cells[6][76].value.ToString());
            edtPhone.SendKeys(xlRange.Cells[6][77].value.ToString());
            edtPassword.SendKeys(xlRange.Cells[6][78].value.ToString());
            edtRePassword.SendKeys(xlRange.Cells[6][79].value.ToString());
            chkBoxSignUp.Click();
            btnSignUp.Click();
            if (driver.Url != "https://localhost:44366/Users/Login")
            {
                sucOrFail = "Passed";
                dataWorksheet.Cells[12][75] = sucOrFail;
            }
            else
            {
                sucOrFail = "Failed";
                dataWorksheet.Cells[12][75] = sucOrFail;
            }
            Thread.Sleep(5000);
        }

        //Đăng ký thất bại do sai yêu cầu SDT
        [TestMethod]
        public void ID_29_DangKySaiSDT()
        {
            prepareSignUp();
            string sucOrFail;

            // Điền thông tin vào các trường
            edtName.SendKeys(xlRange.Cells[6][80].value.ToString());
            edtEmail.SendKeys(xlRange.Cells[6][81].value.ToString());
            edtPhone.SendKeys(xlRange.Cells[6][82].value.ToString());
            edtPassword.SendKeys(xlRange.Cells[6][83].value.ToString());
            edtRePassword.SendKeys(xlRange.Cells[6][84].value.ToString());
            chkBoxSignUp.Click();
            btnSignUp.Click();
            if (driver.Url != "https://localhost:44366/Users/Login")
            {
                sucOrFail = "Passed";
                dataWorksheet.Cells[12][80] = sucOrFail;
            }
            else
            {
                sucOrFail = "Failed";
                dataWorksheet.Cells[12][80] = sucOrFail;
            }
            Thread.Sleep(5000);
        }

        //Đăng ký thất bại do "Nhập lại mật khẩu" không trùng với "Nhập mật khẩu"
        [TestMethod]
        public void ID_30_MkKhongGiongXacNhan()
        {
            prepareSignUp();
            string sucOrFail;

            // Điền thông tin vào các trường
            edtName.SendKeys(xlRange.Cells[6][85].value.ToString());
            edtEmail.SendKeys(xlRange.Cells[6][86].value.ToString());
            edtPhone.SendKeys(xlRange.Cells[6][87].value.ToString());
            edtPassword.SendKeys(xlRange.Cells[6][88].value.ToString());
            edtRePassword.SendKeys(xlRange.Cells[6][89].value.ToString());
            chkBoxSignUp.Click();
            btnSignUp.Click();
            if (driver.Url != "https://localhost:44366/Users/Login")
            {
                sucOrFail = "Passed";
                dataWorksheet.Cells[12][85] = sucOrFail;
            }
            else
            {
                sucOrFail = "Failed";
                dataWorksheet.Cells[12][85] = sucOrFail;
            }
            Thread.Sleep(8000);
        }
        //Đăng ký nhưng chưa click vào CheckBox
        [TestMethod]
        public void ID_31_CheckBoxChuaClick()
        {
            prepareSignUp();
            string sucOrFail;

            // Điền thông tin vào các trường
            edtName.SendKeys(xlRange.Cells[6][90].value.ToString());
            edtEmail.SendKeys(xlRange.Cells[6][91].value.ToString());
            edtPhone.SendKeys(xlRange.Cells[6][92].value.ToString());
            edtPassword.SendKeys(xlRange.Cells[6][93].value.ToString());
            edtRePassword.SendKeys(xlRange.Cells[6][94].value.ToString());
            btnSignUp.Click();
            if (driver.Url != "https://localhost:44366/Users/Login")
            {
                sucOrFail = "Passed";
                dataWorksheet.Cells[12][90] = sucOrFail;
                Console.WriteLine("Passed");
            }
            else
            {
                sucOrFail = "Failed";
                dataWorksheet.Cells[12][90] = sucOrFail;
                Console.WriteLine("Failed");
            }
            Thread.Sleep(5000);
        }

        //Để trống thông tin đăng ký 

        [TestMethod]
        public void ID_32_DeTrongThongTin()
        {
            prepareSignUp();
            string sucOrFail;
            btnSignUp.Click();
            if (driver.Url != "https://localhost:44366/Users/Login")
            {
                sucOrFail = "Passed";
                dataWorksheet.Cells[12][95] = sucOrFail;
            }
            else
            {
                sucOrFail = "Failed";
                dataWorksheet.Cells[12][95] = sucOrFail;
            }
            Thread.Sleep(5000);
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
}
