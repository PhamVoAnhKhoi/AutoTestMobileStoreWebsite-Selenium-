using Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Security.Principal;
using System.Threading;

using Excel = Microsoft.Office.Interop.Excel;

namespace UnitTestProject1
{
    [TestClass]
    public class QLUser
    {
        IWebDriver driver;

        IWebElement btnNavRegist;
        IWebElement btnMenuSideBar, menuSideBar;
        IWebElement sideBarQLUsers;
        IWebElement btnEditAccount, btnDetailsAccount;
        IWebElement UserName, UserEmail, PhoneNumber, UserPassword, AvaterImage;
        IWebElement edtName, edtEmail, edtPhone, edtPassword, edtRePassword, chkBoxSignUp,
                            btnSignUp;
        IWebElement btnLogin, btnLogOut; 
        IWebElement edtAccount;


        Excel.Application dataApp;
        Excel.Workbook dataWorkbook;
        Excel.Worksheet dataWorksheet;
        Excel.Range xlRange;

        int iRegistName, iRegistEmail, iRegistPhone, iRegistPassword, iRegistRePassword;

        int iLogInAccount, iLogInPassword;

        string edtUserEmail;
        string edtUserPassword;

        [TestInitialize]
        public void OpenWebsite()
        {
            driver = new ChromeDriver();
            driver.Url = "https://localhost:44366/Users/Login";
            driver.Navigate();
            dataApp = new Excel.Application();
            dataWorkbook = dataApp.Workbooks.Open(@"D:\\Learn\\DBCLPM\\DoAn_DBCLPM\\DoAn_DBCLPM\\Final_TestCase.xlsx");
            dataWorksheet = dataWorkbook.Sheets[1];
            xlRange = dataWorksheet.UsedRange;
        }

        public void SignUpSuccess(int iName, int iEmail, int iPhone, int iPassword, int iRePassword)
        {
            edtName = driver.FindElement(By.Name("UserName"));
            edtEmail = driver.FindElement(By.Name("UserEmail"));
            edtPhone = driver.FindElement(By.Name("PhoneNumber"));
            edtPassword = driver.FindElement(By.Name("UserPassword"));
            edtRePassword = driver.FindElement(By.Name("RePassword"));
            btnSignUp = driver.FindElement(By.XPath("/html/body/div[3]/div/div/div[2]/form/div/div[6]/div[2]/button"));
            chkBoxSignUp = driver.FindElement(By.XPath("/html/body/div[3]/div/div/div[2]/form/div/div[6]/div[1]"));

            
            // Điền thông tin vào các trường
            edtName.SendKeys(xlRange.Cells[6][iName].value.ToString());
            edtEmail.SendKeys(xlRange.Cells[6][iEmail].value.ToString());
            edtPhone.SendKeys(xlRange.Cells[6][iPhone].value.ToString());
            edtPassword.SendKeys(xlRange.Cells[6][iPassword].value.ToString());
            edtRePassword.SendKeys(xlRange.Cells[6][iRePassword].value.ToString());

            // Chọn checkbox đồng ý đăng ký

            chkBoxSignUp.Click();
            btnSignUp.Click();         
            
            if (driver.Url == "https://localhost:44366/Users/Login")
            {
                Console.WriteLine("Đăng ký tài khoản thành công");
            }
            else
            {
                Console.WriteLine("Đăng ký thất bại");
            }

        }

        public void LoginSuccess(int iAccount, int iPassword)
        {
            //UserAccount
            edtAccount = driver.FindElement(By.Name("UserEmail"));

            //UserPassword
            edtPassword = driver.FindElement(By.Name("UserPassword"));

            //Test btnLogin
            btnLogin = driver.FindElement(By.XPath("/html/body/div[3]/div/div/div[2]/form/input"));

            // Lấy data từ Excel để truyền vào 
            edtAccount.SendKeys(xlRange.Cells[6][iAccount].value.ToString());
            edtPassword.SendKeys(xlRange.Cells[6][iPassword].value.ToString());

            // Nhấp vào nút
            btnLogin.Click();

            if (driver.Url == "https://localhost:44366/Admin/AdminHome")
            {
                Console.WriteLine("Đăng nhập trang Admin thành công");
            }
            else
            {
                Console.WriteLine("Đăng nhập thất bại");
            }
        }

        public void NavigateQLUser()
        {
            sideBarQLUsers = driver.FindElement(By.XPath("/html/body/div[1]/div[2]/ul/li[2]/a"));
             sideBarQLUsers.Click();
             if (driver.Url == "https://localhost:44366/Admin/CustomersAdmin")
             {
                 Console.WriteLine("Chuyển hướng đến trang QLUsers");
             }
             else
             {
                 Console.WriteLine("Không chuyển hướng đến trang User");
             }
        }
        //
        // Test Button close sideBar
        //
        [TestMethod]
        public void ID_40_AnSideBar()
        {
            string sucOrFail;
            iLogInAccount = 126;
            iLogInPassword = 127;
            LoginSuccess(iLogInAccount, iLogInPassword);
            // Tìm và nhấn vào nút trên thanh taskbar
            btnMenuSideBar = driver.FindElement(By.XPath("/html/body/div[2]/div[1]/div[1]/label"));
            btnMenuSideBar.Click();

            // Kiểm tra xem menu đã hiển thị ra hay không
            menuSideBar = driver.FindElement(By.XPath("/html/body/div[1]"));
            if (!menuSideBar.Displayed)
            {
                sucOrFail = "Pass";
                Console.WriteLine("Menu sidebar đã bị ẩn.");
                dataWorksheet.Cells[12][126] = sucOrFail;
            }
            else
            {
                sucOrFail = "Fail";
                Console.WriteLine("Menu không bị ẩn.");
                dataWorksheet.Cells[12][126] = sucOrFail;
            }      
            Thread.Sleep(7000);
        }

        //
        // Test Button open sideBar
        //
        [TestMethod]
        public void ID_41_HienThiSideBar()
        {
            string sucOrFail;
            iLogInAccount = 128;
            iLogInPassword = 129;
            LoginSuccess(iLogInAccount, iLogInPassword);
            // Tìm và nhấn vào nút trên thanh taskbar
            btnMenuSideBar = driver.FindElement(By.XPath("/html/body/div[2]/div[1]/div[1]/label"));
            btnMenuSideBar.Click();
            Thread.Sleep(1000);
            btnMenuSideBar.Click();
            // Kiểm tra xem menu đã hiển thị ra hay không
            menuSideBar = driver.FindElement(By.XPath("/html/body/div[1]"));
            if (!menuSideBar.Displayed)
            {
                sucOrFail = "Fail";
                Console.WriteLine("Menu sidebar không hiển thị.");
                dataWorksheet.Cells[12][128] = sucOrFail;
            }
            else
            {
                sucOrFail = "Pass";
                Console.WriteLine("Menu sidebar hiển thị.");
                dataWorksheet.Cells[12][128] = sucOrFail;
            }
            Thread.Sleep(7000);
        }

        //
        //Test thông tin sau khi đăng ký 
        //
        [TestMethod]
        public void ID_34_ThongTinSauDangKy()
        {
            string sucOrFail;
            //Assert.IsTrue(IsAccountInfoCorrect(), "Thông tin tài khoản không chính xác sau khi đăng ký.");
            //IsAccountInfoCorrect();
            if (IsAccountInfoCorrect() == true)
            {
                sucOrFail = "Pass";
                dataWorksheet.Cells[12][99] = sucOrFail;
                Console.WriteLine("Đã lưu vào Excel");
            }
            else
            {
                sucOrFail = "Fail";
                dataWorksheet.Cells[12][99] = sucOrFail;
                Console.WriteLine("Chưa lưu vào Excel");
            }
            Thread.Sleep(15000);
        }
        private bool IsAccountInfoCorrect()
        {
            
            btnNavRegist = driver.FindElement(By.XPath("/html/body/div[3]/div/div/div[2]/form/div[3]/a"));
            btnNavRegist.Click();

            iRegistName = 99;
            iRegistEmail = 100;
            iRegistPhone = 101;
            iRegistPassword = 102;
            iRegistRePassword = 103;
            SignUpSuccess(iRegistName, iRegistEmail, iRegistPhone, iRegistPassword, iRegistRePassword);


            iLogInAccount = 104;
            iLogInPassword = 105;
            LoginSuccess(iLogInAccount, iLogInPassword);


            NavigateQLUser();

            // Tìm các phần tử trên trang quản lý tài khoản để lấy thông tin tài khoản
            UserName = driver.FindElement(By.XPath("/html/body/div[2]/div[2]/table/tbody/tr[10]/td[1]"));
            UserEmail = driver.FindElement(By.XPath("/html/body/div[2]/div[2]/table/tbody/tr[10]/td[2]"));
            PhoneNumber = driver.FindElement(By.XPath("/html/body/div[2]/div[2]/table/tbody/tr[10]/td[3]"));
            // Ví dụ: tìm các phần tử khác để lấy thông tin khác nếu cần

            // Lấy thông tin từ các phần tử
            string displayedUsername = UserName.Text;
            string displayedEmail = UserEmail.Text;
            string displayedPhone = PhoneNumber.Text;
            // Ví dụ: lấy thông tin khác nếu cần

            // Lấy dữ liệu từ ô Excel
            string expectedUsername = xlRange.Cells[6][99].Value.ToString(); // Thay thế bằng thông tin tên người dùng mong đợi
            string expectedEmail = xlRange.Cells[6][100].Value.ToString(); // Thay thế bằng địa chỉ email mong đợi
            string expectedPhone = xlRange.Cells[6][101].Value.ToString(); // Thay thế bằng số điện thoại mong đợi

            // Ví dụ: thêm thông tin khác nếu cần


            // Kiểm tra xem thông tin hiển thị có chính xác không
            

            if (displayedUsername == expectedUsername && displayedEmail == expectedEmail && displayedPhone == expectedPhone)
            {
                
                Console.WriteLine("Thông tin chính xác");
                return true;

            }
            else
            {
                
                Console.WriteLine("Thông tin khoong chính xác");
                return false;
            }
        }

        //
        //Test nhấn vào Button Edit
        //
        [TestMethod]
        public void ID_38_NhanButtonChinhSua()
        {
            string sucOrFail;
            iLogInAccount = 123;
            iLogInPassword = 124;
            LoginSuccess(iLogInAccount, iLogInPassword);

            NavigateQLUser();

            btnEditAccount = driver.FindElement(By.XPath("/html/body/div[2]/div[2]/table/tbody/tr[2]/td[4]/a[1]"));
            btnEditAccount.Click();

            // Kiểm tra xem URL hiện tại có trùng với URL của trang web mà bạn mong muốn không
            string editUrl = "https://localhost:44366/Admin/CustomersAdmin/Edit/1";
            //Assert.AreEqual(editUrl, driver.Url);
            if (editUrl == driver.Url)
            {
                sucOrFail = "Pass";
                dataWorksheet.Cells[12][123] = sucOrFail;
                Console.WriteLine("Success");
            }
            else
            {
                sucOrFail = "Fail";
                dataWorksheet.Cells[12][123] = sucOrFail;
                Console.WriteLine("Fail");
            }
            Thread.Sleep(7000);
        }

        //
        //Test nhấn vào Button Details
        //
        [TestMethod]
        public void ID_37_XemChiTietUser()
        {

            string sucOrFail;
            iLogInAccount = 120;
            iLogInPassword = 121;
            LoginSuccess(iLogInAccount, iLogInPassword);

            NavigateQLUser();

            btnDetailsAccount = driver.FindElement(By.XPath("/html/body/div[2]/div[2]/table/tbody/tr[2]/td[4]/a[2]"));
            btnDetailsAccount.Click();
            // Kiểm tra xem URL hiện tại có trùng với URL của trang web mà bạn mong muốn không
            string detailsUrl = "https://localhost:44366/Admin/CustomersAdmin";
            
            if (detailsUrl == driver.Url)
            {
                sucOrFail = "Pass";
                dataWorksheet.Cells[12][120] = sucOrFail;
                Console.WriteLine("Success");
            }
            else
            {
                sucOrFail = "Fail";
                dataWorksheet.Cells[12][120] = sucOrFail;
                Console.WriteLine("Fail");
            }
            Thread.Sleep(7000);
        }

        //
        //Test Edit Account 
        //
        [TestMethod]
        public void ID_35_ChinhSuaUser()
        {

            string sucOrFail;
            iLogInAccount = 106;
            iLogInPassword = 107;
            LoginSuccess(iLogInAccount, iLogInPassword);

            NavigateQLUser();

            // Thực hiện chỉnh sửa thông tin tài khoản
            EditAccountInfo();
            // Kiểm tra xem thay đổi đã được lưu thành công
            if (IsEditSuccessful())
            {
                sucOrFail = "Pass";
                dataWorksheet.Cells[12][106] = sucOrFail;
                Console.WriteLine("Đã lưu vào Excel");
            }
            else
            {
                sucOrFail = "Fail";
                dataWorksheet.Cells[12][106] = sucOrFail;
                Console.WriteLine("Chưa lưu vào Excel");
            }
            
            
        }
        private void EditAccountInfo()
        {
            btnEditAccount = driver.FindElement(By.XPath("/html/body/div[2]/div[2]/table/tbody/tr[2]/td[4]/a[1]"));
            btnEditAccount.Click();
            // Điền thông tin mới vào các trường chỉnh sửa thông tin tài khoản
            UserName = driver.FindElement(By.Id("UserName"));
            UserName.Clear();
            UserName.SendKeys(xlRange.Cells[6][108].value.ToString());

            UserEmail = driver.FindElement(By.Id("UserEmail"));
            UserEmail.Clear();
            UserEmail.SendKeys(xlRange.Cells[6][109].value.ToString());

            PhoneNumber = driver.FindElement(By.Id("PhoneNumber"));
            PhoneNumber.Clear();
            PhoneNumber.SendKeys(xlRange.Cells[6][110].value.ToString());

            UserPassword = driver.FindElement(By.Id("UserPassword"));
            UserPassword.Clear();
            UserPassword.SendKeys(xlRange.Cells[6][111].value.ToString());

            AvaterImage = driver.FindElement(By.Id("AvatarImage"));
            AvaterImage.Clear();
            AvaterImage.SendKeys(xlRange.Cells[6][112].value.ToString());
            // Thực hiện các bước khác nếu cần (ví dụ: thay đổi mật khẩu, v.v.)

            edtUserEmail = UserEmail.GetAttribute("value").ToString();
            edtUserPassword = UserPassword.GetAttribute("value").ToString();

            // Lưu thay đổi bằng cách nhấn nút lưu
            IWebElement saveButton = driver.FindElement(By.XPath("/html/body/div[2]/div[2]/form/div/div[6]/div/input"));
            saveButton.Click();
        }

        private bool IsEditSuccessful()
        {
            // Kiểm tra xem URL hiện tại có trùng khớp với URL của trang Quản Lý Tài Khoản hay không
            string expectedUrl = "https://localhost:44366/Admin/CustomersAdmin";
            if (driver.Url == expectedUrl)
            {
                Console.WriteLine(" Edit success");
                return true;
            }
            else
            {
                Console.WriteLine("Edit failed");
                return false;
            }
        }
        //Test button Back To List
        [TestMethod]
        public void ID_42_BackToList()
        {
            string sucOrFail;
            iLogInAccount = 131;
            iLogInPassword = 132;
            LoginSuccess(iLogInAccount, iLogInPassword);

            NavigateQLUser();

            btnEditAccount = driver.FindElement(By.XPath("/html/body/div[2]/div[2]/table/tbody/tr[2]/td[4]/a[1]"));
            btnEditAccount.Click();

            IWebElement hrefBackToList = driver.FindElement(By.XPath("/html/body/div[2]/div[2]/div/a"));
            hrefBackToList.Click();
            if (driver.Url == "https://localhost:44366/Admin/CustomersAdmin")
            {
                sucOrFail = "Pass";
                dataWorksheet.Cells[12][131] = sucOrFail;
                Console.WriteLine("Chuyển hướng đến trang đúng.");
            }
            else
            {
                sucOrFail = "Fail";
                dataWorksheet.Cells[12][131] = sucOrFail;
                Console.WriteLine("Chuyển hướng không thành công hoặc đến trang không mong muốn.");
            }
            Thread.Sleep(7000);
        }

        //Test button Edit Account in details user
        [TestMethod]
        public void ID_43_ChinhSuaTrongChiTiet()
        {
            string sucOrFail;
            iLogInAccount = 135;
            iLogInPassword = 136;
            LoginSuccess(iLogInAccount, iLogInPassword);

            NavigateQLUser();

            btnDetailsAccount = driver.FindElement(By.XPath("/html/body/div[2]/div[2]/table/tbody/tr[2]/td[4]/a[2]"));
            btnDetailsAccount.Click();

            btnEditAccount = driver.FindElement(By.XPath("/html/body/div[2]/div[2]/div/p/a[1]"));
            btnEditAccount.Click();

            if (driver.Url == "https://localhost:44366/Admin/CustomersAdmin/Edit/1")
            {
                sucOrFail = "Pass";
                dataWorksheet.Cells[12][135] = sucOrFail;
                Console.WriteLine("Chuyển hướng đến trang đúng.");
            }
            else
            {
                sucOrFail = "Fail";
                dataWorksheet.Cells[12][135] = sucOrFail;
                Console.WriteLine("Chuyển hướng không thành công hoặc đến trang không mong muốn.");
            }
            Thread.Sleep(7000);
        }

        //Test button Turn back in details user
        [TestMethod]
        public void ID_44_QuayLaiTrongChiTiet()
        {
            string sucOrFail;
            iLogInAccount = 139;
            iLogInPassword = 140;
            LoginSuccess(iLogInAccount, iLogInPassword);

            NavigateQLUser();

            btnDetailsAccount = driver.FindElement(By.XPath("/html/body/div[2]/div[2]/table/tbody/tr[2]/td[4]/a[2]"));
            btnDetailsAccount.Click();

            IWebElement btnTurnBack = driver.FindElement(By.XPath("/html/body/div[2]/div[2]/div/p/a[2]"));
            btnTurnBack.Click();

            if (driver.Url == "https://localhost:44366/Admin/CustomersAdmin")
            {
                sucOrFail = "Pass";
                dataWorksheet.Cells[12][139] = sucOrFail;
                Console.WriteLine("Chuyển hướng đến trang đúng.");
            }
            else
            {
                sucOrFail = "Fail";
                dataWorksheet.Cells[12][139] = sucOrFail;
                Console.WriteLine("Chuyển hướng không thành công hoặc đến trang không mong muốn.");
            }
            Thread.Sleep(7000);
        }

        ///////// Test đăng nhập vào tài khoản vừa edit thông tin
        [TestMethod]
        public void ID_36_DangNhapUserVuaEdit()
        {
            iLogInAccount = 113;
            iLogInPassword = 114;
            LoginSuccess(iLogInAccount, iLogInPassword);

            NavigateQLUser();

            EditAccountInfo();

            //UserName = driver.FindElement(By.XPath("//*[@id=\"UserName\"]"));
            //UserPassword = driver.FindElement(By.XPath("//*[@id=\"UserPassword\"]"));

            

            //LogOut
            btnLogOut = driver.FindElement(By.XPath("/html/body/div[2]/div[1]/div[2]/div/a"));
            btnLogOut.Click();
            if(driver.Url == "https://localhost:44366/Users/Login")
            {
                Console.WriteLine("Log out successfully");
            }
            else
            {
                Console.WriteLine("Log out fail");
            }
            CheckEditLogIn();
        }

        public void CheckEditLogIn()
        {
            string sucOrFail;
            //UserAccount
            edtAccount = driver.FindElement(By.Name("UserEmail"));

            //UserPassword
            edtPassword = driver.FindElement(By.Name("UserPassword"));

            //Test btnLogin
            btnLogin = driver.FindElement(By.XPath("/html/body/div[3]/div/div/div[2]/form/input"));

            // Lấy data từ Excel để truyền vào 
            edtAccount.SendKeys(edtUserEmail);
            edtPassword.SendKeys(edtUserPassword);
            Thread.Sleep(1000);
            // Nhấp vào nút
            btnLogin.Click();
            
            if (driver.Url == "https://localhost:44366/")
            {
                sucOrFail = "Pass";
                dataWorksheet.Cells[12][113] = sucOrFail;
                Console.WriteLine("Success");
            }
            else
            {
                sucOrFail = "Fail";
                dataWorksheet.Cells[12][113] = sucOrFail;
                Console.WriteLine("Fail");
            }

        }

        
        [TestCleanup]
        public void Cleanup()
        {
            dataWorkbook.Save();
            dataWorkbook.Close();
            dataApp.Quit();
            driver.Quit();
        }
    }
}
