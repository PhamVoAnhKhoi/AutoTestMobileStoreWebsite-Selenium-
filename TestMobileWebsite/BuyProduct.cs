using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;

namespace UnitTestProject1
{
    [TestClass]
    public class BuyProduct
    {
        Excel.Application datatApp;
        Excel.Workbook dataWorkBook;
        Excel.Worksheet dataWorksheet;
        Excel.Range xlRange;

        IWebDriver driver;
        [TestInitialize]
        public void OpenWebsite()
        {
            driver = new ChromeDriver();

            driver.Url = "https://localhost:44366/Users/Login";
            driver.Navigate();

            datatApp = new Excel.Application();
            dataWorkBook = datatApp.Workbooks.Open(@"D:\\Learn\\DBCLPM\\DoAn_DBCLPM\\DoAn_DBCLPM\\Final_TestCase - Copy.xlsx");
            dataWorksheet = dataWorkBook.Sheets[1];
            xlRange = dataWorksheet.UsedRange;     
        }
        [TestCleanup]
        public void Cleanup()
        {
            dataWorkBook.Save();
            dataWorkBook.Close();
            datatApp.Quit();
            driver.Quit();
        }


        // *********************TEST SELENIUM CHỨC NĂNG MUA HÀNG **************************************/

        /*Chọn sản phẩm,chỉnh số lượng và thêm vào giỏ*/
        [TestMethod]
        [Priority(1)]
        public void ID_001_MuaHang()
        {
           
            driver.FindElement(By.Name("UserEmail"));
            driver.FindElement(By.Name("UserPassword"));
            driver.FindElement(By.XPath("/html/body/div[3]/div/div/div[2]/form/input")).SendKeys(Keys.Enter);
            
            string Suc = " success";
            string Failed = " failed";
            driver.FindElement(By.Name("UserEmail")).SendKeys(xlRange.Cells[6][52].value.ToString());
            driver.FindElement(By.Name("UserPassword")).SendKeys(xlRange.Cells[6][53].value.ToString());
            driver.FindElement(By.XPath("/html/body/div[3]/div/div/div[2]/form/input")).Click();

           
            Thread.Sleep(2000);

            if (!driver.Url.Equals("https://localhost:44366/Users/Login"))
            {
                driver.Navigate().GoToUrl("  https://localhost:44366/");
                driver.FindElement(By.XPath("/html/body/div[3]/div[2]/div/div[2]/div[2]/div[1]/div/div[8]/div/a/div/img")).Click();
                // Tìm và click vào button để tăng, giảm số lượng sản phẩm
                driver.FindElement(By.XPath("/html/body/div[3]/div/div[1]/div[2]/div[4]/form[1]/button")).Click();
                Thread.Sleep(1000);
                driver.FindElement(By.XPath("/html/body/div[3]/div/div/div[2]/div/form/input[2]")).Click();
                // Thử đặt hàng  khi chưa nhập dia chỉ(báo lỗi:Please fill out this field) 
                driver.FindElement(By.XPath("/html/body/div[3]/div/div[2]/form/div/div[2]/div/button[1]")).Click();
                Thread.Sleep(2000);

                driver.FindElement(By.XPath("/html/body/div[3]/div/div[2]/form/div/div[1]/table/tbody/tr[4]/td/textarea")).SendKeys("828-sư vạn hạnh-q10-tphcm");
                Thread.Sleep(1000);
                driver.FindElement(By.XPath("/html/body/div[3]/div/div[2]/form/div/div[2]/div/button[1]")).Click();
                Thread.Sleep(2000);
                if (driver.Url == "https://localhost:44366/Order/GetOrder/1020")
                {
                    xlRange[11][4].value = Suc;
                }
                else
                {
                    xlRange[11][4].value = Failed;
                }

            }

        }
        [TestMethod]
        [Priority(2)]
        public void ID_002_ThemSpYeuThich()
        {
            
            driver.FindElement(By.Name("UserEmail"));
            driver.FindElement(By.Name("UserPassword"));
            driver.FindElement(By.XPath("/html/body/div[3]/div/div/div[2]/form/input")).SendKeys(Keys.Enter);
           
            string Suc = " success";
            string Failed = " failed";
            driver.FindElement(By.Name("UserEmail")).SendKeys(xlRange.Cells[6][52].value.ToString());
            driver.FindElement(By.Name("UserPassword")).SendKeys(xlRange.Cells[6][53].value.ToString());
            driver.FindElement(By.XPath("/html/body/div[3]/div/div/div[2]/form/input")).Click();

           

            if (!driver.Url.Equals("https://localhost:44366/Users/Login"))
            {
                driver.FindElement(By.CssSelector(".product-list:nth-child(2) .owl-item:nth-child(8) .card-img-top")).Click();
                Thread.Sleep(1000);
                driver.FindElement(By.CssSelector(".bi-heart-fill")).Click();
                Thread.Sleep(1000);
                driver.FindElement(By.CssSelector(".logo-image")).Click();
                Thread.Sleep(1000);
                driver.FindElement(By.CssSelector(".user-avatar")).Click();
                Thread.Sleep(1000);
                driver.FindElement(By.LinkText("Sản phẩm yêu thích")).Click();
                Thread.Sleep(1000);
                if (driver.Url != "https://localhost:44366/Users/Login")
                {
                    xlRange[11][8].value = Suc;
                }
                else
                {
                    xlRange[11][8].value = Failed;
                }
                Thread.Sleep(2000);
            }


        }
        [TestMethod]
        [Priority(3)]
        public void ID_003_KtraIconGioHang()
        {
            
            driver.FindElement(By.Name("UserEmail"));
            driver.FindElement(By.Name("UserPassword"));
            driver.FindElement(By.XPath("/html/body/div[3]/div/div/div[2]/form/input")).SendKeys(Keys.Enter);
            
            string Suc = " success";
            string Failed = " failed";
            driver.FindElement(By.Name("UserEmail")).SendKeys(xlRange.Cells[6][52].value.ToString());
            driver.FindElement(By.Name("UserPassword")).SendKeys(xlRange.Cells[6][53].value.ToString());
            driver.FindElement(By.XPath("/html/body/div[3]/div/div/div[2]/form/input")).Click();

            if (driver.Url != "https://localhost:44366/Users/Login")
            {
                xlRange[11][12].value = Suc;
            }
            else
            {
                xlRange[11][12].value = Failed;
            }
            Thread.Sleep(2000);
            if (!driver.Url.Equals("https://localhost:44366/Users/Login"))
            {
                driver.FindElement(By.CssSelector(".bi-cart3")).Click();
                Thread.Sleep(1000);
                driver.FindElement(By.LinkText("Tiếp tục mua sắm >>")).Click();
                Thread.Sleep(1000);
                driver.FindElement(By.CssSelector(".product-list:nth-child(2) .owl-item:nth-child(8) .card-img-top")).Click();
                Thread.Sleep(1000);
                driver.FindElement(By.Id("counter-plus")).Click();
                Thread.Sleep(1000);
                driver.FindElement(By.CssSelector(".add-to-cart-btn")).Click();
                Thread.Sleep(1000);
                driver.FindElement(By.LinkText("Tiếp tục mua sắm >>")).Click();
                Thread.Sleep(1000);
                driver.FindElement(By.CssSelector(".bi-cart3")).Click();
                Thread.Sleep(1000);

            }

        }
        [TestMethod]
        [Priority(4)]
        public void ID_004_KtraSLSPIconGiohang()
        {
            
            driver.FindElement(By.Name("UserEmail"));
            driver.FindElement(By.Name("UserPassword"));
            driver.FindElement(By.XPath("/html/body/div[3]/div/div/div[2]/form/input")).SendKeys(Keys.Enter);
            
            string Suc = " success";
            string Failed = " failed";
            driver.FindElement(By.Name("UserEmail")).SendKeys(xlRange.Cells[6][52].value.ToString());
            driver.FindElement(By.Name("UserPassword")).SendKeys(xlRange.Cells[6][53].value.ToString());
            driver.FindElement(By.XPath("/html/body/div[3]/div/div/div[2]/form/input")).Click();


            Thread.Sleep(2000);
            if (!driver.Url.Equals("https://localhost:44366/Users/Login"))
            {
                driver.FindElement(By.CssSelector(".product-list:nth-child(2) .owl-item:nth-child(8) .card-img-top")).Click();
                Thread.Sleep(1000);
                driver.FindElement(By.XPath("/html/body/div[3]/div/div[1]/div[2]/div[4]/form[1]/button")).Click();
                Thread.Sleep(1000);
                driver.FindElement(By.CssSelector(".cart-item:nth-child(2) .delete-product-cart > .bi")).Click();
                Thread.Sleep(1000);
                driver.FindElement(By.CssSelector(".bi-trash3-fill")).Click();
                Thread.Sleep(1000);

                // Kiểm tra số lượng sản phẩm trong biểu tượng giỏ hàng
                int itemsInCart = int.Parse(driver.FindElement(By.CssSelector("div.count-product")).Text);
                if (itemsInCart == 0)
                {
                    Console.WriteLine("Item deleted successfully from the cart.");
                    xlRange[11][16].value = Suc;
                }
                else
                {
                    Console.WriteLine("Failed to delete item from the cart.");
                    xlRange[11][16].value = Failed;
                }

            }


        }
        [TestMethod]
        [Priority(5)]
        public void ID_005_BinhLuan()
        {
          
            driver.FindElement(By.Name("UserEmail"));
            driver.FindElement(By.Name("UserPassword"));
            driver.FindElement(By.XPath("/html/body/div[3]/div/div/div[2]/form/input")).SendKeys(Keys.Enter);
           
            string Suc = " success";
            string Failed = " failed";
            driver.FindElement(By.Name("UserEmail")).SendKeys(xlRange.Cells[6][52].value.ToString());
            driver.FindElement(By.Name("UserPassword")).SendKeys(xlRange.Cells[6][53].value.ToString());
            driver.FindElement(By.XPath("/html/body/div[3]/div/div/div[2]/form/input")).Click();

            if (driver.Url != "https://localhost:44366/Users/Login")
            {
                xlRange[11][20].value = Suc;
            }
            else
            {
                xlRange[11][20].value = Failed;
            }

            Thread.Sleep(2000);
            driver.FindElement(By.CssSelector(".product-list:nth-child(2) .owl-item:nth-child(8) .card-img-top")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.XPath("/html/body/div[3]/div/div[4]/div[1]/div[2]/form/button")).Click(); // Nhấn gửi khi chưa viết bình luận
            Thread.Sleep(1000);
            driver.FindElement(By.XPath("/html/body/div[3]/div/div[4]/div[1]/div[2]/form/textarea")).SendKeys("Tuyệt vời");
            Thread.Sleep(1000);
            driver.FindElement(By.XPath("/html/body/div[3]/div/div[4]/div[1]/div[2]/form/button")).Click();
            Thread.Sleep(1000);

        }
        [TestMethod]
        [Priority(6)]
        public void ID_006_XoaBinhLuan()
        {
           
            driver.FindElement(By.Name("UserEmail"));
            driver.FindElement(By.Name("UserPassword"));
            driver.FindElement(By.XPath("/html/body/div[3]/div/div/div[2]/form/input")).SendKeys(Keys.Enter);
           
            string Suc = " success";
            string Failed = " failed";
            driver.FindElement(By.Name("UserEmail")).SendKeys(xlRange.Cells[6][52].value.ToString());
            driver.FindElement(By.Name("UserPassword")).SendKeys(xlRange.Cells[6][53].value.ToString());
            driver.FindElement(By.XPath("/html/body/div[3]/div/div/div[2]/form/input")).Click();

            if (driver.Url != "https://localhost:44366/Users/Login")
            {
                xlRange[11][24].value = Suc;
            }
            else
            {
                xlRange[11][24].value = Failed;
            }

            Thread.Sleep(2000);
            driver.FindElement(By.CssSelector(".product-list:nth-child(2) .owl-item:nth-child(8) .card-img-top")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.XPath("/html/body/div[3]/div/div[4]/div[1]/div[2]/form/textarea")).SendKeys("Tuyệt vời");
            Thread.Sleep(1000);
            driver.FindElement(By.XPath("/html/body/div[3]/div/div[4]/div[1]/div[2]/form/button")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.XPath("/html/body/div[3]/div/div[4]/div[2]/div[1]/div[2]/a/i")).Click();
            Thread.Sleep(1000);

        }
        [TestMethod]
        [Priority(7)]

        public void ID_007_MuaSPHetHang()
        {
          
            driver.FindElement(By.Name("UserEmail"));
            driver.FindElement(By.Name("UserPassword"));
            driver.FindElement(By.XPath("/html/body/div[3]/div/div/div[2]/form/input")).SendKeys(Keys.Enter);
           
            string Suc = " success";
            string Failed = " failed";
            driver.FindElement(By.Name("UserEmail")).SendKeys(xlRange.Cells[6][52].value.ToString());
            driver.FindElement(By.Name("UserPassword")).SendKeys(xlRange.Cells[6][53].value.ToString());
            driver.FindElement(By.XPath("/html/body/div[3]/div/div/div[2]/form/input")).Click();

            if (driver.Url != "https://localhost:44366/Users/Login")
            {
                xlRange[11][29].value = Suc;
            }
            else
            {
                xlRange[11][29].value = Failed;
            }
            Thread.Sleep(2000);
            if (!driver.Url.Equals("https://localhost:44366/Users/Login"))
            {
                driver.FindElement(By.CssSelector(".product-list:nth-child(3) .owl-item:nth-child(8) .card-img-top")).Click();
                Thread.Sleep(1000);
                driver.FindElement(By.CssSelector(".logo-image")).Click();
                Thread.Sleep(1000);
            }

            IWebElement addToCartButton = driver.FindElement(By.ClassName("add-to-cart-btn"));
            if (addToCartButton.GetAttribute("class").Contains("sold-out-btn"))
            {
                Console.WriteLine("Sản phẩm đã hết hàng và không thể thêm vào giỏ hàng.");
            }
            else
            {
                // Nếu sản phẩm không hết hàng, thực hiện các hành động khác ở đây
                addToCartButton.Click();
                Console.WriteLine("Sản phẩm đã được thêm vào giỏ hàng.");
            }



        }
        [TestMethod]
        [Priority(8)]
        public void ID_008_MuaSpTuTrangChiTiet()
        {
           
            driver.FindElement(By.Name("UserEmail"));
            driver.FindElement(By.Name("UserPassword"));
            driver.FindElement(By.XPath("/html/body/div[3]/div/div/div[2]/form/input")).SendKeys(Keys.Enter);
           
            string Suc = " success";
            string Failed = " failed";
            driver.FindElement(By.Name("UserEmail")).SendKeys(xlRange.Cells[6][52].value.ToString());
            driver.FindElement(By.Name("UserPassword")).SendKeys(xlRange.Cells[6][53].value.ToString());
            driver.FindElement(By.XPath("/html/body/div[3]/div/div/div[2]/form/input")).Click();

            if (driver.Url != "https://localhost:44366/Users/Login")
            {
                xlRange[11][32].value = Suc;
            }
            else
            {
                xlRange[11][32].value = Failed;
            }
            Thread.Sleep(2000);

            if (!driver.Url.Equals("https://localhost:44366/Users/Login"))
            {
                driver.FindElement(By.CssSelector(".product-list:nth-child(2) .owl-item:nth-child(8) .card-img-top")).Click();
                Thread.Sleep(1000);
                driver.FindElement(By.Name("searchString")).Click();
                Thread.Sleep(1000);
                driver.FindElement(By.Name("searchString")).SendKeys("iphone 11");
                Thread.Sleep(1000);
                driver.FindElement(By.Name("searchString")).SendKeys(Keys.Enter);
                Thread.Sleep(1000);
                driver.FindElement(By.CssSelector(".card-img-top")).Click();
                Thread.Sleep(1000);

            }


        }
        [TestMethod]
        [Priority(9)]
        public void ID_009_MuaKhiChuaDangNhap()
        {
            driver.Navigate().GoToUrl("https://localhost:44366/");
            Thread.Sleep(1000);
            driver.FindElement(By.CssSelector(".product-list:nth-child(2) .owl-item:nth-child(8) a")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.Id("counter-plus")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.CssSelector(".add-to-cart-btn")).Click();
            Thread.Sleep(1000);
            string Suc = "Pass";
            string Failed = "Fail";
            if (driver.Url != "https://localhost:44366/Users/Login")
            {
                xlRange[11][36].value = Suc;
            }
            else
            {
                xlRange[11][36].value = Failed;
            }
            Thread.Sleep(2000);

        }
        [TestMethod]
        [Priority(11)]
        public void ID_010_TimSpTuTrangChiTiet()
        {
            driver.Navigate().GoToUrl("https://localhost:44366/Users/Login");
            driver.Manage().Window.Size = new System.Drawing.Size(1536, 824);
            driver.FindElement(By.Name("UserEmail"));
            driver.FindElement(By.Name("UserPassword"));
            driver.FindElement(By.XPath("/html/body/div[3]/div/div/div[2]/form/input")).SendKeys(Keys.Enter);
            string Suc = "Pass";
            string Failed = "Fail";
            driver.FindElement(By.Name("UserEmail")).SendKeys(xlRange.Cells[6][52].value.ToString());
            driver.FindElement(By.Name("UserPassword")).SendKeys(xlRange.Cells[6][53].value.ToString());
            driver.FindElement(By.XPath("/html/body/div[3]/div/div/div[2]/form/input")).Click();

            

            if (!driver.Url.Equals("https://localhost:44366/Users/Login"))
            {
                driver.FindElement(By.Name("searchString")).Click();
                Thread.Sleep(1000);
                driver.FindElement(By.Name("searchString")).SendKeys("iphone 11");
                Thread.Sleep(1000);
                driver.FindElement(By.CssSelector(".bi-search")).Click();
                Thread.Sleep(1000);
                driver.FindElement(By.CssSelector(".card-img-top")).Click();
                Thread.Sleep(1000);
                driver.FindElement(By.CssSelector(".owl-item:nth-child(8) .card-img-top")).Click();
                Thread.Sleep(1000);
                driver.FindElement(By.CssSelector(".owl-item:nth-child(6) .card-img-top")).Click();
                Thread.Sleep(1000);
                driver.FindElement(By.CssSelector(".logo-image")).Click();
                Thread.Sleep(1000);
                
            }
            if (driver.Url != "https://localhost:44366/Users/Login")
            {
                xlRange[11][40].value = Suc;
            }
            else
            {
                xlRange[11][40].value = Failed;
            }
            Thread.Sleep(2000);
        }
        [TestMethod]
        [Priority(12)]
        public void ID_011_HuyDonHang()
        {
            driver.Navigate().GoToUrl("https://localhost:44366/Users/Login");
            driver.Manage().Window.Size = new System.Drawing.Size(1536, 824);
            driver.FindElement(By.Name("UserEmail"));
            driver.FindElement(By.Name("UserPassword"));
            driver.FindElement(By.XPath("/html/body/div[3]/div/div/div[2]/form/input")).SendKeys(Keys.Enter);
            
            string Suc = "Pass";
            string Failed = "Fail";
            driver.FindElement(By.Name("UserEmail")).SendKeys(xlRange.Cells[6][52].value.ToString());
            driver.FindElement(By.Name("UserPassword")).SendKeys(xlRange.Cells[6][53].value.ToString());
            driver.FindElement(By.XPath("/html/body/div[3]/div/div/div[2]/form/input")).Click();

            if (driver.Url != "https://localhost:44366/Users/Login")
            {
                xlRange[11][41].value = Suc;
            }
            else
            {
                xlRange[11][41].value = Failed;
            }
            Thread.Sleep(2000);

            if (!driver.Url.Equals("https://localhost:44366/Users/Login"))
            {
                driver.FindElement(By.CssSelector(".user-avatar")).Click();
                driver.FindElement(By.LinkText("Đơn hàng của bạn")).Click();
                driver.FindElement(By.CssSelector(".cart-item:nth-child(3) a")).Click();

                driver.FindElement(By.LinkText("Hủy đơn hàng")).Click();
                if (driver.Url.Equals("https://localhost:44366/Order/GetOrder/1013"))
                {
                    Console.WriteLine("Success");
                }
                else
                {
                    Console.WriteLine("failed");
                }           
            }
        }
        [TestMethod]
        [Priority(13)]
        public void ID_012_TimSanPhamKhongCo()
        {
            driver.Navigate().GoToUrl("https://localhost:44366/Users/Login");
            driver.Manage().Window.Size = new System.Drawing.Size(1536, 824);
            driver.FindElement(By.Name("UserEmail"));
            driver.FindElement(By.Name("UserPassword"));
            driver.FindElement(By.XPath("/html/body/div[3]/div/div/div[2]/form/input")).SendKeys(Keys.Enter);
            
            string Suc = "Pass";
            string Failed = "Fail";
            driver.FindElement(By.Name("UserEmail")).SendKeys(xlRange.Cells[6][52].value.ToString());
            driver.FindElement(By.Name("UserPassword")).SendKeys(xlRange.Cells[6][53].value.ToString());
            driver.FindElement(By.XPath("/html/body/div[3]/div/div/div[2]/form/input")).Click();

            

            if (!driver.Url.Equals("https://localhost:44366/Users/Login"))
            {
                // Tìm phần tử input tìm kiếm và nhập vào từ khóa tìm kiếm không tồn tại
                driver.FindElement(By.Name("searchString")).Click();
                Thread.Sleep(1000);
                driver.FindElement(By.Name("searchString")).SendKeys("laptop");
                Thread.Sleep(1000);
                driver.FindElement(By.CssSelector(".bi-search")).Click();
                Thread.Sleep(1000);

                // Kiểm tra xem có thông báo không tìm thấy sản phẩm hay không
                try
                {
                    IWebElement notFoundMessage = driver.FindElement(By.CssSelector(".not-found-message"));
                    if (notFoundMessage.Displayed)
                    {
                        Console.WriteLine("Không tìm thấy sản phẩm.");
                    }
                }
                catch (NoSuchElementException)
                {
                    Console.WriteLine("Đã xảy ra lỗi khi kiểm tra tìm kiếm sản phẩm không tồn tại.");
                }

                Thread.Sleep(1000); // Đợi một lát trước khi đóng trình duyệt
            }
            if (driver.Url != "https://localhost:44366/Users/Login")
            {
                xlRange[11][42].value = Suc;
            }
            else
            {
                xlRange[11][42].value = Failed;
            }
            Thread.Sleep(2000);
        }
        // *********************TEST SELENIUM XEM LICH SU MUA HÀNG **************************************/
        [TestMethod]
        [Priority(14)]
        public void ID_013_XemLichSuMua()
        {
            driver.Navigate().GoToUrl("https://localhost:44366/Users/Login");
            driver.Manage().Window.Size = new System.Drawing.Size(1536, 824);
            driver.FindElement(By.Name("UserEmail"));
            driver.FindElement(By.Name("UserPassword"));
            driver.FindElement(By.XPath("/html/body/div[3]/div/div/div[2]/form/input")).SendKeys(Keys.Enter);
            
            string Suc = "Pass";
            string Failed = "Fail";
            driver.FindElement(By.Name("UserEmail")).SendKeys(xlRange.Cells[6][52].value.ToString());
            driver.FindElement(By.Name("UserPassword")).SendKeys(xlRange.Cells[6][53].value.ToString());
            driver.FindElement(By.XPath("/html/body/div[3]/div/div/div[2]/form/input")).Click();

            

            if (!driver.Url.Equals("https://localhost:44366/Users/Login"))
            {
                driver.FindElement(By.CssSelector("a > p:nth-child(2)")).Click();
                Thread.Sleep(1000);
                driver.FindElement(By.LinkText("Đơn hàng của bạn")).Click();
                Thread.Sleep(1000);
                driver.FindElement(By.CssSelector(".cart-item:nth-child(3) a")).Click();
                Thread.Sleep(1000);

                
            }
            if (driver.Url != "https://localhost:44366/Users/Login")
            {
                xlRange[11][43].value = Suc;
            }
            else
            {
                xlRange[11][43].value = Failed;
            }
            Thread.Sleep(2000);

        }
        // *********************TEST SELENIUM XEM THONG KE BAO CAO **************************************/
        [TestMethod]
        [Priority(15)]
        public void ID_014_XemBaoCao_ThongKe()
        {
            driver.Navigate().GoToUrl("https://localhost:44366/Users/Login");
            driver.Manage().Window.Size = new System.Drawing.Size(1536, 824);
            driver.FindElement(By.Name("UserEmail"));
            driver.FindElement(By.Name("UserPassword"));
            driver.FindElement(By.XPath("/html/body/div[3]/div/div/div[2]/form/input")).SendKeys(Keys.Enter);
           
            string Suc = " pass";
            string Failed = "failed";
            driver.FindElement(By.Name("UserEmail")).SendKeys(xlRange.Cells[6][106].value.ToString());
            driver.FindElement(By.Name("UserPassword")).SendKeys(xlRange.Cells[6][107].value.ToString());
            driver.FindElement(By.XPath("/html/body/div[3]/div/div/div[2]/form/input")).Click();

            

            if (!driver.Url.Equals("https://localhost:44366/Users/Login"))
            {
                driver.FindElement(By.CssSelector("li:nth-child(1) span:nth-child(2)")).Click();
                Thread.Sleep(1000);
                driver.FindElement(By.CssSelector("tbody:nth-child(1) > tr:nth-child(2) > td")).Click();
                Thread.Sleep(1000);
                driver.FindElement(By.CssSelector("tbody:nth-child(1)")).Click();
                Thread.Sleep(2000);
                driver.FindElement(By.LinkText("Xem chi tiết")).Click();
                driver.FindElement(By.XPath("/html/body/div[2]/div[1]/div[2]/div/a")).Click();
                Thread.Sleep(2000);
            }
            if (driver.Url != "https://localhost:44366/Users/Login")
            {
                xlRange[12][44].value = Suc;
            }
            else
            {
                xlRange[12][44].value = Failed;
            }
            Thread.Sleep(2000);

        }
        [TestMethod]
        [Priority(16)]
        public void ID_015_CapNhatSoLuongSp()
        {
            driver.Navigate().GoToUrl("https://localhost:44366/Users/Login");
            driver.Manage().Window.Size = new System.Drawing.Size(1536, 824);
            driver.FindElement(By.Name("UserEmail"));
            driver.FindElement(By.Name("UserPassword"));
            driver.FindElement(By.XPath("/html/body/div[3]/div/div/div[2]/form/input")).SendKeys(Keys.Enter);
            string Suc = " success";
            string Failed = "failed";
            driver.FindElement(By.Name("UserEmail")).SendKeys(xlRange.Cells[6][52].value.ToString());
            driver.FindElement(By.Name("UserPassword")).SendKeys(xlRange.Cells[6][53].value.ToString());
            driver.FindElement(By.XPath("/html/body/div[3]/div/div/div[2]/form/input")).Click();
            int i = 0;

            if (!driver.Url.Equals("https://localhost:44366/Users/Login"))
            {
                driver.Navigate().GoToUrl("  https://localhost:44366/");
                driver.FindElement(By.XPath("/html/body/div[3]/div[2]/div/div[2]/div[2]/div[1]/div/div[8]/div/a/div/img")).Click();
                // Tìm và click vào button để tăng, giảm số lượng sản phẩm

                driver.FindElement(By.Id("counter-plus")).Click(); // Click tăng lần thứ 1
                Thread.Sleep(1000);
                driver.FindElement(By.Id("counter-plus")).Click(); // Click tăng lần thứ 2
                Thread.Sleep(1000);
                driver.FindElement(By.Id("counter-plus")).Click(); // Click tăng lần thứ 3
                Thread.Sleep(1000);
                driver.FindElement(By.Id("counter-minus")).Click(); // Click giảm lần thứ 1
                Thread.Sleep(1000);            
            }
            if (driver.Url != "https://localhost:44366/Users/Login")
            {
                xlRange[12][45].value = Suc;
            }
            else
            {
                xlRange[12][45].value = Failed;
            }
            Thread.Sleep(2000);
        }
        [TestMethod]
        [Priority(17)]
        public void ID_016_MuaHangKhiChuaDangNhap()
        {
            driver.Navigate().GoToUrl("https://localhost:44366/");
            Thread.Sleep(1000);
            driver.FindElement(By.CssSelector(".user- > p")).Click();
            Thread.Sleep(1000);
            if (driver.Url.Equals("https://localhost:44366/Users/Login"))
            {
                Console.WriteLine("Success");
            }
            else
            {
                Console.WriteLine("failed");
            }

            if (driver.Url.Equals("https://localhost:44366/Users/Login"))
            {
                // Nếu trang hiện tại là trang đăng nhập, bạn cần thực hiện các bước đăng nhập hoặc chuyển hướng đến trang đăng nhập
                // Ví dụ: Điều hướng đến trang đăng nhập
                driver.Manage().Window.Size = new System.Drawing.Size(1536, 824);
                driver.FindElement(By.Name("UserEmail"));
                driver.FindElement(By.Name("UserPassword"));
                driver.FindElement(By.XPath("/html/body/div[3]/div/div/div[2]/form/input")).SendKeys(Keys.Enter);

                string Suc = "Pass";
                string Failed = "Fail";
                driver.FindElement(By.Name("UserEmail")).SendKeys(xlRange.Cells[6][52].value.ToString());
                driver.FindElement(By.Name("UserPassword")).SendKeys(xlRange.Cells[6][53].value.ToString());
                driver.FindElement(By.XPath("/html/body/div[3]/div/div/div[2]/form/input")).Click();

                if (driver.Url != "https://localhost:44366/Users/Login")
                {
                    xlRange[11][46].value = Suc;
                }
                else
                {
                    xlRange[11][46].value = Failed;
                }
                Thread.Sleep(2000);
                // Sau khi đăng nhập thành công, bạn có thể kiểm tra lại trang giỏ hàng hoặc tiếp tục các thao tác khác
                driver.FindElement(By.CssSelector(".user- > p")).Click();
                Thread.Sleep(1000);
            }

            driver.Close();

        }

        [TestMethod]
        // [Priority(18)]
        public void ID_017_KiemTraSoLuongDonHang()
        {
            driver.FindElement(By.Name("UserEmail"));
            driver.FindElement(By.Name("UserPassword"));
            driver.FindElement(By.XPath("/html/body/div[3]/div/div/div[2]/form/input")).SendKeys(Keys.Enter);

            string Suc = "Pass";
            string Failed = "Fail";
            // Đăng nhập Admin
            driver.FindElement(By.Name("UserEmail")).SendKeys(xlRange.Cells[6][106].value.ToString());
            driver.FindElement(By.Name("UserPassword")).SendKeys(xlRange.Cells[6][107].value.ToString());
            driver.FindElement(By.XPath("/html/body/div[3]/div/div/div[2]/form/input")).Click();

            Thread.Sleep(2000);

            IWebElement count_pro_buy = driver.FindElement(By.XPath("/html/body/div[2]/div[2]/div[1]/div[3]/div[1]/h4"));
            string count_Buy = count_pro_buy.Text;
            int count_buy = Convert.ToInt32(count_Buy);

            driver.FindElement(By.XPath("/html/body/div[2]/div[1]/div[2]/div/a")).Click();

            driver.Navigate().GoToUrl("https://localhost:44366/Users/Login");
            driver.Manage().Window.Size = new System.Drawing.Size(1536, 824);
            driver.FindElement(By.Name("UserEmail"));
            driver.FindElement(By.Name("UserPassword"));
            driver.FindElement(By.XPath("/html/body/div[3]/div/div/div[2]/form/input")).SendKeys(Keys.Enter);
           

            driver.FindElement(By.Name("UserEmail")).SendKeys(xlRange.Cells[6][52].value.ToString());
            driver.FindElement(By.Name("UserPassword")).SendKeys(xlRange.Cells[6][53].value.ToString());
            driver.FindElement(By.XPath("/html/body/div[3]/div/div/div[2]/form/input")).Click();

            Thread.Sleep(2000);

            if (!driver.Url.Equals("https://localhost:44366/Users/Login"))
            {
                driver.Navigate().GoToUrl("  https://localhost:44366/");
                driver.FindElement(By.XPath("/html/body/div[3]/div[2]/div/div[2]/div[2]/div[1]/div/div[8]/div/a/div/img")).Click();
                // Tìm và click vào button để tăng, giảm số lượng sản phẩm
                driver.FindElement(By.Id("counter-plus")).Click(); // Click tăng lần thứ 1
                Thread.Sleep(1000);
                driver.FindElement(By.Id("counter-plus")).Click(); // Click tăng lần thứ 2
                Thread.Sleep(1000);
                driver.FindElement(By.Id("counter-plus")).Click(); // Click tăng lần thứ 3
                Thread.Sleep(1000);
                driver.FindElement(By.Id("counter-minus")).Click(); // Click giảm lần thứ 1
                Thread.Sleep(1000);
                driver.FindElement(By.XPath("/html/body/div[3]/div/div[1]/div[2]/div[4]/form[1]/button")).Click();
                Thread.Sleep(1000);
                driver.FindElement(By.XPath("/html/body/div[3]/div/div/div[2]/div/form/input[2]")).Click();
                // Thử đặt hàng  khi chưa nhập dia chỉ(báo lỗi:Please fill out this field) 
                driver.FindElement(By.XPath("/html/body/div[3]/div/div[2]/form/div/div[2]/div/button[1]")).Click();
                Thread.Sleep(2000);

                driver.FindElement(By.XPath("/html/body/div[3]/div/div[2]/form/div/div[1]/table/tbody/tr[4]/td/textarea")).SendKeys("828-sư vạn hạnh-q10-tphcm");
                Thread.Sleep(1000);
                driver.FindElement(By.XPath("/html/body/div[3]/div/div[2]/form/div/div[2]/div/button[1]")).Click();
                Thread.Sleep(1000);
                driver.FindElement(By.CssSelector("a > p:nth-child(2)")).Click();
                Thread.Sleep(1000);
                driver.FindElement(By.LinkText("Đăng xuất")).Click();
                Thread.Sleep(1000);
                driver.FindElement(By.Name("UserEmail"));
                Thread.Sleep(1000);
                driver.FindElement(By.Name("UserPassword"));
                Thread.Sleep(1000);
                driver.FindElement(By.XPath("/html/body/div[3]/div/div/div[2]/form/input")).SendKeys(Keys.Enter);
                
                driver.FindElement(By.Name("UserEmail")).SendKeys(xlRange.Cells[6][106].value.ToString());
                driver.FindElement(By.Name("UserPassword")).SendKeys(xlRange.Cells[6][107].value.ToString());
                driver.FindElement(By.XPath("/html/body/div[3]/div/div/div[2]/form/input")).Click();
            }
            IWebElement count_pro_buy2 = driver.FindElement(By.XPath("/html/body/div[2]/div[2]/div[1]/div[3]/div[1]/h4"));
            string count_Buy2 = count_pro_buy2.Text;
            int count_buy2 = Convert.ToInt32(count_Buy2);
            if (count_buy2 == count_buy + 1)
            {
                Console.WriteLine("Success", count_buy2);
                xlRange[12][47].value = Suc;
            }
            else
            {
                Console.WriteLine("failse");
                xlRange[12][47].value = Failed;
            }
            Thread.Sleep(2000);
        }

    }
}
