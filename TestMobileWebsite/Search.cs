using Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;

namespace UnitTestProject1
{
    [TestClass]
    public class Search
    {
        IWebDriver driver;
        IWebElement searchBar, btnSearch, searchResultElement;
        IWebElement nameProduct;

        Excel.Application dataApp;
        Excel.Workbook dataWorkBook;
        Excel.Worksheet dataWorksheet;
        Excel.Range xlRange;


        string dataSearch;

        [TestInitialize]
        public void OpenWebsite()
        {
            driver = new ChromeDriver();
            driver.Url = "https://localhost:44366/";
            driver.Navigate();
            dataApp = new Excel.Application();
            dataWorkBook = dataApp.Workbooks.Open(@"D:\\Learn\\DBCLPM\\DoAn_DBCLPM\\DoAn_DBCLPM\\Final_TestCase.xlsx");
            dataWorksheet = dataWorkBook.Sheets[1];
            xlRange = dataWorksheet.UsedRange;
        }
        [TestCleanup]
        public void Cleanup()
        {
            dataWorkBook.Save();
            dataWorkBook.Close();
            dataApp.Quit();
            driver.Quit();
        }
        public void prepareToSearch(int value)
        {
           

        }

        // Search không nhập thông tin
        [TestMethod]
        public void TestMethod1()
        {
            searchBar = driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[2]/form/input"));
            searchBar.Click();

            btnSearch = driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[2]/form/button"));
            btnSearch.Click();
            searchResultElement = driver.FindElement(By.ClassName("card"));
            
            if (!searchResultElement.Displayed)
            {
                Console.WriteLine("Success");
            }
            else
            {
                Console.WriteLine("Fail");
            }
        }
        //Tìm kiếm sản phẩm theo tên 
        [TestMethod]
        public void SearchNamePro()
        {
            searchBar = driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[2]/form/input"));
            searchBar.Click();
            dataSearch = xlRange.Cells[6][225].value.ToString();
            searchBar.SendKeys(dataSearch);
            
            //dataSearch = "Xiaomi Redmi Note 11 Pro 5G 128GB";
            

            btnSearch = driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[2]/form/button"));
            btnSearch.Click();

            searchResultElement = driver.FindElement(By.ClassName("card"));

            if (searchResultElement.Displayed)
            {
                Console.WriteLine("Search success");
                nameProduct = driver.FindElement(By.ClassName("card-text"));
                string expectedResult = nameProduct.Text;
                if(expectedResult == dataSearch)
                {
                    Console.WriteLine("Search true product");
                }
                else
                {
                    Console.WriteLine("Search fail product");
                }
            }
            else
            {
                Console.WriteLine("Fail");
            }
            Thread.Sleep(5000);
        }

        //Tìm kiếm sản phẩm theo thương hiệu 
        [TestMethod]
        public void SearchCatePro()
        {
            searchBar = driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[2]/form/input"));
            searchBar.Click();
            dataSearch = xlRange.Cells[6][226].value.ToString();
            searchBar.SendKeys(dataSearch);


            btnSearch = driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[2]/form/button"));
            btnSearch.Click();

            searchResultElement = driver.FindElement(By.ClassName("card"));

            if (searchResultElement.Displayed)
            {
                Console.WriteLine("Search success");
                nameProduct = driver.FindElement(By.ClassName("card-text"));
                string expectedResult = nameProduct.Text;
                if (expectedResult == dataSearch)
                {
                    Console.WriteLine("Search true product");
                }
                else
                {
                    Console.WriteLine("Search fail product");
                }
            }
            else
            {
                Console.WriteLine("Fail");
            }
            Thread.Sleep(5000);
        }

    }
}
