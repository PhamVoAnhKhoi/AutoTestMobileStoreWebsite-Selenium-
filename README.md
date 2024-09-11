# Auto Test for Mobile Store Website

- Giới thiệu:
    Đây là một dự án tự động kiểm thử (auto test) dành cho một trang web bán điện thoại sử dụng Selenium và C#. Dự án này được thiết kế để kiểm tra tính năng và chất lượng của trang web một cách tự động, nhằm đảm bảo rằng các chức năng chính của trang web hoạt động đúng như mong đợi.

- Mục tiêu:
    + Kiểm thử tự động: Tự động hóa các kịch bản kiểm thử để đảm bảo tính năng của trang web hoạt động đúng.
    + Kiểm tra chức năng chính: Đảm bảo các chức năng như tìm kiếm sản phẩm, thêm sản phẩm vào giỏ hàng, và thực hiện thanh toán hoạt động chính xác.
    + Đảm bảo chất lượng: Phát hiện lỗi và vấn đề sớm để cải thiện chất lượng sản phẩm trước khi phát hành.

- Công nghệ sử dụng:
    + Ngôn ngữ lập trình: C#
    + Thư viện kiểm thử: Selenium WebDriver
    + Khung kiểm thử: NUnit (.Net Framework)

- Cài đặt:
1. **Cài đặt các gói cần thiết**: Sử dụng NuGet để cài đặt các gói Selenium WebDriver và NUnit.
   ```bash
   dotnet add package Selenium.WebDriver
   dotnet add package Microsoft.CSharp
   dotnet add package Microsoft.Office.Interop.Excel
   dotnet add package Selenium.WebDriver.ChromeDriver
   dotnet add package Selenium.WebDriver.ChromeDriver
   dotnet add package NUnit
2. **Chạy website kiểm thử**: Vì đây là project cá nhân trong quá trình mình học tập nên trước khi chạy cần lưu ý như sau:
     - Download project website về máy
     - Connect database với SQL server
     - Điều chỉnh lại đường dẫn đặt file excel   
