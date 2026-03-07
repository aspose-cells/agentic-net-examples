using System;
using Aspose.Cells;

namespace HeaderFooterMarginsDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Get the PageSetup object to configure margins and header/footer
            PageSetup pageSetup = sheet.PageSetup;

            // Set header and footer margins (in centimeters)
            pageSetup.HeaderMargin = 2.0;   // 2 cm from the top of the page
            pageSetup.FooterMargin = 1.5;   // 1.5 cm from the bottom of the page

            // Align header/footer with the page margins (optional, default is true)
            pageSetup.IsHFAlignMargins = true;

            // Add some sample data to visualize the margins
            sheet.Cells["A1"].PutValue("Header/Footer Margin Demo");
            sheet.Cells["A2"].PutValue("Header margin = 2 cm, Footer margin = 1.5 cm");

            // Define header text for left, center, and right sections
            pageSetup.SetHeader(0, "&\"Arial,Bold\"&12Header Left");   // Left section
            pageSetup.SetHeader(1, "&\"Arial\"&10Header Center");     // Center section
            pageSetup.SetHeader(2, "&\"Arial\"&10Header Right");      // Right section

            // Define footer text for left, center, and right sections
            pageSetup.SetFooter(0, "&\"Arial\"&10Footer Left");       // Left section
            pageSetup.SetFooter(1, "Page &P of &N");                  // Center section (page number)
            pageSetup.SetFooter(2, "&D");                            // Right section (date)

            // Save the workbook to an Excel file
            workbook.Save("HeaderFooterMarginsDemo.xlsx");
        }
    }
}