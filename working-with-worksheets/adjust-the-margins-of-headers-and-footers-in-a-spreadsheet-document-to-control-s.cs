using System;
using Aspose.Cells;

namespace AdjustHeaderFooterMargins
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Access the PageSetup object for margin configuration
            PageSetup pageSetup = worksheet.PageSetup;

            // Set header margin (distance from top of page to header) in centimeters
            pageSetup.HeaderMargin = 2.0; // 2 cm

            // Set footer margin (distance from bottom of page to footer) in centimeters
            pageSetup.FooterMargin = 1.5; // 1.5 cm

            // Optionally align header/footer margins with page margins
            pageSetup.IsHFAlignMargins = true;

            // Add sample content to visualize the margins
            worksheet.Cells["A1"].PutValue("Header/Footer Margin Demo");
            worksheet.Cells["A2"].PutValue("Header margin set to 2 cm, Footer margin set to 1.5 cm.");

            // Set header and footer text so the margins are noticeable when printed
            pageSetup.SetHeader(0, "Left Header");
            pageSetup.SetHeader(1, "Center Header");
            pageSetup.SetHeader(2, "Right Header");

            pageSetup.SetFooter(0, "Left Footer");
            pageSetup.SetFooter(1, "Page &P of &N");
            pageSetup.SetFooter(2, "Right Footer");

            // Save the workbook to an Excel file
            workbook.Save("HeaderFooterMarginsDemo.xlsx");
        }
    }
}