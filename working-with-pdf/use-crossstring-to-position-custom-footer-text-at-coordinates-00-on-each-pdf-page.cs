using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add some sample data so the sheet is not empty
        sheet.Cells["A1"].PutValue("Sample content for PDF");

        // Access the PageSetup object to configure the footer
        PageSetup pageSetup = sheet.PageSetup;

        // Remove any distance between the footer and the bottom edge of the page
        // FooterMargin is in centimeters; setting it to 0 places the footer at the very bottom (0,0)
        pageSetup.FooterMargin = 0;

        // Optionally remove the bottom margin to ensure the footer touches the page edge
        pageSetup.BottomMargin = 0;

        // Set the left section of the footer with the desired custom text
        // The left section starts at the leftmost position of the footer area
        pageSetup.SetFooter(0, "Custom Footer Text");

        // Save the workbook as a PDF; the footer will appear at coordinates (0,0) on each page
        workbook.Save("CustomFooter.pdf", SaveFormat.Pdf);
    }
}