using System;
using Aspose.Cells;

class PageMarginsDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add some sample data to visualize the margins
        sheet.Cells["A1"].PutValue("Demonstrating page margins");
        sheet.Cells["A2"].PutValue("Top, Bottom, Left, Right margins are set.");

        // Get the PageSetup object for the worksheet
        PageSetup pageSetup = sheet.PageSetup;

        // Set margins in centimeters
        pageSetup.TopMargin = 2.0;      // 2 cm top margin
        pageSetup.BottomMargin = 1.5;   // 1.5 cm bottom margin
        pageSetup.LeftMargin = 1.0;     // 1 cm left margin
        pageSetup.RightMargin = 1.0;    // 1 cm right margin

        // Optionally set the same margins in inches (overwrites the previous values)
        pageSetup.TopMarginInch = 0.8;      // ~2 cm
        pageSetup.BottomMarginInch = 0.6;   // ~1.5 cm
        pageSetup.LeftMarginInch = 0.4;     // ~1 cm
        pageSetup.RightMarginInch = 0.4;    // ~1 cm

        // Save the workbook as PDF to see the effect of the margins
        workbook.Save("PageMarginsDemo.pdf", SaveFormat.Pdf);
    }
}