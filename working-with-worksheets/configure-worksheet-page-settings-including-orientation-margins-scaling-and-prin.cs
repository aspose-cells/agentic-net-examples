using System;
using Aspose.Cells;

class WorksheetPageSetupDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Get the PageSetup object for the worksheet
        PageSetup pageSetup = sheet.PageSetup;

        // Define the range that will be printed
        pageSetup.PrintArea = "A1:G30";

        // Repeat the first two rows and first two columns on each printed page
        pageSetup.PrintTitleRows = "$1:$2";
        pageSetup.PrintTitleColumns = "$A:$B";

        // Set page orientation to Landscape
        pageSetup.Orientation = PageOrientationType.Landscape;

        // Choose A4 paper size
        pageSetup.PaperSize = PaperSizeType.PaperA4;

        // Set margins in centimeters
        pageSetup.TopMargin = 1.5;      // top margin
        pageSetup.BottomMargin = 1.5;   // bottom margin
        pageSetup.LeftMargin = 2.0;     // left margin
        pageSetup.RightMargin = 2.0;    // right margin

        // Optionally set the same margins in inches (1 cm ≈ 0.3937 inch)
        pageSetup.TopMarginInch = 0.59;
        pageSetup.BottomMarginInch = 0.59;
        pageSetup.LeftMarginInch = 0.79;
        pageSetup.RightMarginInch = 0.79;

        // Center the printed sheet horizontally and vertically
        pageSetup.CenterHorizontally = true;
        pageSetup.CenterVertically = true;

        // Scale the worksheet to fit one page wide; height adjusts automatically
        pageSetup.FitToPagesWide = 1;
        pageSetup.FitToPagesTall = 0; // 0 means all rows on as many pages as needed

        // Alternatively, you could set a fixed zoom percentage (e.g., 80%)
        // pageSetup.Zoom = 80;

        // Additional print options
        pageSetup.PrintGridlines = true;   // print cell gridlines
        pageSetup.PrintHeadings = false;   // do not print row/column headings
        pageSetup.BlackAndWhite = false;   // print in color

        // Save the workbook with the configured page settings
        workbook.Save("WorksheetPageSetupDemo.xlsx");
    }
}