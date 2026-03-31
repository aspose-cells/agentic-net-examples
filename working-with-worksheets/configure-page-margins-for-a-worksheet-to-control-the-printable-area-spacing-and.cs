using System;
using Aspose.Cells;

namespace AsposeCellsMarginDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the PageSetup object for the worksheet
            PageSetup pageSetup = worksheet.PageSetup;

            // Set margins (values are in centimeters)
            pageSetup.TopMargin = 2.0;      // 2 cm top margin
            pageSetup.BottomMargin = 1.5;   // 1.5 cm bottom margin
            pageSetup.LeftMargin = 1.0;     // 1 cm left margin
            pageSetup.RightMargin = 1.0;    // 1 cm right margin

            // Optionally set margins in inches (overwrites the cm values if used)
            // pageSetup.TopMarginInch = 0.79; // 2 cm ≈ 0.79 inches
            // pageSetup.BottomMarginInch = 0.59; // 1.5 cm ≈ 0.59 inches
            // pageSetup.LeftMarginInch = 0.39; // 1 cm ≈ 0.39 inches
            // pageSetup.RightMarginInch = 0.39; // 1 cm ≈ 0.39 inches

            // Additional page setup settings (optional)
            pageSetup.PaperSize = PaperSizeType.PaperA4;
            pageSetup.Orientation = PageOrientationType.Portrait;
            pageSetup.CenterHorizontally = true;
            pageSetup.CenterVertically = true;

            // Save the workbook (lifecycle rule: save)
            workbook.Save("WorksheetWithMargins.xlsx");

            Console.WriteLine("Workbook saved with custom page margins.");
        }
    }
}