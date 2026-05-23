using System;
using Aspose.Cells;

namespace WorkbookToPdfWithMargins
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add some sample data (optional, just to have content)
            sheet.Cells["A1"].PutValue("Demo of PDF conversion with custom margins");
            sheet.Cells["A2"].PutValue("All margins are set to 0.5 inches");

            // Set page margins to 0.5 inches on each side
            sheet.PageSetup.LeftMarginInch = 0.5;
            sheet.PageSetup.RightMarginInch = 0.5;
            sheet.PageSetup.TopMarginInch = 0.5;
            sheet.PageSetup.BottomMarginInch = 0.5;

            // Save the workbook as PDF
            string outputPath = "WorkbookWithMargins.pdf";
            workbook.Save(outputPath, SaveFormat.Pdf);

            Console.WriteLine($"Workbook saved to PDF with 0.5 inch margins: {outputPath}");
        }
    }
}