using System;
using Aspose.Cells;

namespace AsposeCellsMarginExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Set custom page margins in inches
            worksheet.PageSetup.TopMarginInch = 0.5;    // Top margin = 0.5 inches
            worksheet.PageSetup.BottomMarginInch = 0.5; // Bottom margin = 0.5 inches
            worksheet.PageSetup.LeftMarginInch = 0.3;   // Left margin = 0.3 inches
            worksheet.PageSetup.RightMarginInch = 0.3;  // Right margin = 0.3 inches

            // Optional: add some data to visualize the margins when printed
            worksheet.Cells["A1"].PutValue("Margin demonstration");
            worksheet.Cells["A2"].PutValue("Top: 0.5\", Bottom: 0.5\", Left: 0.3\", Right: 0.3\"");

            // Save the workbook to a file
            workbook.Save("CustomMargins.xlsx");
        }
    }
}