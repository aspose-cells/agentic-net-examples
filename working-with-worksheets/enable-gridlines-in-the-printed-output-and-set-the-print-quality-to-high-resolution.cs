using System;
using Aspose.Cells;

namespace AsposeCellsPrintDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add some sample data
            worksheet.Cells["A1"].PutValue("Print Gridlines Demo");
            worksheet.Cells["A2"].PutValue("This sheet will print with gridlines and high quality.");

            // Enable gridlines in the printed output
            worksheet.PageSetup.PrintGridlines = true;

            // Set a high print quality (e.g., 300 DPI)
            worksheet.PageSetup.PrintQuality = 300;

            // Save the workbook
            workbook.Save("PrintGridlinesHighQuality.xlsx");
        }
    }
}