using System;
using Aspose.Cells;

namespace AsposeCellsPrintAreaRemoval
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // creates a blank workbook

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Example: set a print area first (to demonstrate removal)
            worksheet.PageSetup.PrintArea = "A1:C10";

            // Remove any defined print area by setting it to an empty string
            // This resets the worksheet to use the default printing behavior (entire sheet)
            worksheet.PageSetup.PrintArea = string.Empty;

            // Save the workbook to verify that the print area has been cleared
            workbook.Save("PrintAreaRemoved.xlsx", SaveFormat.Xlsx);
        }
    }
}