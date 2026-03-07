using System;
using Aspose.Cells;

namespace AsposeCellsPrintAreaDemo
{
    class ClearPrintAreaExample
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Set a print area initially (for demonstration)
            sheet.PageSetup.PrintArea = "A1:C10";

            // Clear the print area by assigning an empty string
            sheet.PageSetup.PrintArea = string.Empty;

            // Save the workbook (lifecycle: save)
            workbook.Save("ClearPrintAreaDemo.xlsx");

            Console.WriteLine("Print area cleared and workbook saved.");
        }
    }
}