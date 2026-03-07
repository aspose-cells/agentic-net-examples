using Aspose.Cells;
using System;

class ClearPrintAreaDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Set a print area (for demonstration purposes)
        sheet.PageSetup.PrintArea = "A1:C10";

        // Clear the print area by assigning an empty string
        sheet.PageSetup.PrintArea = string.Empty;

        // Save the workbook
        workbook.Save("ClearPrintAreaDemo.xlsx");
    }
}