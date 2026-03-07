using System;
using Aspose.Cells;

class ClearPrintAreaDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // (Optional) Set a print area to demonstrate clearing
        worksheet.PageSetup.PrintArea = "A1:C10";

        // Clear the print area by assigning an empty string
        worksheet.PageSetup.PrintArea = string.Empty;

        // Save the workbook
        workbook.Save("ClearPrintAreaDemo.xlsx");
    }
}