using System;
using Aspose.Cells;

class ClearPrintAreaDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Set a print area (for demonstration)
        worksheet.PageSetup.PrintArea = "A1:C10";

        // Verify that the print area is set
        Console.WriteLine("Print area before clearing: " + worksheet.PageSetup.PrintArea);

        // Clear the print area by assigning an empty string
        worksheet.PageSetup.PrintArea = string.Empty;

        // Verify that the print area has been cleared
        Console.WriteLine("Print area after clearing: " + (string.IsNullOrEmpty(worksheet.PageSetup.PrintArea) ? "(cleared)" : worksheet.PageSetup.PrintArea));

        // Save the workbook (the file will not contain any print area definition)
        workbook.Save("ClearPrintAreaDemo.xlsx");

        // Dispose the workbook (optional but good practice)
        workbook.Dispose();
    }
}