using System;
using Aspose.Cells;

class RemovePrinterSettings
{
    static void Main()
    {
        // Load the existing Excel file
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);   // uses the Workbook(string) constructor

        // Clear printer settings for each worksheet
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // The PrinterSettings property holds a byte array; setting it to null removes any existing settings
            sheet.PageSetup.PrinterSettings = null;
        }

        // Save the workbook with the printer settings removed
        string outputPath = "output.xlsx";
        workbook.Save(outputPath);                    // uses the Save(string) method
    }
}