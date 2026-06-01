using System;
using Aspose.Cells;

class VerifyIgnoreErrorExport
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Insert a formula that will generate a division by zero error (#DIV/0!)
        worksheet.Cells["A1"].Formula = "=1/0";

        // Force calculation so the error is realized in the cell
        workbook.CalculateFormula();

        // Configure PDF save options to ignore errors during rendering
        PdfSaveOptions saveOptions = new PdfSaveOptions();
        saveOptions.IgnoreError = true; // Hide errors while exporting

        // Attempt to export the workbook; no exception should be thrown
        try
        {
            workbook.Save("ExportWithIgnoreError.pdf", saveOptions);
            Console.WriteLine("Export succeeded without throwing an exception.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Export failed with exception: " + ex.Message);
        }
    }
}