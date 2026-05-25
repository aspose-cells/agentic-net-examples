using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Enable QuotePrefixToStyle so that a leading apostrophe is stored as a style flag
        workbook.Settings.QuotePrefixToStyle = true;

        // Add a value that starts with a single quote (apostrophe)
        var cell = workbook.Worksheets[0].Cells["A1"];
        cell.PutValue("'Leading apostrophe text");

        // Verify that the QuotePrefix style is applied
        Console.WriteLine("QuotePrefix style applied: " + cell.GetStyle().QuotePrefix);
        Console.WriteLine("Cell displayed value: " + cell.StringValue);

        // Configure PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Ensure the workbook's default font is checked for Unicode characters
            CheckWorkbookDefaultFont = true
        };

        // Save the workbook to PDF
        workbook.Save("output.pdf", pdfOptions);

        Console.WriteLine("Workbook successfully saved to PDF with leading apostrophe displayed.");
    }
}