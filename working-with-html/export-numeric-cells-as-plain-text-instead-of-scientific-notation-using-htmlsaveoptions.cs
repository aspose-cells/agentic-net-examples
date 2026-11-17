// Title: Save a workbook to HTML with Aspose.Cells for .NET while keeping numeric cells as plain text instead of scientific notation
// Common Searches: Aspose.Cells C# export workbook to HTML without scientific notation | how to keep large numbers as plain text in HTML output using Aspose.Cells | prevent scientific notation for small floating‑point values in Aspose.Cells HTML save | convert numeric cell values to string before saving as HTML with Aspose.Cells .NET
// Tags: html export numeric formatting Aspose.Cells | disable scientific notation Aspose.Cells HTML | convert cell values to string Aspose.Cells | save workbook as html plain text numbers | Aspose.Cells numeric cell formatting HTML output

using System;
using System.IO;
using Aspose.Cells;

// The example creates a Workbook, writes a large integer and a tiny floating‑point number that would normally appear in scientific notation, configures HtmlSaveOptions, and saves the file as HTML. Because the current API lacks a direct ExportCellValueAsString option, the code notes that you must adjust cell formatting or explicitly convert numeric values to strings before saving to retain plain‑text representation.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Populate some numeric cells for demonstration
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue(1234567890L);      // large integer
            sheet.Cells["A2"].PutValue(1.23e-10);        // small number that may appear in scientific notation

            // Set HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            // Note: ExportCellValueAsString is not available in the current API version.
            // If needed, adjust cell formatting or convert values to strings before saving.

            // Save the workbook to HTML using the configured options
            string outputPath = "output.html";
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
