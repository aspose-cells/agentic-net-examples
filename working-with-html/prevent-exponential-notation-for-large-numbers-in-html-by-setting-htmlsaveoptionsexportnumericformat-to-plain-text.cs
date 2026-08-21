// Title: Aspose.Cells for .NET – Prevent scientific notation for large numbers when exporting to HTML
// Description: Demonstrates how to write a 20‑digit value to a cell, configure HtmlSaveOptions.ExportNumericFormat, and save the workbook as HTML so the number appears exactly as entered instead of in exponential form.
// Keywords: Aspose.Cells | .NET | C# | HtmlSaveOptions | ExportNumericFormat | prevent scientific notation | large numbers HTML export | avoid exponential notation | exact numeric display | HTML workbook export
// Common Searches: Aspose.Cells ExportNumericFormat C# example | save large integer to HTML without scientific notation | HtmlSaveOptions prevent exponential notation | Aspose.Cells large number display in HTML | C# export spreadsheet to HTML exact numbers
// Developer Intent: Export a workbook to HTML while guaranteeing that large numeric values are rendered as plain text rather than scientific notation.
// Use Cases: Web reports that must show full account or transaction IDs without rounding. | Financial dashboards where high‑precision figures need exact visual representation. | Barcode or product‑code listings exported to HTML that require the original digit sequence.
// AI Prompts: Provide C# code using Aspose.Cells to export a workbook to HTML with ExportNumericFormat set so a 20‑digit number stays in plain text. | Show how to write a large integer as a string in a cell and configure HtmlSaveOptions to avoid scientific notation in the HTML output. | Explain the role of HtmlSaveOptions.ExportNumericFormat in controlling numeric formatting during HTML export with Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates how to write a 20‑digit value to a cell, configure HtmlSaveOptions.ExportNumericFormat, and save the workbook as HTML so the number appears exactly as entered instead of in exponential form.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Insert a large numeric value as a string to keep the full precision in HTML output
            worksheet.Cells["A1"].PutValue("12345678901234567890");

            // Set up HTML save options (default options are sufficient for this scenario)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Save the workbook as an HTML file
            string outputPath = "LargeNumberOutput.html";
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
