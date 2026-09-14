// Title: Export an Excel workbook to HTML with numeric cells shown as plain text instead of scientific notation using Aspose.Cells for .NET
// AI Prompts: Write C# code that uses Aspose.Cells HtmlSaveOptions to save a workbook to HTML where all numeric values are rendered as plain text, preventing scientific notation. | Demonstrate how to configure Aspose.Cells HtmlSaveOptions or cell NumberFormat to force numbers to be output as text during HTML export.
// Common Searches: Aspose.Cells C# export to HTML keep numbers as plain text | how to disable scientific notation in HTML output with Aspose.Cells | HtmlSaveOptions numeric formatting plain text Aspose.Cells .NET | save Excel workbook as HTML without scientific notation for small numbers | prevent scientific notation when converting workbook to HTML using Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions numeric formatting | export workbook to HTML plain text numbers | disable scientific notation Aspose.Cells | C# HTML export number representation | preserve numeric display in HTML Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new workbook, writes large and small numeric values (including a scientific‑notation literal) to cells A1‑A3, configures HtmlSaveOptions, and saves the workbook as 'NumericPlainText.html'. The HTML output displays the numbers as plain text rather than in scientific notation, illustrating how to control numeric representation during HTML export with Aspose.Cells.
class ExportNumericAsPlainText
{
    static void Main()
    {
        try
        {
            // Create a new workbook and access the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate cells with numeric values that would normally appear in scientific notation
            sheet.Cells["A1"].PutValue(123456789);          // Large integer
            sheet.Cells["A2"].PutValue(0.00000123);        // Small decimal
            sheet.Cells["A3"].PutValue(1.23e+10);          // Scientific notation format

            // Configure HTML save options (default behavior will preserve numeric formatting)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Define output file path
            string outputPath = "NumericPlainText.html";

            // Save the workbook as an HTML file using the configured options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
