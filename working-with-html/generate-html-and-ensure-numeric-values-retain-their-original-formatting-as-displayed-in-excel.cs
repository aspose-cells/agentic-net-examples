// Title: Generate a self‑contained HTML file from an Excel workbook in C# while keeping original numeric cell formatting using Aspose.Cells
// AI Prompts: Write C# code that loads an .xlsx file, sets HtmlSaveOptions to embed worksheet images as Base64 strings and export all worksheets, then saves the workbook as a single HTML document preserving the Excel number formats. | Adjust the Aspose.Cells HTML export to output only the active worksheet and disable Base64 image embedding, while still retaining the original numeric formatting of cells. | Add a custom number‑format rule to HtmlSaveOptions so that every numeric cell appears with exactly two decimal places in the generated HTML.
// Common Searches: how to export an entire Excel workbook to a single HTML file with embedded images using Aspose.Cells .NET | preserve Excel number formatting when converting .xlsx to HTML in C# | Aspose.Cells HtmlSaveOptions export all worksheets and keep numeric formats | C# generate self‑contained HTML from Excel with Base64 images and original cell formatting | convert Excel to HTML while maintaining decimal precision using Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions all worksheets | C# self-contained HTML from Excel | Base64 image embedding in HTML export | numeric formatting preservation in HTML conversion | Excel to HTML conversion with original number formats

using System;
using System.IO;
using Aspose.Cells;

// The example loads an input.xlsx file with Aspose.Cells, configures HtmlSaveOptions to export all worksheets and embed images as Base64, and saves the workbook as output.html. The HTML output retains the exact numeric formatting displayed in the original Excel workbook.
class ExcelToHtml
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Ensure the input Excel file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the file
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Export all worksheets (not only the active one)
                ExportActiveWorksheetOnly = false,
                // Embed images as Base64 to keep HTML self‑contained
                ExportImagesAsBase64 = true
                // Numeric formatting is preserved by default in HTML export
            };

            // Save the workbook as an HTML file using the configured options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
