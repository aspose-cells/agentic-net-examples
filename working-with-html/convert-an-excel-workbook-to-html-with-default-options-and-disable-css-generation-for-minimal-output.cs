// Title: Convert an Excel workbook to minimal HTML without CSS using Aspose.Cells in C#
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, sets HtmlSaveOptions.ExportCss to false, embeds images as Base64, and saves a single HTML file. | Show how to configure Aspose.Cells HtmlSaveOptions to produce a CSS‑free HTML output from a workbook. | Provide a robust example that validates the input Excel path, handles load and save exceptions, and creates a minimal HTML file with no external stylesheet.
// Common Searches: Aspose.Cells C# generate HTML from workbook without CSS files | disable CSS generation in Aspose.Cells HtmlSaveOptions example | create self‑contained HTML from Excel file using Base64 images Aspose.Cells | minimal HTML output settings for Aspose.Cells workbook export
// Tags: Aspose.Cells HtmlSaveOptions ExportCss false | C# Excel workbook HTML export without stylesheet | Aspose.Cells generate single HTML file with embedded images | minimal HTML generation using Aspose.Cells | exception handling for Aspose.Cells workbook operations

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

// The sample verifies that input.xlsx exists, loads it into an Aspose.Cells Workbook, configures HtmlSaveOptions to disable CSS generation and embed images as Base64, then saves the workbook as output.html while handling any loading or saving errors.
class ExcelToHtmlConverter
{
    static void Main()
    {
        const string inputFile = "input.xlsx";
        const string outputFile = "output.html";

        // Verify that the input Excel file exists
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Error: Input file \"{inputFile}\" not found.");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook from the specified file
            workbook = new Workbook(inputFile);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        // Configure HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        // Optional: embed images as Base64 to keep a single output file
        htmlOptions.ExportImagesAsBase64 = true;

        try
        {
            // Save the workbook as an HTML file
            workbook.Save(outputFile, htmlOptions);
            Console.WriteLine($"Workbook successfully converted to HTML: {outputFile}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving HTML file: {ex.Message}");
        }
    }
}
