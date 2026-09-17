// Title: Convert an Excel workbook to a single HTML file while preserving conditional formatting colors using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an .xlsx file, enables HtmlSaveOptions.ExportConditionalFormatting and ExportSingleFile, embeds images as Base64, and saves the workbook as one HTML document. | Show how to detect and set the ExportConditionalFormatting and ExportSingleFile properties at runtime to ensure conditional formatting colors appear correctly in the generated HTML.
// Common Searches: how to keep conditional formatting colors when exporting Excel to HTML with Aspose.Cells C# | Aspose.Cells HtmlSaveOptions ExportSingleFile example | C# convert .xlsx to single HTML file preserving conditional formatting | export conditional formatting icons as base64 in Aspose.Cells HTML output | save workbook as HTML with conditional formatting using Aspose.Cells .NET
// Tags: Aspose.Cells HtmlSaveOptions ExportConditionalFormatting | Aspose.Cells ExportSingleFile single HTML output | C# embed images as Base64 Aspose.Cells HTML | Excel conditional formatting colors HTML export | Aspose.Cells HTML conversion with conditional formatting

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample loads an input.xlsx workbook, configures HtmlSaveOptions to embed images as Base64, dynamically enables ExportConditionalFormatting and ExportSingleFile when available, and saves the workbook as a single output.html file that retains the original conditional formatting colors.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                // Export images (including conditional formatting icons) as Base64 strings
                ExportImagesAsBase64 = true
            };

            // Set ExportConditionalFormatting if the property exists in the current version
            var exportCfProp = typeof(HtmlSaveOptions).GetProperty("ExportConditionalFormatting");
            if (exportCfProp != null && exportCfProp.CanWrite)
            {
                exportCfProp.SetValue(htmlOptions, true);
            }

            // Set ExportSingleFile if the property exists in the current version
            var exportSingleProp = typeof(HtmlSaveOptions).GetProperty("ExportSingleFile");
            if (exportSingleProp != null && exportSingleProp.CanWrite)
            {
                exportSingleProp.SetValue(htmlOptions, true);
            }

            // Save the workbook as HTML with the specified options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
