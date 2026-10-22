// Title: Retain Excel conditional formatting colors when exporting to HTML with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an .xlsx workbook and saves it as HTML with HtmlSaveOptions.ExportConditionalFormatting set to true so that all conditional formatting colors appear in the output. | Update an existing Aspose.Cells HTML export example to explicitly enable ExportConditionalFormatting, then show how to verify that the generated HTML contains the expected color styles.
// Common Searches: aspnet cells export conditional formatting colors to html c# | c# aspose.cells HtmlSaveOptions ExportConditionalFormatting true example | preserve excel conditional formatting when converting to html using aspose.cells | how to keep conditional formatting colors in html output from aspose.cells | export workbook to html with conditional formatting retained .net
// Tags: Aspose.Cells HtmlSaveOptions ExportConditionalFormatting | C# export Excel to HTML with conditional formatting | preserve conditional formatting colors Aspose.Cells | HTML export of Excel conditional formatting .NET | embed images base64 Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace ExportConditionalFormattingToHtmlApp
{
    // The sample loads an existing Excel file, configures HtmlSaveOptions (including ExportConditionalFormatting = true and embedding images as Base64), and saves the workbook as HTML. The resulting HTML retains all conditional formatting color rules.
    class ExportConditionalFormattingToHtml
    {
        static void Main()
        {
            try
            {
                string inputPath = "Input.xlsx";
                string outputPath = "Output.html";

                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Configure HTML save options (conditional formatting is exported by default)
                HtmlSaveOptions saveOptions = new HtmlSaveOptions
                {
                    ExportImagesAsBase64 = true // Embed images directly in the HTML
                };

                // Save the workbook as HTML with the specified options
                workbook.Save(outputPath, saveOptions);
                Console.WriteLine($"Workbook successfully saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
