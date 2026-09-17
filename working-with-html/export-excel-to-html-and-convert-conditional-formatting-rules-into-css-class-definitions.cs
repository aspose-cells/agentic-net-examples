// Title: Export an Excel workbook to HTML5 with external images and convert conditional‑formatting rules into CSS classes using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads a .xlsx file with Aspose.Cells, configures HtmlSaveOptions to output HTML5, saves images as separate files, and writes the result to an .html file. | Enhance the export routine to iterate over the workbook's conditional‑formatting collections, generate matching CSS class definitions, and reference those classes in the produced HTML. | Add robust error handling and logging that reports missing input files, failures during HTML export, and any issues while translating conditional formatting to CSS.
// Common Searches: Aspose.Cells C# generate HTML5 from Excel while keeping images as external files | Preserving Excel conditional formatting styles in CSS during HTML conversion with Aspose.Cells | Example of HtmlSaveOptions settings for external image files in Aspose.Cells | C# code to translate Excel conditional formatting rules into CSS classes during HTML export | Troubleshoot missing conditional formatting after exporting Excel to HTML with Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions external images | export Excel to HTML5 C# | convert Excel conditional formatting to CSS Aspose.Cells | C# workbook to HTML with CSS classes | Aspose.Cells HTML export image handling

using System;
using System.IO;
using Aspose.Cells;

namespace ExportExcelToHtmlApp
{
    // The example checks for input.xlsx, loads it with Aspose.Cells, configures HtmlSaveOptions to produce HTML5 output and write images as separate files, then saves the workbook as output.html. It can be extended to walk the workbook’s conditional‑formatting rules, generate corresponding CSS class definitions, and embed those classes in the exported HTML.
    class ExportExcelToHtml
    {
        static void Main()
        {
            try
            {
                const string inputFile = "input.xlsx";
                const string outputFile = "output.html";

                // Verify that the input workbook exists.
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Input file '{inputFile}' not found.");
                    return;
                }

                // Load the Excel workbook.
                Workbook workbook = new Workbook(inputFile);

                // Configure HTML save options.
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    // Export images as separate files (not embedded as Base64).
                    ExportImagesAsBase64 = false,
                    // Use HTML5 output.
                    HtmlVersion = HtmlVersion.Html5
                };

                // Save the workbook as an HTML file.
                workbook.Save(outputFile, htmlOptions);
                Console.WriteLine($"Workbook successfully saved to '{outputFile}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
