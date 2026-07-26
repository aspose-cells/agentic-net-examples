// Title: Export Excel to HTML with Inline Styles Only (DisableCss) – Aspose.Cells for .NET
// Description: Demonstrates how to save a workbook as HTML using Aspose.Cells while suppressing external CSS. The HtmlSaveOptions.DisableCss flag forces all formatting to be written as inline style attributes, producing a single‑file HTML output.
// Keywords: Aspose.Cells HTML export | DisableCss | inline styles | C# Excel to HTML | no external stylesheet | HtmlSaveOptions
// Common Searches: Aspose.Cells disable CSS when exporting to HTML | HTML export with inline styles only .NET | How to turn off CSS generation in Aspose.Cells | Save Excel as HTML without external CSS file | HtmlSaveOptions.DisableCss example
// Developer Intent: Create an HTML representation of an Excel workbook where all cell formatting is embedded directly in the markup, eliminating the need for a separate CSS file.
// Use Cases: Embedding workbook data in email bodies that prohibit external style sheets. | Generating lightweight, self‑contained HTML reports for quick preview in browsers. | Packaging documentation where managing additional CSS resources is undesirable.
// AI Prompts: Write C# code that loads an existing .xlsx file, sets HtmlSaveOptions.DisableCss to true, and saves it as a single HTML file with a custom name. | Explain how to combine DisableCss with ExportImagesAsBase64 and PageSetup options for a fully self‑contained HTML export. | Provide a step‑by‑step guide to convert each worksheet in a workbook to separate HTML files using only inline styles.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to save a workbook as HTML using Aspose.Cells while suppressing external CSS. The HtmlSaveOptions.DisableCss flag forces all formatting to be written as inline style attributes, producing a single‑file HTML output.
    public class DisableCssHtmlExport
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (in-memory)
                Workbook workbook = new Workbook();

                // Access the first worksheet and add some sample data
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Hello");
                sheet.Cells["B1"].PutValue("World");
                sheet.Cells["A1"].GetStyle().Font.IsBold = true;
                sheet.Cells["B1"].GetStyle().Font.Color = Color.Blue;

                // Create HtmlSaveOptions and disable CSS generation (use inline styles only)
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    DisableCss = true // Inline styles will be applied, no external CSS file
                };

                // Save the workbook as HTML with the specified options
                string outputPath = "HtmlWithInlineStyles.html";
                workbook.Save(outputPath, htmlOptions);

                Console.WriteLine($"Workbook saved to '{outputPath}' with DisableCss = true.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            DisableCssHtmlExport.Run();
        }
    }
}
