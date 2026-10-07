// Title: How to export an Aspose.Cells workbook to HTML using HtmlCrossType.FitToCell and AddTooltipText to show full cell text on hover
// AI Prompts: Write C# code that creates a workbook, inserts a long string into a cell, and saves it as HTML with HtmlSaveOptions set to use HtmlCrossType.FitToCell and to include hover text for each cell. | Update an existing Aspose.Cells example to apply HtmlCrossType.FitToCell in the HTML save options while keeping the hover feature active for overflowing cell content.
// Common Searches: Aspose.Cells export to HTML with tooltip for long cell values | C# HtmlSaveOptions HtmlCrossType.FitToCell example | prevent text overflow in Aspose.Cells HTML output | add hover tooltip to cells when saving workbook as HTML using Aspose.Cells
// Tags: HTML cell fitting option Aspose.Cells | AddTooltipText HTML export | cell text overflow handling Aspose.Cells | tooltip for long cell values C#

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example creates a workbook, writes a lengthy string to cell B2, configures HtmlSaveOptions with HtmlCrossType.FitToCell and AddTooltipText enabled, and saves the workbook as an HTML file so the full text appears as a tooltip on hover, preventing overflow.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                var workbook = new Workbook();

                // Access the first worksheet
                var sheet = workbook.Worksheets[0];

                // Put a long text into a cell to demonstrate overflow handling
                sheet.Cells["B2"].PutValue("This is a very long piece of text that would normally overflow the cell boundaries when exported to HTML.");

                // Configure HTML save options
                var htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
                {
                    // Add the full cell text as a tooltip (hover text) to avoid overflow loss
                    AddTooltipText = true
                };

                // Define output file path
                string outputPath = "Output.html";

                // Ensure the directory for the output file exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as HTML using the configured options
                workbook.Save(outputPath, htmlOptions);
                Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while processing the workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
