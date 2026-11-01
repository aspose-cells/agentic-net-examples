// Title: How to export an Excel workbook to HTML with preserved text wrapping in cells using Aspose.Cells for .NET
// AI Prompts: Write C# code that enables IsTextWrapped for a specific cell, sets an appropriate column width, and saves the workbook as HTML with Aspose.Cells. | Show how to configure HtmlSaveOptions to embed images as Base64 when exporting a workbook that contains wrapped text.
// Common Searches: Aspose.Cells C# export worksheet to HTML keep cell text wrap | save Excel as HTML with wrapped text using Aspose.Cells .NET | how to enable text wrapping in HTML output from Aspose.Cells workbook
// Tags: Aspose.Cells export HTML with text wrap | C# set cell IsTextWrapped Aspose.Cells | HtmlSaveOptions embed images base64 Aspose | adjust column width for wrapped text Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlWrapExample
{
    // The example creates a workbook, writes a long string to cell A1, enables text wrapping for that cell, adjusts column A width, configures HtmlSaveOptions to embed images as Base64, and saves the result as WrappedTextOutput.html.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Set a long text value in cell A1
                Cell cell = sheet.Cells["A1"];
                cell.PutValue("This is a very long piece of text that should wrap inside the cell when exported to HTML.");

                // Enable text wrapping for the cell
                Style style = cell.GetStyle();
                style.IsTextWrapped = true; // Preserve wrap text
                cell.SetStyle(style);

                // Adjust the column width to demonstrate wrapping
                sheet.Cells.SetColumnWidth(0, 20); // Column A width

                // Prepare HTML save options
                HtmlSaveOptions saveOptions = new HtmlSaveOptions
                {
                    // Optional: embed images as base64 to keep a single HTML file
                    ExportImagesAsBase64 = true
                };

                // Save the workbook as an HTML file
                string outputPath = "WrappedTextOutput.html";
                workbook.Save(outputPath, saveOptions);

                Console.WriteLine($"Workbook saved as HTML to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
