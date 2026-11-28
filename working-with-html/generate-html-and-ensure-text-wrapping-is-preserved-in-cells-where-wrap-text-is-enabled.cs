// Title: Export Excel workbook to HTML with cell text wrapping preserved using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# export Excel to HTML keep text wrap enabled | How to retain cell text wrapping when saving workbook as HTML with Aspose.Cells | C# generate self‑contained HTML from Excel with wrapped cell content | Preserve wrapped text in HTML output using Aspose.Cells SaveOptions | Auto‑fit row height for wrapped text before HTML export Aspose.Cells
// Tags: Aspose.Cells HTML export with text wrapping | C# preserve cell wrap in HTML output | Excel to self‑contained HTML Aspose.Cells | auto‑fit row height wrapped text Aspose | Export images as Base64 Aspose.Cells HTML | Set IsTextWrapped style Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

namespace AsposeCellsHtmlWrapDemo
{
    // The example creates a workbook, inserts a long string into cell A1, enables the IsTextWrapped style, auto‑fits the row height, and saves the file as a self‑contained HTML document with images embedded as Base64, ensuring that the wrapped text appears correctly in the generated HTML.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Define long text that should wrap inside the cell
                string longText = "This is a very long piece of text that is intended to demonstrate how text wrapping works when the cell is exported to HTML. " +
                                  "The text should appear on multiple lines inside the same cell in the generated HTML file.";

                // Put the long text into cell A1
                Cell cell = cells["A1"];
                cell.PutValue(longText);

                // Enable text wrapping for the cell
                Style style = cell.GetStyle();
                style.IsTextWrapped = true;
                cell.SetStyle(style);

                // Auto‑fit the row height so that all wrapped lines are visible
                worksheet.AutoFitRow(0);

                // Prepare HTML save options
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
                {
                    // Export images (if any) as Base64 strings to keep the HTML self‑contained.
                    ExportImagesAsBase64 = true
                };

                // Define output file path
                string outputPath = "WrappedTextOutput.html";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as HTML
                workbook.Save(outputPath, htmlOptions);

                Console.WriteLine($"Workbook successfully saved as HTML to '{outputPath}'.");
                Console.WriteLine("Open the file in a browser to verify that the text in cell A1 is wrapped.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while generating the HTML file:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
