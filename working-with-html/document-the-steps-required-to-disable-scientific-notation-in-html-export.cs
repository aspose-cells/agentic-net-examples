// Title: How to Prevent Scientific Notation in HTML Export with Aspose.Cells for .NET (C#)
// Description: Step‑by‑step guide that creates an in‑memory workbook, applies the Text number format (Number = 49) to cells that could be rendered in exponent form, and saves the file as HTML using HtmlSaveOptions. The resulting HTML displays the original numeric strings without scientific notation.
// Keywords: Aspose.Cells HTML export | disable exponent notation | C# scientific notation | Number format 49 | text format cells Aspose | .NET large numbers HTML | small decimal HTML output | prevent scientific notation Aspose.Cells | Aspose.Cells tutorial
// Common Searches: Aspose.Cells keep numbers as plain text in HTML | remove scientific notation when saving workbook to HTML | C# export to HTML without exponent format | how to set cell style to text for HTML output Aspose | large number display in Aspose.Cells HTML export
// Developer Intent: Export a workbook to HTML while ensuring all numeric values appear in their original decimal representation.
// Use Cases: Show product IDs or transaction numbers exactly as entered in web reports. | Present precise measurement values (e.g., 0.00000012345) on dashboards without auto‑conversion. | Create HTML invoices where every numeric field must retain its literal format.
// AI Prompts: Generate C# code that uses Aspose.Cells to save a workbook as HTML and forces every numeric cell to use the Text format (Number = 49). | Explain how to apply a Text style to specific cells before HTML export to avoid exponent notation in Aspose.Cells for .NET. | Provide a concise tutorial on configuring HtmlSaveOptions and cell styles to keep numbers unchanged during HTML conversion.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Step‑by‑step guide that creates an in‑memory workbook, applies the Text number format (Number = 49) to cells that could be rendered in exponent form, and saves the file as HTML using HtmlSaveOptions. The resulting HTML displays the original numeric strings without scientific notation.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook (in-memory)
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate cells with values that could be shown in scientific notation
                // Set the style to Text (Number format 49) to preserve plain representation
                Cell cellA1 = sheet.Cells["A1"];
                cellA1.PutValue(123456789012345L);
                cellA1.GetStyle().Number = 49; // Text format
                cellA1.SetStyle(cellA1.GetStyle());

                Cell cellA2 = sheet.Cells["A2"];
                cellA2.PutValue(0.00000012345);
                cellA2.GetStyle().Number = 49; // Text format
                cellA2.SetStyle(cellA2.GetStyle());

                // Configure HTML save options (default options are sufficient)
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

                string outputPath = "output.html";

                // Ensure the directory for the output file exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as HTML
                workbook.Save(outputPath, htmlOptions);

                Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
