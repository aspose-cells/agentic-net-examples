// Title: Export an entire Excel workbook to a single HTML file while keeping the original worksheet order using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx workbook with Aspose.Cells, configures HtmlSaveOptions to include all worksheets, and saves the result as one HTML file preserving the sheet sequence. | Show how to set Aspose.Cells HtmlSaveOptions.ExportActiveWorksheetOnly = false to convert a multi‑sheet Excel file to HTML without reordering the sheets.
// Common Searches: asp.net convert multi-sheet Excel to single HTML file preserving sheet order | c# Aspose.Cells HtmlSaveOptions export all worksheets to one HTML page | how to keep worksheet sequence when saving Excel as HTML using Aspose.Cells | save workbook as HTML with original sheet order Aspose.Cells .NET example
// Tags: Aspose.Cells HtmlSaveOptions export all worksheets | C# convert Excel workbook to single HTML page | maintain original sheet sequence Aspose.Cells | ExportActiveWorksheetOnly false example | multi-sheet workbook HTML export Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The program checks for InputWorkbook.xlsx, loads it with Aspose.Cells, sets HtmlSaveOptions.ExportActiveWorksheetOnly to false so every worksheet is included, and saves the workbook as OutputWorkbook.html, ensuring the original sheet order is retained.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputFile = "InputWorkbook.xlsx";
                string outputFile = "OutputWorkbook.html";

                // Verify that the input workbook exists
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Input file not found: {inputFile}");
                    return;
                }

                // Load the workbook from the file
                Workbook workbook = new Workbook(inputFile);

                // Set up HTML save options (export all worksheets)
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    ExportActiveWorksheetOnly = false
                };

                // Save the workbook as HTML using the options
                workbook.Save(outputFile, htmlOptions);
                Console.WriteLine($"Workbook successfully exported to: {outputFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
