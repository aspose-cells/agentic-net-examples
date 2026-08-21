// Title: Export Named Ranges and Tables to Separate HTML Files with Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook, retrieves all defined named ranges and tables using GetNamedRangesAndTables(), converts each range to HTML with HtmlSaveOptions, sanitizes the range name for a safe file name, and writes the HTML output to individual .html files. Includes basic error handling and console feedback.
// Keywords: Aspose.Cells export named range HTML | C# Aspose.Cells GetNamedRangesAndTables | range.ToHtml example | save Excel named ranges as HTML | Aspose.Cells HTML documentation | sanitize file names C#
// Common Searches: export named ranges to html aspnet cells | c# convert excel named range to html file | aspocells getnamedrangesandtables usage | how to save each excel table as html c# | safe filename generation for exported ranges
// Developer Intent: Create individual HTML documents for every named range and table in a workbook to support documentation or web publishing.
// Use Cases: Produce HTML reference sheets for all data model ranges. | Generate web‑ready snapshots of tables for dashboards. | Track changes to range definitions via version‑controlled HTML files.
// AI Prompts: Write C# code that loads an Excel file with Aspose.Cells, extracts all named ranges and tables, and saves each as an HTML file with a sanitized filename. | Provide a method that takes a Workbook object and returns a dictionary mapping each range name to its HTML string using Aspose.Cells. | Explain how to configure HtmlSaveOptions to include gridlines, column headers, and custom CSS when exporting named ranges to HTML.

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsNamedRangesToHtml
{
    // Loads an Excel workbook, retrieves all defined named ranges and tables using GetNamedRangesAndTables(), converts each range to HTML with HtmlSaveOptions, sanitizes the range name for a safe file name, and writes the HTML output to individual .html files. Includes basic error handling and console feedback.
    class Program
    {
        static void Main()
        {
            const string inputPath = "InputWorkbook.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Retrieve all named ranges (and tables) defined in the workbook
                AsposeRange[] namedRanges = workbook.Worksheets.GetNamedRangesAndTables();

                if (namedRanges == null || namedRanges.Length == 0)
                {
                    Console.WriteLine("No named ranges found in the workbook.");
                    return;
                }

                // Iterate through each named range and export it to an individual HTML file
                for (int i = 0; i < namedRanges.Length; i++)
                {
                    AsposeRange range = namedRanges[i];

                    // Prepare HTML save options – default ExportNamedRangeAnchors (true) is kept
                    HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

                    // Convert the range to HTML; the method returns the HTML content as a byte array
                    byte[] htmlBytes = range.ToHtml(htmlOptions);

                    // Determine a safe file name: use the range name if available, otherwise use the index
                    string safeName = string.IsNullOrWhiteSpace(range.Name) ? $"Range_{i + 1}" : range.Name;

                    // Replace any characters that are invalid in file names
                    foreach (char c in Path.GetInvalidFileNameChars())
                    {
                        safeName = safeName.Replace(c, '_');
                    }

                    string outputPath = $"{safeName}.html";

                    // Write the HTML bytes to the file system
                    File.WriteAllBytes(outputPath, htmlBytes);

                    Console.WriteLine($"Exported named range '{range.Name}' to '{outputPath}'.");
                }

                Console.WriteLine("All named ranges have been exported to HTML.");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
