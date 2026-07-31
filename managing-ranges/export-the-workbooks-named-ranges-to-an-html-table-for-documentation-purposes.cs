// Title: Export Excel Named Ranges and Tables to a Single HTML Document using Aspose.Cells for .NET
// Description: Loads an Excel workbook, extracts all defined named ranges and tables with GetNamedRangesAndTables, converts each range to HTML using HtmlSaveOptions (ExportNamedRangeAnchors = true), and merges the results into one HTML file that shows the range name followed by its table representation.
// Keywords: Aspose.Cells | C# export named ranges to HTML | GetNamedRangesAndTables | ExportNamedRangeAnchors | Excel to HTML conversion | named range documentation | Aspose.Range ToHtml | C# Excel HTML example | Aspose.Cells .NET tutorial | export named ranges .NET
// Common Searches: How to export all named ranges from an Excel workbook to HTML with Aspose.Cells | C# code to convert Excel named ranges and tables into HTML tables | Aspose.Cells example for generating HTML documentation of named ranges | Export named ranges to a single HTML file using .NET | Aspose.Cells GetNamedRangesAndTables HTML output
// Developer Intent: Create a single HTML file that lists every named range and table from an Excel workbook as separate tables for documentation or review.
// Use Cases: Provide auditors with a searchable HTML reference of all named ranges in a financial model. | Automate documentation of data‑validation ranges before publishing a shared spreadsheet. | Include an HTML snapshot of workbook named ranges in version‑control change logs.
// AI Prompts: Write a C# method that loads an Excel file with Aspose.Cells, iterates over all named ranges and tables, converts each to HTML with ExportNamedRangeAnchors enabled, and saves a combined HTML document. | Explain how the ExportNamedRangeAnchors option influences the HTML produced by Aspose.Range.ToHtml. | Suggest improvements for error handling, logging, and output‑path validation in the ExportNamedRangesToHtml example.

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Alias to avoid conflict with System.Range
    using AsposeRange = Aspose.Cells.Range;

    // Loads an Excel workbook, extracts all defined named ranges and tables with GetNamedRangesAndTables, converts each range to HTML using HtmlSaveOptions (ExportNamedRangeAnchors = true), and merges the results into one HTML file that shows the range name followed by its table representation.
    public class ExportNamedRangesToHtml
    {
        public static void Run(string inputFilePath, string outputHtmlPath)
        {
            try
            {
                // Verify input file exists
                if (!File.Exists(inputFilePath))
                {
                    Console.WriteLine($"Input file not found: {inputFilePath}");
                    return;
                }

                // Load the workbook from the specified file
                Workbook workbook = new Workbook(inputFilePath);

                // Retrieve all named ranges (and tables) defined in the workbook
                AsposeRange[] namedRanges = workbook.Worksheets.GetNamedRangesAndTables();

                // Prepare HTML save options; ensure named range anchors are exported
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    ExportNamedRangeAnchors = true
                };

                // Build a single HTML document that contains each named range as a separate table
                StringBuilder htmlBuilder = new StringBuilder();
                htmlBuilder.AppendLine("<!DOCTYPE html>");
                htmlBuilder.AppendLine("<html>");
                htmlBuilder.AppendLine("<head>");
                htmlBuilder.AppendLine("<meta charset=\"UTF-8\"/>");
                htmlBuilder.AppendLine("<title>Named Ranges Documentation</title>");
                htmlBuilder.AppendLine("</head>");
                htmlBuilder.AppendLine("<body>");
                htmlBuilder.AppendLine("<h1>Named Ranges</h1>");

                if (namedRanges != null && namedRanges.Length > 0)
                {
                    foreach (AsposeRange range in namedRanges)
                    {
                        // Convert the range to HTML using the provided ToHtml method
                        byte[] htmlBytes = range.ToHtml(htmlOptions);
                        string rangeHtml = Encoding.UTF8.GetString(htmlBytes);

                        // Add a heading for the named range and its HTML representation
                        htmlBuilder.AppendLine($"<h2>{range.Name}</h2>");
                        htmlBuilder.AppendLine(rangeHtml);
                    }
                }
                else
                {
                    htmlBuilder.AppendLine("<p>No named ranges found in the workbook.</p>");
                }

                htmlBuilder.AppendLine("</body>");
                htmlBuilder.AppendLine("</html>");

                // Write the combined HTML to the output file
                File.WriteAllText(outputHtmlPath, htmlBuilder.ToString(), Encoding.UTF8);
                Console.WriteLine($"HTML exported successfully to: {outputHtmlPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        // Entry point for the console application
        public static void Main(string[] args)
        {
            // Expecting two arguments: input Excel file path and output HTML file path
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: ExportNamedRangesToHtml <inputExcelPath> <outputHtmlPath>");
                return;
            }

            string inputPath = args[0];
            string outputPath = args[1];

            ExportNamedRangesToHtml.Run(inputPath, outputPath);
        }
    }
}
