// Title: How to autofit columns and rows after loading an HTML file into an Aspose.Cells workbook and export it back to HTML with headings preserved (C#)
// AI Prompts: Load an HTML document into a Workbook, call AutoFitColumns and AutoFitRows on the first worksheet, then save the workbook as HTML with ExportHeadings enabled using HtmlSaveOptions. | Show how to configure HtmlSaveOptions in C# to retain table headings when exporting a resized worksheet back to HTML with Aspose.Cells.
// Common Searches: C# Aspose.Cells autofit columns after importing HTML file | Preserve table headings when saving workbook to HTML with Aspose.Cells | How to use HtmlSaveOptions ExportHeadings in Aspose.Cells .NET | Resize rows and columns in Aspose.Cells before exporting to HTML | Load HTML into Aspose.Cells workbook and re-export with same layout
// Tags: autofit columns rows Aspose.Cells | htmlsaveoptions exportheadings C# | load html workbook Aspose.Cells | export workbook to html preserve layout | auto resize worksheet cells Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

namespace AsposeCellsHtmlExample
{
    // The example loads an HTML file into an Aspose.Cells Workbook, checks for a worksheet, applies AutoFitColumns and AutoFitRows to the first sheet, configures HtmlSaveOptions with ExportHeadings enabled, ensures the output directory exists, and saves the workbook back to HTML while preserving the original table structure.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths to the source HTML file and the output HTML file
            string inputHtmlPath = "input.html";
            string outputHtmlPath = "output.html";

            try
            {
                // Verify that the input HTML file exists
                if (!File.Exists(inputHtmlPath))
                {
                    Console.WriteLine($"Input file not found: {inputHtmlPath}");
                    return;
                }

                // Load the HTML content into a workbook
                Workbook workbook = new Workbook(inputHtmlPath);

                // Ensure there is at least one worksheet
                if (workbook.Worksheets.Count == 0)
                {
                    Console.WriteLine("The workbook does not contain any worksheets.");
                    return;
                }

                // Autofit columns and rows in the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                sheet.AutoFitColumns();
                sheet.AutoFitRows();

                // Configure HTML export options
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
                {
                    // Export column and row headings to preserve table structure
                    ExportHeadings = true
                };

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputHtmlPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as HTML with the specified options
                workbook.Save(outputHtmlPath, htmlOptions);
                Console.WriteLine($"HTML exported successfully to: {outputHtmlPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
