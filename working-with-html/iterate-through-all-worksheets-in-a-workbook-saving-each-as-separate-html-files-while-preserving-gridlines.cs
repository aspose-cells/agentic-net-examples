// Title: Export Each Excel Worksheet to Separate HTML Files with Gridlines – Aspose.Cells for .NET
// Description: Loads a workbook, iterates its worksheets, skips evaluation‑warning sheets, activates each sheet, and saves it as an HTML file with gridlines and a sanitized filename using Aspose.Cells.
// Keywords: Aspose.Cells HTML export | export worksheet to html | gridlines | C# Aspose.Cells | save each sheet as html | HtmlSaveOptions ExportGridLines | sanitize filename Aspose | iterate workbook worksheets
// Common Searches: Aspose.Cells export each sheet to html | C# save excel worksheets as separate html files | keep gridlines when converting excel to html | skip evaluation warning sheet Aspose.Cells | sanitize worksheet name for file output
// Developer Intent: Generate individual HTML files for all worksheets while preserving gridlines.
// Use Cases: Publish each sheet of a multi‑tab report as a web‑ready HTML page. | Automate conversion of Excel templates into separate HTML files for a documentation portal. | Process server‑side workbooks, omit evaluation‑warning tabs, and store each sheet with safe filenames.
// AI Prompts: Show a C# example that uses Aspose.Cells to convert every worksheet in a workbook to an HTML file with gridlines and sanitized filenames. | Explain how to add custom CSS to the generated HTML while keeping gridlines intact. | Provide a script to batch‑process a folder of Excel files, exporting each worksheet to HTML using the same options.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads a workbook, iterates its worksheets, skips evaluation‑warning sheets, activates each sheet, and saves it as an HTML file with gridlines and a sanitized filename using Aspose.Cells.
    class SaveWorksheetsToHtml
    {
        static void Main()
        {
            const string inputPath = "input.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through each worksheet
                for (int i = 0; i < workbook.Worksheets.Count; i++)
                {
                    Worksheet sheet = workbook.Worksheets[i];

                    // Skip Aspose evaluation warning sheets
                    if (sheet.Name.StartsWith("Evaluation Warning", StringComparison.OrdinalIgnoreCase))
                        continue;

                    // Set the current worksheet as active
                    workbook.Worksheets.ActiveSheetIndex = i;

                    // Configure HTML save options
                    HtmlSaveOptions options = new HtmlSaveOptions
                    {
                        ExportActiveWorksheetOnly = true,
                        ExportGridLines = true
                    };

                    // Build output file name based on worksheet name
                    string safeSheetName = sheet.Name.Replace(Path.GetInvalidFileNameChars(), '_');
                    string outputFile = $"{safeSheetName}.html";

                    try
                    {
                        // Save the active worksheet as an individual HTML file
                        workbook.Save(outputFile, options);
                        Console.WriteLine($"Worksheet \"{sheet.Name}\" saved to \"{outputFile}\".");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to save worksheet \"{sheet.Name}\": {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
            }
        }
    }

    // Extension method to replace invalid filename characters
    static class StringExtensions
    {
        public static string Replace(this string str, char[] chars, char replacement)
        {
            foreach (char c in chars)
                str = str.Replace(c, replacement);
            return str;
        }
    }
}
