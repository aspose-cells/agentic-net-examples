// Title: C# example: Convert an Excel workbook to HTML with auto‑fitted column widths using Aspose.Cells
// Common Searches: Aspose.Cells C# export Excel to HTML with auto‑fit columns | How to keep column widths compact when converting .xlsx to HTML in .NET | Export all worksheets to a single HTML file with grid lines using Aspose.Cells | C# save workbook as HTML preserving column layout with HtmlSaveOptions | AutoFitColumns before HTML conversion in Aspose.Cells C#
// Tags: Aspose.Cells HtmlSaveOptions ExportGridLines | AutoFitColumns before HTML export | C# Excel to HTML multiple worksheets | compact column widths HTML conversion | save workbook as HTML with column optimization | Aspose.Cells column width handling in HTML

using System;
using System.IO;
using Aspose.Cells;

namespace ExcelToHtmlConversion
{
    // The program loads an Excel workbook, auto‑fits columns in every worksheet, configures HtmlSaveOptions to export all sheets with grid lines, and saves the result as an HTML file.
    class Program
    {
        static void Main(string[] args)
        {
            // Validate command‑line arguments
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: ExcelToHtmlConversion <input.xlsx> <output.html>");
                return;
            }

            string inputPath = args[0];
            string outputPath = args[1];

            // Ensure the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" does not exist.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);
                Console.WriteLine($"Loaded workbook \"{inputPath}\" successfully.");

                // AutoFit columns in each worksheet for compact widths
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    sheet.AutoFitColumns();
                    Console.WriteLine($"AutoFitted columns in worksheet \"{sheet.Name}\".");
                }

                // Configure HTML save options
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
                {
                    ExportActiveWorksheetOnly = false, // Export all worksheets
                    ExportGridLines = true              // Keep grid lines for readability
                };

                // Save the workbook as HTML
                workbook.Save(outputPath, htmlOptions);
                Console.WriteLine($"Workbook saved as HTML to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An exception occurred during conversion:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
