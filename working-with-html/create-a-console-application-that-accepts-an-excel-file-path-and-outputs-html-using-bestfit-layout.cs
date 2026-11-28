// Title: Convert an Excel workbook to a single HTML file with auto‑fit columns and rows using Aspose.Cells in a C# console application
// Common Searches: C# console program to export Excel to HTML with auto‑fit column widths using Aspose.Cells | How to generate a single HTML file from an Excel workbook preserving grid lines and embedding images as Base64 in .NET | Aspose.Cells HtmlSaveOptions settings for best‑fit layout and presentation‑friendly output | Command‑line tool for converting .xlsx to .html with auto‑fitted rows and columns | Saving all worksheets to one HTML file with Aspose.Cells while embedding images
// Tags: auto-fit columns rows Aspose.Cells HTML conversion | embed images as Base64 Aspose.Cells HtmlSaveOptions | preserve grid lines Excel to HTML C# | presentation-preference layout Aspose.Cells | console app Excel to HTML conversion

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace ExcelToHtmlBestFit
{
    // A C# console application that takes an Excel file path, validates it, loads the workbook with Aspose.Cells, auto‑fits all columns and rows in each worksheet, configures HtmlSaveOptions to export every sheet, keep grid lines, embed images as Base64, and enable a presentation‑friendly layout, then saves the result as a single HTML file (defaulting to the same folder) and reports success or errors.
    class Program
    {
        static void Main(string[] args)
        {
            // Validate arguments
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: ExcelToHtmlBestFit <input_excel_path> [output_html_path]");
                return;
            }

            string inputPath = args[0];
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" does not exist.");
                return;
            }

            // Determine output path
            string outputPath;
            if (args.Length >= 2 && !string.IsNullOrWhiteSpace(args[1]))
            {
                outputPath = args[1];
            }
            else
            {
                string directory = Path.GetDirectoryName(inputPath) ?? Directory.GetCurrentDirectory();
                string fileNameWithoutExt = Path.GetFileNameWithoutExtension(inputPath);
                outputPath = Path.Combine(directory, fileNameWithoutExt + ".html");
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Apply best‑fit layout to every worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Auto‑fit all columns in the used range
                    sheet.AutoFitColumns();

                    // Auto‑fit all rows in the used range
                    sheet.AutoFitRows();
                }

                // Configure HTML save options
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    // Export the whole workbook (not only the active sheet)
                    ExportActiveWorksheetOnly = false,

                    // Keep grid lines visible for a faithful layout
                    ExportGridLines = true,

                    // Embed images as Base64 strings to keep a single output file
                    ExportImagesAsBase64 = true,

                    // Use a presentation‑friendly layout (optional, improves appearance)
                    PresentationPreference = true
                };

                // Save as HTML
                workbook.Save(outputPath, htmlOptions);

                Console.WriteLine($"Successfully converted \"{inputPath}\" to HTML.");
                Console.WriteLine($"Output file: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during conversion:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
