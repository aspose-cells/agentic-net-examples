// Title: Convert a folder of XLSX workbooks to compact HTML with unused style removal using Aspose.Cells for .NET
// Common Searches: C# code to export multiple Excel files to HTML with Aspose.Cells and remove unused styles | How to batch process XLSX files and generate minimal HTML using Aspose.Cells .NET | Aspose.Cells RemoveUnusedStyles before saving workbook as HTML in a loop | Save Excel workbooks as HTML without gridlines and separate CSS using Aspose.Cells | Create output folder and export all .xlsx files to HTML with Aspose.Cells in C#
// Tags: batch xlsx to html Aspose.Cells | remove unused workbook styles Aspose.Cells | minimal html export HtmlSaveOptions C# | export excel as html without gridlines Aspose | process multiple Excel files programmatically .NET

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace BatchHtmlExport
{
    // The sample iterates over every .xlsx file in a specified directory, loads each workbook with Aspose.Cells, calls RemoveUnusedStyles to discard unused styles, configures HtmlSaveOptions for a compact HTML output (no gridlines, headings, base64 images, or separate CSS), and saves the result as an .html file in a designated output folder, handling errors and logging progress.
    class Program
    {
        static void Main(string[] args)
        {
            // Determine input and output directories
            string inputDir = args.Length > 0 ? args[0] : Directory.GetCurrentDirectory();
            string outputDir = args.Length > 1 ? args[1] : Path.Combine(Directory.GetCurrentDirectory(), "HtmlOutput");

            // Validate input directory
            if (!Directory.Exists(inputDir))
            {
                Console.Error.WriteLine($"Input directory does not exist: {inputDir}");
                return;
            }

            // Ensure output directory exists
            if (!Directory.Exists(outputDir))
            {
                try
                {
                    Directory.CreateDirectory(outputDir);
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"Failed to create output directory '{outputDir}': {ex.Message}");
                    return;
                }
            }

            // Get all .xlsx files in the input directory (non‑recursive)
            string[] xlsxFiles = Directory.GetFiles(inputDir, "*.xlsx", SearchOption.TopDirectoryOnly);
            if (xlsxFiles.Length == 0)
            {
                Console.WriteLine($"No .xlsx files found in '{inputDir}'.");
                return;
            }

            Console.WriteLine($"Processing {xlsxFiles.Length} file(s) from '{inputDir}' and saving HTML to '{outputDir}'.");

            foreach (string xlsxPath in xlsxFiles)
            {
                string fileName = Path.GetFileNameWithoutExtension(xlsxPath);
                string htmlPath = Path.Combine(outputDir, fileName + ".html");

                try
                {
                    // Verify the file still exists before loading
                    if (!File.Exists(xlsxPath))
                    {
                        Console.Error.WriteLine($"File not found: {xlsxPath}");
                        continue;
                    }

                    // Load the workbook
                    Workbook workbook = new Workbook(xlsxPath);

                    // Remove any styles that are not used in the workbook
                    workbook.RemoveUnusedStyles();

                    // Configure HTML save options for a compact output
                    HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                    {
                        ExportGridLines = false,
                        ExportRowColumnHeadings = false, // Updated property (ExportHeadings is obsolete)
                        ExportImagesAsBase64 = false,
                        ExportWorksheetCSSSeparately = false,
                        Encoding = Encoding.UTF8,
                        ExportActiveWorksheetOnly = false
                    };

                    // Save the workbook as HTML using the options overload
                    workbook.Save(htmlPath, htmlOptions);

                    Console.WriteLine($"Successfully saved: {htmlPath}");
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"Error processing '{xlsxPath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch processing completed.");
        }
    }
}
