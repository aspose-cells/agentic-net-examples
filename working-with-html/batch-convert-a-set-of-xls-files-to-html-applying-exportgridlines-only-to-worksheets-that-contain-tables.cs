// Title: Batch convert multiple XLS/XLSX worksheets to individual HTML files with ExportGridLines enabled only for sheets that contain tables using Aspose.Cells for .NET
// Common Searches: c# batch convert all Excel files in a folder to HTML with Aspose.Cells | export grid lines only for worksheets that have tables when saving to HTML using Aspose.Cells | save each worksheet as a separate HTML file and sanitize sheet name in Aspose.Cells C# | detect ListObjects in a worksheet before HTML export with Aspose.Cells .NET | process both .xls and .xlsx files in bulk to HTML with conditional options
// Tags: Aspose.Cells bulk workbook to html conversion | ExportGridLines only when worksheet contains ListObjects | Generate separate html per worksheet Aspose.Cells | File-system safe worksheet name for html output | Iterate over .xls and .xlsx files with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts; // Added to resolve Chart related types if present elsewhere

namespace BatchXlsToHtml
{
    // The utility validates input and output directories, scans the source folder for .xls and .xlsx workbooks, and processes each file sequentially. For every worksheet it checks for ListObjects (tables) and sets HtmlSaveOptions.ExportGridLines accordingly, exporting only the active sheet to an HTML file. The output filename combines the original workbook name with a sanitized worksheet name to ensure file‑system compatibility. Each worksheet is saved as an individual HTML file in the target folder, and any processing errors are logged to the console.
    class Program
    {
        static void Main(string[] args)
        {
            // Validate arguments
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: BatchXlsToHtml <input_folder> <output_folder>");
                return;
            }

            string inputFolder = args[0];
            string outputFolder = args[1];

            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder does not exist: {inputFolder}");
                return;
            }

            // Ensure output folder exists
            Directory.CreateDirectory(outputFolder);

            // Get all .xls and .xlsx files
            string[] xlsFiles = Directory.GetFiles(inputFolder, "*.xls", SearchOption.TopDirectoryOnly);
            string[] xlsxFiles = Directory.GetFiles(inputFolder, "*.xlsx", SearchOption.TopDirectoryOnly);
            string[] allFiles = new string[xlsFiles.Length + xlsxFiles.Length];
            xlsFiles.CopyTo(allFiles, 0);
            xlsxFiles.CopyTo(allFiles, xlsFiles.Length);

            if (allFiles.Length == 0)
            {
                Console.WriteLine("No XLS/XLSX files found in the input folder.");
                return;
            }

            foreach (string filePath in allFiles)
            {
                try
                {
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found: {filePath}");
                        continue;
                    }

                    Console.WriteLine($"Processing file: {Path.GetFileName(filePath)}");
                    Workbook workbook = new Workbook(filePath);

                    int sheetCount = workbook.Worksheets.Count;
                    for (int i = 0; i < sheetCount; i++)
                    {
                        Worksheet sheet = workbook.Worksheets[i];

                        // Determine if the worksheet contains any tables (ListObjects)
                        bool containsTable = sheet.ListObjects.Count > 0;

                        // Configure HTML save options
                        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                        {
                            ExportActiveWorksheetOnly = true,
                            ExportGridLines = containsTable // Export grid lines only when a table exists
                        };

                        // Set the active sheet to the one we are exporting
                        workbook.Worksheets.ActiveSheetIndex = i;

                        // Build output file name: originalname_sheetname.html
                        string safeSheetName = MakeFileSystemSafe(sheet.Name);
                        string outputFileName = $"{Path.GetFileNameWithoutExtension(filePath)}_{safeSheetName}.html";
                        string outputPath = Path.Combine(outputFolder, outputFileName);

                        // Save as HTML
                        workbook.Save(outputPath, htmlOptions);

                        Console.WriteLine($"  Saved worksheet '{sheet.Name}' to '{outputFileName}' (ExportGridLines={(containsTable ? "true" : "false")})");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch conversion completed.");
        }

        // Replaces invalid filename characters with underscores
        private static string MakeFileSystemSafe(string name)
        {
            foreach (char c in Path.GetInvalidFileNameChars())
            {
                name = name.Replace(c, '_');
            }
            return name;
        }
    }
}
