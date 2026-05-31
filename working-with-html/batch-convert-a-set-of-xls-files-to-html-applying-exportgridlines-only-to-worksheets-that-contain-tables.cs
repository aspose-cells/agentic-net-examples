using System;
using System.IO;
using Aspose.Cells;

namespace BatchConversion
{
    class BatchXlsToHtml
    {
        static void Main()
        {
            // Folder containing the source XLS/XLSX files
            string sourceFolder = @"C:\InputXls";
            // Folder where the HTML files will be saved
            string outputFolder = @"C:\OutputHtml";

            try
            {
                // Verify source folder exists
                if (!Directory.Exists(sourceFolder))
                {
                    Console.WriteLine($"Source folder not found: {sourceFolder}");
                    return;
                }

                // Ensure the output directory exists
                Directory.CreateDirectory(outputFolder);

                // Get all .xls and .xlsx files in the source folder
                string[] xlsFiles = Directory.GetFiles(sourceFolder, "*.xls", SearchOption.TopDirectoryOnly);
                string[] xlsxFiles = Directory.GetFiles(sourceFolder, "*.xlsx", SearchOption.TopDirectoryOnly);
                string[] files = new string[xlsFiles.Length + xlsxFiles.Length];
                xlsFiles.CopyTo(files, 0);
                xlsxFiles.CopyTo(files, xlsFiles.Length);

                foreach (string filePath in files)
                {
                    // Guard against missing files (should not happen, but safe)
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found, skipping: {filePath}");
                        continue;
                    }

                    // Load the workbook from the file
                    Workbook workbook = new Workbook(filePath);

                    // Process each worksheet
                    for (int i = 0; i < workbook.Worksheets.Count; i++)
                    {
                        Worksheet sheet = workbook.Worksheets[i];

                        // Determine if the worksheet contains any tables (ListObjects)
                        bool hasTable = sheet.ListObjects.Count > 0;

                        // Configure HTML save options
                        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                        {
                            ExportGridLines = hasTable,          // Export gridlines only when a table exists
                            ExportActiveWorksheetOnly = true    // Export only the active worksheet
                        };

                        // Set the current worksheet as active
                        workbook.Worksheets.ActiveSheetIndex = i;

                        // Build a safe output file name: OriginalFile_SheetName.html
                        string sheetNameSafe = sheet.Name.Replace(" ", "_");
                        string outputFileName = $"{Path.GetFileNameWithoutExtension(filePath)}_{sheetNameSafe}.html";
                        string outputPath = Path.Combine(outputFolder, outputFileName);

                        // Save the active worksheet as HTML
                        workbook.Save(outputPath, htmlOptions);
                    }
                }

                Console.WriteLine("Batch conversion completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}