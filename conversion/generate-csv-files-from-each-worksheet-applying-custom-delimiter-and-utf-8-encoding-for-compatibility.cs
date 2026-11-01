// Title: Export every worksheet from an Excel workbook to separate UTF-8 CSV files with a custom delimiter using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export each Excel sheet to its own CSV file with semicolon delimiter | How to save worksheets as UTF-8 encoded CSV using Aspose.Cells TxtSaveOptions | Aspose.Cells .NET generate CSV per worksheet with custom separator | Programmatically create CSV files from all sheets in a workbook using Aspose.Cells | Set custom CSV delimiter and UTF-8 encoding when exporting Excel with Aspose.Cells
// Tags: Aspose.Cells export worksheet to CSV with custom delimiter | C# TxtSaveOptions CSV separator setting | UTF-8 CSV export from Excel using Aspose.Cells | save each Excel sheet as separate CSV .NET | Aspose.Cells per-sheet CSV generation

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsCsvExport
{
    // Loads an Excel workbook, iterates through its worksheets, and saves each sheet as an individual CSV file using Aspose.Cells with a semicolon as the separator and UTF-8 encoding.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source Excel workbook
            string inputFilePath = @"C:\Data\SourceWorkbook.xlsx";

            // Folder where CSV files will be saved
            string outputFolder = @"C:\Data\CsvExports";

            // Custom delimiter to use in CSV files (e.g., semicolon)
            char customDelimiter = ';';

            try
            {
                // Verify that the input file exists
                if (!File.Exists(inputFilePath))
                {
                    Console.WriteLine($"Input file not found: {inputFilePath}");
                    return;
                }

                // Ensure the output directory exists
                Directory.CreateDirectory(outputFolder);

                // Load the workbook (lifecycle rule: load)
                Workbook workbook = new Workbook(inputFilePath);

                // Preserve the original active sheet index
                int originalActiveSheet = workbook.Worksheets.ActiveSheetIndex;

                // Iterate through each worksheet in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    try
                    {
                        // Configure CSV save options with custom delimiter and UTF-8 encoding
                        TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv)
                        {
                            Separator = customDelimiter,   // Apply custom delimiter
                            Encoding = Encoding.UTF8       // Ensure UTF-8 encoding
                        };

                        // Set the current worksheet as the active sheet for export
                        workbook.Worksheets.ActiveSheetIndex = sheet.Index;

                        // Build the output CSV file path (one file per worksheet)
                        string csvFilePath = Path.Combine(outputFolder, $"{sheet.Name}.csv");

                        // Save the current worksheet as a CSV file (lifecycle rule: save)
                        workbook.Save(csvFilePath, csvOptions);
                    }
                    catch (Exception exSheet)
                    {
                        Console.WriteLine($"Failed to export sheet '{sheet.Name}': {exSheet.Message}");
                    }
                }

                // Restore the original active sheet index
                workbook.Worksheets.ActiveSheetIndex = originalActiveSheet;

                Console.WriteLine("CSV export completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred during CSV export: {ex.Message}");
            }
        }
    }
}
