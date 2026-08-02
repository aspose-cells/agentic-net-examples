// Title: Batch convert Excel workbooks to CSV with Aspose.Cells – trim leading blanks, preserve headers, export active sheet (C#)
// Description: A C# console utility that scans a folder for .xls, .xlsx, and .xlsm files, loads each workbook with Aspose.Cells, and saves the active worksheet as a CSV file. TxtSaveOptions are configured with TrimLeadingBlankRowAndColumn=true to drop empty rows/columns while keeping column headers. The resulting .csv files are written beside the original workbooks.
// Keywords: Aspose.Cells CSV conversion C# | batch Excel to CSV C# | trim leading blanks Aspose.Cells | export active worksheet to CSV | TxtSaveOptions SaveFormat.Csv | remove empty rows columns CSV | C# convert multiple Excel files | Aspose.Cells file processing example | data extraction from Excel C# | GitHub Aspose.Cells CSV sample
// Common Searches: C# batch convert Excel to CSV Aspose.Cells | How to remove empty rows when saving Excel as CSV using Aspose | Save only active sheet to CSV with Aspose.Cells | Trim leading blank columns in CSV output Aspose.Cells | Convert all .xls files in a directory to CSV programmatically
// Developer Intent: Convert every Excel file in a given folder to a CSV that contains only the active sheet, with leading blank rows and columns removed.
// Use Cases: Automate daily export of report workbooks to clean CSV files for loading into a data warehouse. | Process user‑submitted Excel templates, stripping empty rows/columns before feeding the data to downstream analytics pipelines. | Generate CSV extracts from the active sheet of each workbook to supply legacy systems that accept only CSV input. | Create a scheduled job that converts a batch of financial spreadsheets to CSV for archival and audit purposes.
// AI Prompts: Write C# code using Aspose.Cells to batch convert all .xls, .xlsx, and .xlsm files in a directory to CSV, trimming leading blanks and exporting only the active sheet. | Show how to add logging to the batch conversion program so that each successful conversion and any errors are written to a log file. | Explain the impact of TxtSaveOptions.TrimLeadingBlankRowAndColumn and ExportAllSheets on the CSV output produced by Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace BatchExcelToCsv
{
    // A C# console utility that scans a folder for .xls, .xlsx, and .xlsm files, loads each workbook with Aspose.Cells, and saves the active worksheet as a CSV file. TxtSaveOptions are configured with TrimLeadingBlankRowAndColumn=true to drop empty rows/columns while keeping column headers. The resulting .csv files are written beside the original workbooks.
    class Program
    {
        static void Main()
        {
            // Folder containing the Excel files to convert
            string sourceFolder = @"C:\InputExcelFiles";

            // Verify the source folder exists
            if (!Directory.Exists(sourceFolder))
            {
                Console.WriteLine($"Source folder not found: {sourceFolder}");
                return;
            }

            try
            {
                // Get all files in the folder (filter later by extension)
                string[] files = Directory.GetFiles(sourceFolder, "*.*", SearchOption.TopDirectoryOnly);

                foreach (string filePath in files)
                {
                    // Filter only supported Excel extensions
                    string ext = Path.GetExtension(filePath).ToLowerInvariant();
                    if (ext != ".xlsx" && ext != ".xls" && ext != ".xlsm")
                        continue;

                    // Ensure the file actually exists before loading
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found (skipped): {filePath}");
                        continue;
                    }

                    try
                    {
                        // Load the workbook
                        Workbook workbook = new Workbook(filePath);

                        // Configure CSV save options
                        TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv)
                        {
                            // Trim leading blank rows and columns
                            TrimLeadingBlankRowAndColumn = true,
                            // Export only the active sheet
                            ExportAllSheets = false
                        };

                        // Determine output CSV file path (same name, .csv extension)
                        string csvPath = Path.ChangeExtension(filePath, ".csv");

                        // Save the workbook as CSV using the configured options
                        workbook.Save(csvPath, csvOptions);

                        Console.WriteLine($"Converted '{Path.GetFileName(filePath)}' to '{Path.GetFileName(csvPath)}'.");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing file '{Path.GetFileName(filePath)}': {ex.Message}");
                    }
                }

                Console.WriteLine("Batch conversion completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
