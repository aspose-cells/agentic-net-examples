// Title: Batch convert Excel files to CSV with Aspose.Cells – trim blanks, keep headers
// Description: A C# console app that scans a directory for .xlsx, .xls, and .xlsm workbooks, loads each with Aspose.Cells, applies TxtSaveOptions to remove leading empty rows/columns and export only the active sheet, then writes a CSV file with the same base name to a target folder.
// Keywords: Aspose.Cells | C# batch Excel to CSV | trim leading blanks | export active sheet | TxtSaveOptions CSV | folder processing | Excel to CSV automation | remove empty rows columns
// Common Searches: C# Aspose.Cells convert folder of Excel files to CSV | remove leading empty rows when saving CSV with Aspose.Cells | batch export active worksheet as CSV using Aspose.Cells | Aspose.Cells TxtSaveOptions example for CSV | keep column headers in CSV export Aspose.Cells
// Developer Intent: Programmatically transform every Excel workbook in a given folder into a CSV file, stripping initial blank rows/columns while preserving the first row as column headers.
// Use Cases: Automate nightly ingestion of Excel reports into a CSV‑based data pipeline. | Clean spreadsheets that contain leading empty rows or columns before legacy system import. | Generate CSV snapshots of the active sheet from multiple workbooks in a single batch job.
// AI Prompts: Write C# code using Aspose.Cells to iterate over a directory, load each .xlsx/.xls/.xlsm file, trim leading blank rows and columns, and save the active worksheet as a CSV with the same filename in another folder. | Explain how TxtSaveOptions.TrimLeadingBlankRowAndColumn works when exporting to CSV and why column headers remain intact. | Provide a step‑by‑step guide for batch converting Excel files to CSV with Aspose.Cells, including error handling for unsupported formats.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace BatchExcelToCsv
{
    // A C# console app that scans a directory for .xlsx, .xls, and .xlsm workbooks, loads each with Aspose.Cells, applies TxtSaveOptions to remove leading empty rows/columns and export only the active sheet, then writes a CSV file with the same base name to a target folder.
    class Program
    {
        static void Main()
        {
            // Folder containing the source Excel files
            string sourceFolder = @"C:\InputExcelFiles";

            // Folder where the resulting CSV files will be saved
            string outputFolder = @"C:\OutputCsvFiles";

            // Ensure the output directory exists
            Directory.CreateDirectory(outputFolder);

            // Get all Excel files (XLSX and XLS) in the source folder
            string[] excelFiles = Directory.GetFiles(sourceFolder, "*.*", SearchOption.TopDirectoryOnly);
            foreach (string excelPath in excelFiles)
            {
                // Process only supported Excel formats
                string ext = Path.GetExtension(excelPath).ToLowerInvariant();
                if (ext != ".xlsx" && ext != ".xls" && ext != ".xlsm")
                    continue;

                // Load the workbook (lifecycle: create & load)
                Workbook workbook = new Workbook(excelPath);

                // Configure CSV save options
                TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Csv)
                {
                    // Trim leading blank rows and columns (default is true, set explicitly for clarity)
                    TrimLeadingBlankRowAndColumn = true,

                    // Export only the active sheet (default is false)
                    ExportAllSheets = false
                };

                // Build the output CSV file path
                string csvFileName = Path.GetFileNameWithoutExtension(excelPath) + ".csv";
                string csvPath = Path.Combine(outputFolder, csvFileName);

                // Save the workbook as CSV (lifecycle: save)
                workbook.Save(csvPath, saveOptions);

                Console.WriteLine($"Converted '{excelPath}' to '{csvPath}'.");
            }

            Console.WriteLine("Batch conversion completed.");
        }
    }
}
