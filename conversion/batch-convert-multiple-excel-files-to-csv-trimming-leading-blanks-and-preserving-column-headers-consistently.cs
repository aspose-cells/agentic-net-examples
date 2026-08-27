// Title: Batch convert Excel workbooks to CSV with trimmed leading blanks using Aspose.Cells for .NET
// AI Prompts: Write a C# method that enumerates all Excel files in a directory and converts each to a CSV file, applying TxtSaveOptions to eliminate any leading empty rows or columns. | Modify the batch converter to generate a separate CSV file for every worksheet in each workbook while preserving the original column headers. | Enhance the conversion loop with try‑catch blocks that log failed files to a text report and continue processing the remaining workbooks.
// Common Searches: how to batch convert xlsx and ods files to csv with Aspose.Cells .NET | remove leading empty rows and columns when exporting Excel to csv using Aspose.Cells | save only the active worksheet as csv with Aspose.Cells C# | convert multiple Excel formats to csv programmatically Aspose.Cells | Aspose.Cells TxtSaveOptions trim leading blanks example
// Tags: batch excel to csv Aspose.Cells | TxtSaveOptions trim leading blanks | convert multiple workbook formats csv | export active sheet as csv C# | Aspose.Cells CSV conversion options

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsBatchConversion
{
    // The utility scans a source folder for Excel files (xlsx, xlsm, xls, xlsb, ods), loads each workbook with Aspose.Cells, and saves it as a CSV in a destination folder using TxtSaveOptions that trim leading blank rows and columns while exporting only the active sheet.
    public class ExcelToCsvConverter
    {
        /// <param name="sourceFolder">Folder containing Excel files (e.g., .xlsx, .xlsm).</param>
        /// <param name="destFolder">Folder where the resulting CSV files will be saved.</param>
        public static void BatchConvert(string sourceFolder, string destFolder)
        {
            // Ensure the destination directory exists
            if (!Directory.Exists(destFolder))
                Directory.CreateDirectory(destFolder);

            // Define the Excel file extensions to process
            string[] excelExtensions = new[] { "*.xlsx", "*.xlsm", "*.xls", "*.xlsb", "*.ods" };

            // Iterate over each supported extension
            foreach (var pattern in excelExtensions)
            {
                foreach (var excelPath in Directory.GetFiles(sourceFolder, pattern, SearchOption.TopDirectoryOnly))
                {
                    // Load the workbook (using default LoadOptions)
                    Workbook workbook = new Workbook(excelPath);

                    // Prepare CSV save options
                    TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Csv)
                    {
                        // Trim leading blank rows/columns like Excel does
                        TrimLeadingBlankRowAndColumn = true,

                        // Export only the active sheet (change to true if all sheets are required)
                        ExportAllSheets = false
                    };

                    // Build the output CSV file path (same name, .csv extension)
                    string csvFileName = Path.GetFileNameWithoutExtension(excelPath) + ".csv";
                    string csvPath = Path.Combine(destFolder, csvFileName);

                    // Save the workbook as CSV using the configured options
                    workbook.Save(csvPath, saveOptions);

                    Console.WriteLine($"Converted: {excelPath} -> {csvPath}");
                }
            }
        }

        // Example usage
        public static void Main()
        {
            string sourceFolder = @"C:\InputExcelFiles";
            string destFolder   = @"C:\OutputCsvFiles";

            BatchConvert(sourceFolder, destFolder);
        }
    }
}
