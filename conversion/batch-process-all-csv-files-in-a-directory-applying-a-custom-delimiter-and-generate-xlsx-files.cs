// Title: Batch convert CSV to XLSX with a custom delimiter using Aspose.Cells for .NET (C#)
// Description: A C# console utility that scans a source folder, imports every *.csv file into an Aspose.Cells workbook using a user‑defined delimiter (e.g., semicolon), optionally converts numeric strings, and saves each workbook as an .xlsx file with the same base name in a target directory. Includes folder validation and robust error handling.
// Keywords: Aspose.Cells CSV import | custom delimiter CSV | batch CSV to XLSX | C# .NET Excel conversion | ImportCSV custom separator | SaveFormat.Xlsx | Windows file automation | multiple CSV processing | Excel workbook generation | data migration C#
// Common Searches: Aspose.Cells import CSV with custom delimiter C# | batch convert all CSV files to XLSX .NET | C# program to convert semicolon delimited CSV to Excel | how to process multiple CSV files with Aspose.Cells | automate CSV to XLSX conversion using Aspose
// Developer Intent: Build a command‑line tool that converts every CSV file in a folder to an XLSX workbook, applying a specified delimiter and preserving file names.
// Use Cases: Nightly automation that transforms legacy semicolon‑delimited export files into Excel workbooks for reporting dashboards. | Command‑line batch processor for incoming data feeds, converting each CSV to XLSX before loading into BI pipelines. | Scheduled Windows service that watches a drop folder, converts new CSV files to Excel with a custom separator, and archives the results.
// AI Prompts: Generate C# code using Aspose.Cells to import CSV files with a pipe (|) delimiter, convert them to XLSX, and include detailed error handling for missing or corrupt files. | Show how to recursively scan subdirectories, import each CSV with a custom delimiter, and save the resulting XLSX files preserving the original folder hierarchy using Aspose.Cells. | Create a PowerShell script that invokes a compiled .NET executable to batch convert CSV files with a user‑specified delimiter to XLSX, logging successes and failures.

using System;
using System.IO;
using Aspose.Cells;

// A C# console utility that scans a source folder, imports every *.csv file into an Aspose.Cells workbook using a user‑defined delimiter (e.g., semicolon), optionally converts numeric strings, and saves each workbook as an .xlsx file with the same base name in a target directory. Includes folder validation and robust error handling.
class CsvToXlsxBatch
{
    static void Main()
    {
        // Directory containing the source CSV files
        string sourceDir = @"C:\InputCsv";

        // Directory where the generated XLSX files will be saved
        string destDir = @"C:\OutputXlsx";

        try
        {
            // Verify source directory exists
            if (!Directory.Exists(sourceDir))
            {
                Console.WriteLine($"Source directory does not exist: {sourceDir}");
                return;
            }

            // Ensure the destination directory exists
            Directory.CreateDirectory(destDir);

            // Custom delimiter to be used when importing CSV files (e.g., semicolon)
            string customDelimiter = ";";

            // Retrieve all CSV files from the source directory
            string[] csvFiles = Directory.GetFiles(sourceDir, "*.csv");

            foreach (string csvPath in csvFiles)
            {
                try
                {
                    // Verify the CSV file exists before processing
                    if (!File.Exists(csvPath))
                    {
                        Console.WriteLine($"File not found: {csvPath}");
                        continue;
                    }

                    // Create a new empty workbook
                    Workbook workbook = new Workbook();

                    // Access the first worksheet
                    Worksheet sheet = workbook.Worksheets[0];

                    // Import the CSV file using the custom delimiter,
                    // convert numeric strings to numbers, start at cell A1 (row 0, column 0)
                    sheet.Cells.ImportCSV(csvPath, customDelimiter, true, 0, 0);

                    // Build the output XLSX file path (same name, different extension)
                    string outputFileName = Path.GetFileNameWithoutExtension(csvPath) + ".xlsx";
                    string outputPath = Path.Combine(destDir, outputFileName);

                    // Save the workbook as XLSX
                    workbook.Save(outputPath, SaveFormat.Xlsx);

                    Console.WriteLine($"Converted: {csvPath} -> {outputPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{csvPath}': {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
