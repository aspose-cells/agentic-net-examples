// Title: C# Batch Convert Excel Files in a Folder to Individual JSON Using Aspose.Cells .NET
// Description: A console utility that checks a folder path, enumerates all *.xls and *.xlsx files, loads each workbook with Aspose.Cells, and saves it as a separate JSON file using default JsonSaveOptions. The program writes progress messages to the console and isolates errors per file.
// Keywords: Aspose.Cells | C# | .NET | Excel to JSON conversion | batch export Excel files | process folder of XLSX | JsonSaveOptions | automate Excel JSON export | convert multiple workbooks | console app
// Common Searches: how to batch convert Excel to JSON C# Aspose.Cells | C# code to export all .xls files in a directory as JSON | Aspose.Cells JsonSaveOptions example for multiple workbooks | automate folder processing of Excel files to JSON | convert a folder of Excel spreadsheets to JSON using .NET
// Developer Intent: Programmatically transform every Excel workbook in a specified directory into its own JSON file.
// Use Cases: Create JSON feeds from a collection of legacy Excel reports for API consumption. | Generate version‑controlled JSON snapshots of financial spreadsheets for audit trails. | Schedule nightly export of daily Excel logs to JSON for downstream analytics pipelines.
// AI Prompts: Add recursive sub‑folder traversal to the batch converter and preserve the original folder hierarchy in the JSON output. | Customize JsonSaveOptions to export only selected worksheets or to format dates in ISO 8601. | Replace console logging with a structured log file (e.g., JSON or CSV) that records success and error details for each processed workbook.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsBatchJsonExport
{
    // A console utility that checks a folder path, enumerates all *.xls and *.xlsx files, loads each workbook with Aspose.Cells, and saves it as a separate JSON file using default JsonSaveOptions. The program writes progress messages to the console and isolates errors per file.
    public static class WorkbookJsonExporter
    {
        /// <param name="folderPath">The full path to the folder containing the Excel files.</param>
        public static void ProcessFolder(string folderPath)
        {
            // Validate the folder path
            if (string.IsNullOrWhiteSpace(folderPath) || !Directory.Exists(folderPath))
            {
                Console.WriteLine("The specified folder does not exist.");
                return;
            }

            // Get all .xls and .xlsx files in the folder
            string[] excelFiles = Directory.GetFiles(folderPath, "*.xls*", SearchOption.TopDirectoryOnly);

            if (excelFiles.Length == 0)
            {
                Console.WriteLine("No Excel files found in the folder.");
                return;
            }

            foreach (string excelFilePath in excelFiles)
            {
                try
                {
                    // Ensure the file exists before attempting to load
                    if (!File.Exists(excelFilePath))
                    {
                        Console.WriteLine($"File not found: {excelFilePath}");
                        continue;
                    }

                    // Load the workbook
                    Workbook workbook = new Workbook(excelFilePath);

                    // Configure JSON save options (using defaults; adjust if needed)
                    JsonSaveOptions jsonOptions = new JsonSaveOptions();

                    // Determine the output JSON file path (same name, .json extension)
                    string jsonFilePath = Path.ChangeExtension(excelFilePath, ".json");

                    // Save the workbook as JSON
                    workbook.Save(jsonFilePath, jsonOptions);

                    Console.WriteLine($"Converted '{Path.GetFileName(excelFilePath)}' to '{Path.GetFileName(jsonFilePath)}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{excelFilePath}': {ex.Message}");
                }
            }
        }

        // Example entry point
        public static void Main(string[] args)
        {
            // Expect the folder path as the first argument; otherwise use a default path
            string folder = args.Length > 0 ? args[0] : @"C:\ExcelFiles";

            ProcessFolder(folder);
        }
    }
}
