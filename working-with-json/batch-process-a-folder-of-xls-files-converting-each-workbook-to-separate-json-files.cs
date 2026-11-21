// Title: How to batch‑convert all XLS workbooks in a folder to individual JSON files using Aspose.Cells for .NET (C#)
// Common Searches: C# code to convert every .xls file in a directory to JSON with Aspose.Cells | Aspose.Cells batch conversion of multiple XLS workbooks to JSON in .NET | How to process a folder of Excel 97‑2003 files and save each as a JSON file using Aspose.Cells | Command‑line tool for converting a folder of .xls files to .json using Aspose.Cells | Iterate over .xls files and export to JSON with Aspose.Cells SaveFormat.Json
// Tags: Aspose.Cells bulk XLS to JSON export | C# directory enumeration for Excel files | Workbook.Save with SaveFormat.Json usage | command‑line argument for source folder in .NET | exception handling during multi‑file Excel conversion

using System;
using System.IO;
using Aspose.Cells;

namespace XlsToJsonBatch
{
    // The sample program reads a source folder (provided via command line or defaults to the current directory), enumerates all *.xls files, loads each workbook with Aspose.Cells, and saves it as a JSON file using Workbook.Save with SaveFormat.Json. It logs each conversion and continues processing despite individual file errors.
    class Program
    {
        static void Main(string[] args)
        {
            // Determine source folder: command‑line argument or current directory.
            string sourceFolder = args.Length > 0 ? args[0] : Directory.GetCurrentDirectory();

            // Verify that the folder exists.
            if (!Directory.Exists(sourceFolder))
            {
                Console.WriteLine($"Folder not found: {sourceFolder}");
                return;
            }

            // Get all .xls files in the folder (top‑level only).
            string[] xlsFiles = Directory.GetFiles(sourceFolder, "*.xls", SearchOption.TopDirectoryOnly);

            if (xlsFiles.Length == 0)
            {
                Console.WriteLine("No XLS files found in the specified folder.");
                return;
            }

            foreach (string xlsPath in xlsFiles)
            {
                // Ensure the file actually exists before attempting to load.
                if (!File.Exists(xlsPath))
                {
                    Console.WriteLine($"File not found (skipped): {xlsPath}");
                    continue;
                }

                try
                {
                    // Load the workbook from the XLS file.
                    Workbook workbook = new Workbook(xlsPath);

                    // Build the output JSON file path (e.g., Report.xls -> Report.json).
                    string jsonFileName = Path.GetFileNameWithoutExtension(xlsPath) + ".json";
                    string jsonPath = Path.Combine(sourceFolder, jsonFileName);

                    // Save the workbook as JSON using the correct SaveFormat enum value.
                    workbook.Save(jsonPath, SaveFormat.Json);

                    Console.WriteLine($"Converted: {Path.GetFileName(xlsPath)} -> {jsonFileName}");
                }
                catch (Exception ex)
                {
                    // Log any errors but continue processing other files.
                    Console.WriteLine($"Error processing '{xlsPath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch conversion completed.");
        }
    }
}
