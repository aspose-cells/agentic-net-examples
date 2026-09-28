// Title: Convert every XLS workbook in a directory to a separate JSON file using Aspose.Cells in C#
// AI Prompts: Write a C# console program that scans a given folder for *.xls files, loads each workbook with Aspose.Cells, and saves it as a JSON file with the same name. | Modify the batch converter to also handle *.xlsx files and include each worksheet’s name as a property in the generated JSON. | Add robust error handling and a log file that records successful conversions and any failures during the XLS‑to‑JSON batch process.
// Common Searches: C# batch convert all .xls files in a folder to JSON with Aspose.Cells | How to export multiple Excel workbooks to separate JSON files using .NET | Aspose.Cells SaveFormat.Json example for processing a directory of Excel files | Automate conversion of a folder of legacy XLS spreadsheets to JSON in C# | Loop through files in a directory and save each workbook as JSON using Aspose.Cells
// Tags: Aspose.Cells batch XLS to JSON conversion | C# folder iteration for Excel to JSON | SaveFormat.Json multiple workbooks | automated Excel workbook export .NET | error logging Aspose.Cells conversion

using System;
using System.IO;
using Aspose.Cells;

// A C# console utility that enumerates all *.xls files in a source folder, loads each workbook with Aspose.Cells, and saves each one as an individual JSON file in a target folder, with built‑in checks for missing files and exception handling.
class XlsToJsonBatchConverter
{
    static void Main(string[] args)
    {
        // Path to the folder containing XLS files
        string sourceFolder = @"C:\InputXlsFolder";

        // Path to the folder where JSON files will be saved
        string outputFolder = @"C:\OutputJsonFolder";

        // Verify source folder exists
        if (!Directory.Exists(sourceFolder))
        {
            Console.WriteLine($"Source folder not found: {sourceFolder}");
            return;
        }

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Get all .xls files in the source folder (including .xlsx if needed)
        string[] xlsFiles = Directory.GetFiles(sourceFolder, "*.xls", SearchOption.TopDirectoryOnly);

        foreach (string xlsPath in xlsFiles)
        {
            try
            {
                // Verify the file still exists before loading
                if (!File.Exists(xlsPath))
                {
                    Console.WriteLine($"File not found (skipped): {xlsPath}");
                    continue;
                }

                // Load the workbook from the XLS file
                Workbook workbook = new Workbook(xlsPath);

                // Determine the output JSON file name (same base name, .json extension)
                string fileNameWithoutExt = Path.GetFileNameWithoutExtension(xlsPath);
                string jsonPath = Path.Combine(outputFolder, fileNameWithoutExt + ".json");

                // Save the entire workbook as a JSON file
                workbook.Save(jsonPath, SaveFormat.Json);
                Console.WriteLine($"Converted: {xlsPath} -> {jsonPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{xlsPath}': {ex.Message}");
            }
        }

        Console.WriteLine("Conversion completed. JSON files are located in: " + outputFolder);
    }
}
