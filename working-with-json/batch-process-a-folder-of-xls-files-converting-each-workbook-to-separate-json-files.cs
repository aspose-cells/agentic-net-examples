// Title: Convert every XLS workbook in a folder to separate JSON files with Aspose.Cells for .NET (C# batch example)
// AI Prompts: Write a C# console program that enumerates all .xls files in a given directory and uses Aspose.Cells to save each workbook as a .json file in a specified output folder. | Extend the batch converter to walk through subfolders recursively, preserving the original folder hierarchy when creating the JSON files. | Add comprehensive error handling and logging to the XLS‑to‑JSON batch process, capturing file‑access and conversion exceptions.
// Common Searches: c# Aspose.Cells batch convert xls files to json in a folder | how to export multiple Excel workbooks to json using Aspose.Cells .NET | save each workbook as json while iterating over files in a directory c# | recursive conversion of xls to json with Aspose.Cells example | log errors during bulk Excel to JSON conversion c# Aspose
// Tags: Aspose.Cells XLS to JSON batch conversion | C# directory enumeration for Excel files | Save workbook as JSON using Aspose.Cells | Recursive folder processing with Aspose.Cells | Error handling in bulk Excel conversion C# | Automated Excel to JSON export .NET

using System;
using System.IO;
using Aspose.Cells;

// A C# console application that scans a specified input folder for .xls files, loads each workbook with Aspose.Cells, and saves it as an individual .json file in an output directory, creating the output folder if needed and reporting any conversion errors.
class XlsToJsonBatchConverter
{
    static void Main(string[] args)
    {
        // Input folder containing XLS files
        string inputFolder = @"C:\InputXlsFolder";

        // Output folder where JSON files will be saved
        string outputFolder = @"C:\OutputJsonFolder";

        // Verify input folder exists
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder not found: {inputFolder}");
            return;
        }

        // Ensure output directory exists
        if (!Directory.Exists(outputFolder))
        {
            Directory.CreateDirectory(outputFolder);
        }

        string[] xlsFiles;
        try
        {
            // Get all .xls files in the input folder (non‑recursive)
            xlsFiles = Directory.GetFiles(inputFolder, "*.xls", SearchOption.TopDirectoryOnly);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error accessing input folder: {ex.Message}");
            return;
        }

        foreach (string xlsPath in xlsFiles)
        {
            try
            {
                // Load the workbook from the XLS file
                Workbook workbook = new Workbook(xlsPath);

                // Build the output JSON file path (same name, .json extension)
                string fileNameWithoutExt = Path.GetFileNameWithoutExtension(xlsPath);
                string jsonPath = Path.Combine(outputFolder, fileNameWithoutExt + ".json");

                // Save the workbook as JSON
                workbook.Save(jsonPath, SaveFormat.Json);

                Console.WriteLine($"Converted: {xlsPath} -> {jsonPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{xlsPath}': {ex.Message}");
            }
        }

        Console.WriteLine("Batch conversion completed.");
    }
}
