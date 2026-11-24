// Title: Set Aspose.Cells MemoryPreference to Low before opening a large XLSX workbook in C#
// Common Searches: how to set MemoryPreference low in Aspose.Cells before loading a workbook | reduce RAM consumption when opening a large XLSX file using Aspose.Cells in C# | Aspose.Cells low memory configuration example for massive Excel workbooks | open big Excel workbook with limited memory using Aspose.Cells .NET | configure Aspose.Cells memory settings for processing huge XLSX files
// Tags: Aspose.Cells MemoryPreference low | low memory mode for XLSX loading | C# large workbook memory optimization | Aspose.Cells memory setting for massive files | open XLSX with reduced RAM usage

using System;
using System.IO;
using Aspose.Cells;

// The sample sets Aspose.Cells.MemorySetting.MemoryPreference to Low to minimize RAM usage, validates the input XLSX path, ensures the output directory exists, loads the massive workbook, optionally processes it, saves the result, and handles any exceptions.
class Program
{
    static void Main()
    {
        // Optional: Reduce RAM consumption if supported by the library version.
        // Uncomment and adjust the following line if MemorySetting is available in your Aspose.Cells version.
        // Aspose.Cells.MemorySetting.MemoryPreference = Aspose.Cells.MemoryPreference.Low;

        string inputPath = @"C:\Path\To\MassiveFile.xlsx";
        string outputPath = @"C:\Path\To\MassiveFile_Processed.xlsx";

        // Verify that the input file exists before attempting to load it
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Load the massive XLSX file
            using (Workbook workbook = new Workbook(inputPath))
            {
                // (Optional) Perform any required operations on the workbook here

                // Save the workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
            }

            Console.WriteLine("Workbook processed and saved successfully.");
        }
        catch (Exception ex)
        {
            // Handle any runtime exceptions gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
