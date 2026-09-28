// Title: How to enable MemoryPreference in Aspose.Cells LoadOptions for low‑memory loading of large XLSX workbooks (C#)
// AI Prompts: Generate C# code that creates a LoadOptions instance with MemorySetting.MemoryPreference and uses it to open an XLSX file via Aspose.Cells. | Show an example of loading a workbook containing thousands of rows with Aspose.Cells while minimizing RAM usage, then saving it to a new file. | Explain how to verify the input Excel file exists before applying the MemoryPreference setting in LoadOptions.
// Common Searches: Aspose.Cells C# load large Excel file with low memory consumption | Set MemorySetting.MemoryPreference in LoadOptions for XLSX files | How to reduce RAM usage when opening big spreadsheets with Aspose.Cells | LoadOptions MemoryPreference example for thousands of rows in .NET | C# Aspose.Cells optimized memory loading large workbook
// Tags: LoadOptions MemoryPreference Aspose.Cells | low‑memory workbook loading C# | optimize RAM usage when loading XLSX with Aspose | large spreadsheet memory optimization Aspose.Cells | C# Aspose.Cells LoadOptions configuration | memory‑efficient Excel import Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample checks that the source XLSX file exists, configures LoadOptions with MemorySetting.MemoryPreference to minimize RAM usage, loads the workbook, optionally processes it, and saves the result to a new file while handling any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Create LoadOptions and enable optimized memory usage
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                // Use memory‑optimized setting (less memory consumption)
                MemorySetting = MemorySetting.MemoryPreference
            };

            // Load the workbook using the configured LoadOptions
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // (Optional) Perform any required operations on the workbook here

            // Save the workbook to a new file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
