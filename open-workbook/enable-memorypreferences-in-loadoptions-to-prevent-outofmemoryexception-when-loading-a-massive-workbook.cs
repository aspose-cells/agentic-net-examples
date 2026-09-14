// Title: Load a massive XLSX workbook with Aspose.Cells for .NET using LoadOptions.MemorySetting = MemoryPreference to avoid OutOfMemoryException
// AI Prompts: Write C# code that opens a large .xlsx file with Aspose.Cells, sets LoadOptions.MemorySetting to MemoryPreference, and saves the workbook to a new file. | Show how to verify the input file exists and catch exceptions while loading a huge workbook with memory‑optimized LoadOptions in Aspose.Cells. | Demonstrate configuring LoadOptions for memory‑efficient processing of big Excel spreadsheets in a .NET application using Aspose.Cells.
// Common Searches: asp.net how to open a huge xlsx file with aspose.cells memorypreference | c# load large excel workbook without outofmemoryexception using aspose cells | set loadoptions memorysetting to memorypreference for massive workbook | asp.net core aspose cells memory optimization when loading big spreadsheets
// Tags: Aspose.Cells LoadOptions MemoryPreference | load large XLSX workbook memory optimization | prevent OutOfMemoryException Aspose.Cells | C# Aspose.Cells memory‑efficient workbook loading | massive Excel file handling .NET

using System;
using System.IO;
using Aspose.Cells;

// The example checks that a massive XLSX file exists, configures LoadOptions with MemorySetting = MemoryPreference to reduce memory consumption, loads the workbook using these options, optionally processes it, saves the result to a new file, and handles any exceptions gracefully.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "massive_workbook.xlsx";
            string outputPath = "processed_workbook.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Configure load options with memory preference to reduce memory usage
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                MemorySetting = MemorySetting.MemoryPreference
            };

            // Load the workbook using the configured options
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // (Optional) Perform any required operations on the workbook here

            // Save the processed workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
