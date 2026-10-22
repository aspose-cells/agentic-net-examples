// Title: How to enable the UseOptimizedMemory flag in Aspose.Cells LoadOptions for faster loading of large Excel workbooks (C#)
// Common Searches: Aspose.Cells C# load workbook with thousands of rows using optimized memory | Set MemorySetting.MemoryPreference in LoadOptions to improve Excel load speed | Enable UseOptimizedMemory when opening large .xlsx files with Aspose.Cells | C# example of loading large Excel file with reduced memory usage in Aspose.Cells | Performance tuning Aspose.Cells workbook loading for big datasets | LoadOptions memory optimization for Excel files in .NET
// Tags: Aspose.Cells LoadOptions memory optimization | UseOptimizedMemory flag C# | load large Excel workbook Aspose.Cells | MemorySetting.MemoryPreference performance | reduce memory usage Aspose.Cells workbook loading | high row count Excel loading .NET

using System;
using System.IO;
using Aspose.Cells;

// The sample verifies the input file, configures LoadOptions with MemorySetting.MemoryPreference (UseOptimizedMemory), loads the workbook using these options, and saves it, demonstrating how to lower memory consumption and speed up loading of workbooks containing thousands of rows.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the input workbook
            string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Configure load options to use optimized memory handling
            LoadOptions loadOptions = new LoadOptions
            {
                // Use memory optimization setting (available in Aspose.Cells)
                MemorySetting = MemorySetting.MemoryPreference
            };

            // Load the workbook with the specified options
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // (Optional) Perform any workbook operations here

            // Save the workbook to the output file
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
