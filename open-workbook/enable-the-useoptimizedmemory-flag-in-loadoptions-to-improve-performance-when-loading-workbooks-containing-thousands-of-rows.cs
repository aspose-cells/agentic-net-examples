// Title: Enable Aspose.Cells optimized memory mode (MemoryPreference) for loading large Excel files in .NET
// Description: Shows how to load a workbook containing thousands of rows with Aspose.Cells by setting LoadOptions.MemorySetting to MemoryPreference, reducing RAM consumption, reading a cell value, and saving the workbook with proper error handling.
// Keywords: Aspose.Cells | .NET | LoadOptions | MemoryPreference | optimized memory | large Excel workbook | thousands of rows | reduce RAM usage | UseOptimizedMemory | Excel performance
// Common Searches: Aspose.Cells enable optimized memory loading | Load large Excel workbook with low memory .NET | MemorySetting.MemoryPreference example Aspose.Cells | UseOptimizedMemory flag Aspose.Cells C# | Reduce RAM usage when opening big Excel files
// Developer Intent: Load a workbook with a reduced memory footprint by activating Aspose.Cells' optimized memory setting.
// Use Cases: Read‑only analysis of massive spreadsheets in a desktop app without exhausting system RAM. | Processing large Excel files in a web API where memory quotas are strict. | Batch conversion of thousands of rows to another format while keeping the load‑phase memory low.
// AI Prompts: Provide a C# snippet that configures Aspose.Cells LoadOptions to use MemoryPreference and loads a large Excel file. | Explain how MemorySetting.MemoryPreference differs from the older UseOptimizedMemory flag and when to choose each. | Show how to safely iterate through rows of a huge workbook after loading it with optimized memory mode.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to load a workbook containing thousands of rows with Aspose.Cells by setting LoadOptions.MemorySetting to MemoryPreference, reducing RAM consumption, reading a cell value, and saving the workbook with proper error handling.
    public class OptimizedMemoryLoadDemo
    {
        public static void Run()
        {
            // Path to the source workbook (replace with your actual file)
            string inputPath = "LargeDataWorkbook.xlsx";

            // Prevent FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Create LoadOptions and enable optimized memory usage
                LoadOptions loadOptions = new LoadOptions
                {
                    // MemoryPreference reduces memory consumption at the cost of some performance overhead
                    MemorySetting = MemorySetting.MemoryPreference
                };

                // Load the workbook using the configured LoadOptions
                Workbook workbook = new Workbook(inputPath, loadOptions);

                // (Optional) Perform any read‑only operations here.
                // For example, read the value of the first cell:
                string firstCellValue = workbook.Worksheets[0].Cells["A1"].StringValue;
                Console.WriteLine($"First cell value: {firstCellValue}");

                // Save the workbook after processing
                string outputPath = "OptimizedMemoryOutput.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved with optimized memory setting to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            OptimizedMemoryLoadDemo.Run();
        }
    }
}
