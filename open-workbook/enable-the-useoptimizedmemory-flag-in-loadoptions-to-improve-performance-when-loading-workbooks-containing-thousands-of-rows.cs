// Title: Load a large Excel workbook with optimized memory using Aspose.Cells LoadOptions in C#
// Common Searches: C# Aspose.Cells load large Excel file with MemorySetting.MemoryPreference | how to reduce memory consumption when opening big workbooks with Aspose.Cells | example of using LoadOptions to improve performance for thousands of rows in Excel | Aspose.Cells LoadOptions MemoryPreference for high‑row count worksheets in .NET | optimize workbook loading speed and memory usage in C# with Aspose.Cells
// Tags: Aspose.Cells LoadOptions MemoryPreference | memory‑optimized workbook loading C# | large Excel file load performance Aspose.Cells | high‑row count worksheet memory usage | load workbook with reduced memory .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The sample checks for the Excel file, creates a LoadOptions object with MemorySetting set to MemoryPreference to enable optimized memory usage, loads the workbook using these options, optionally processes it, and saves the result, demonstrating reduced memory consumption when handling workbooks with thousands of rows.
    class Program
    {
        static void Main()
        {
            // Path to the Excel file to be loaded
            string inputPath = @"C:\Data\LargeWorkbook.xlsx";

            // Path to save the workbook after loading (optional)
            string outputPath = @"C:\Data\LargeWorkbook_Optimized.xlsx";

            try
            {
                // Ensure the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Configure LoadOptions to reduce memory consumption
                LoadOptions loadOptions = new LoadOptions
                {
                    // Use memory preference setting for optimized memory usage
                    MemorySetting = MemorySetting.MemoryPreference
                };

                // Load the workbook with the specified LoadOptions
                Workbook workbook = new Workbook(inputPath, loadOptions);

                // (Optional) Perform any operations on the workbook here
                // ...

                // Save the workbook to verify that it loads and saves correctly
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any runtime exceptions gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
