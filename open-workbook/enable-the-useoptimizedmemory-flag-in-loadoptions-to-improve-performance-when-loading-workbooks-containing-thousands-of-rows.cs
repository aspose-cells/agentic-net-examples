using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class OptimizedMemoryLoadDemo
    {
        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }

        public static void Run()
        {
            try
            {
                // Path to the source workbook (could be a large file)
                string sourcePath = "LargeDataWorkbook.xlsx";

                // Verify that the source file exists to avoid FileNotFoundException
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    return;
                }

                // Create LoadOptions and enable memory‑optimized mode
                LoadOptions loadOptions = new LoadOptions
                {
                    MemorySetting = MemorySetting.MemoryPreference // Use optimized memory
                };

                // Load the workbook with the specified options
                Workbook workbook = new Workbook(sourcePath, loadOptions);

                // Example operation: read the first worksheet name
                Worksheet sheet = workbook.Worksheets[0];
                Console.WriteLine($"First worksheet: {sheet.Name}");

                // Save the workbook (optional, could be to a different format)
                string outputPath = "OptimizedMemoryWorkbook.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved with MemoryPreference setting to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}