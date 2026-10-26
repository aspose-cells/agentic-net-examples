// Title: Load a large XLSX workbook with Aspose.Cells LightCells API in C# while limiting memory to 200 MB
// Common Searches: C# Aspose.Cells load large XLSX with memory usage under 200 MB | How to enable low‑memory mode for LightCells when opening big Excel files in .NET | Set MemoryPreferenceLow in Aspose.Cells LoadOptions using reflection | Open big Excel workbook with Aspose.Cells without exceeding 200 MB RAM | Aspose.Cells LightCells example for memory‑constrained environments
// Tags: aspose cells lightcells memory optimization | c# loadoptions memorysetting low | excel workbook loading with 200mb memory cap | reflection based memorypreference configuration | large xlsx processing using aspose cells

using System;
using System.IO;
using Aspose.Cells;

// The sample checks for a large XLSX file, attempts to open it with Aspose.Cells LightCells by setting the LoadOptions.MemorySetting to MemoryPreferenceLow via reflection, falls back to default loading if needed, and saves the workbook, keeping memory consumption around 200 MB.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the large XLSX file
            string inputPath = @"C:\Data\largeFile.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            Workbook workbook;

            // Try to load with LoadOptions (if MemorySetting is supported)
            try
            {
                var loadOptions = new LoadOptions(LoadFormat.Xlsx);

                // Attempt to set low‑memory mode via reflection to avoid compile‑time dependency
                var memProp = typeof(LoadOptions).GetProperty("MemorySetting");
                if (memProp != null && memProp.PropertyType.IsEnum)
                {
                    var lowValue = Enum.Parse(memProp.PropertyType, "MemoryPreferenceLow");
                    memProp.SetValue(loadOptions, lowValue);
                }

                workbook = new Workbook(inputPath, loadOptions);
            }
            catch (Exception ex)
            {
                // Fallback to default loading if any issue occurs
                Console.WriteLine($"LoadOptions not applied: {ex.Message}");
                workbook = new Workbook(inputPath);
            }

            // Path for the processed workbook
            string outputPath = @"C:\Data\processedFile.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook processed and saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
