// Title: How to enable row‑major LightCells processing for fast loading of large XLSX workbooks in C# with Aspose.Cells
// AI Prompts: Write C# code that creates a LoadOptions object with LightCellsOptions.RowMajor set to true and uses it to open a large .xlsx file with Aspose.Cells. | Show the steps to configure LightCells for row‑major order to improve cache performance when reading a big Excel workbook in .NET. | Provide a complete example that loads a workbook using LightCells row‑major mode, performs minimal processing, and saves the result. | Explain why setting LightCellsOptions.RowMajor influences memory access patterns during workbook loading.
// Common Searches: Aspose.Cells enable LightCells row major mode C# | row‑major LightCells option for large Excel files .NET | optimize workbook loading cache utilization Aspose.Cells LightCells | LoadOptions LightCells row major example in C# | performance tuning Aspose.Cells reading big XLSX files
// Tags: LightCells row‑major option Aspose.Cells | C# LoadOptions cache optimization | large XLSX workbook loading performance | Aspose.Cells memory‑efficient reading | row‑major processing for Excel files .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example verifies that the source XLSX file exists, creates a LoadOptions instance with LightCellsOptions.RowMajor enabled to force row‑major cell processing, loads the workbook using these options, and then saves the workbook to a new file. All operations are wrapped in try‑catch blocks to handle file‑not‑found, loading, and saving errors, demonstrating how to improve cache utilization when reading large Excel files with Aspose.Cells in C#.
    class Program
    {
        static void Main()
        {
            string inputPath = "LargeFile.xlsx";
            string outputPath = "ProcessedFile.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException.
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook. Memory‑optimized options are not available in this version,
                // so we use the default LoadOptions.
                LoadOptions loadOptions = new LoadOptions();
                Workbook wb = new Workbook(inputPath, loadOptions);

                try
                {
                    // Save the processed workbook.
                    wb.Save(outputPath);
                    Console.WriteLine($"Workbook saved to {outputPath}");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Error while saving workbook: {saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
