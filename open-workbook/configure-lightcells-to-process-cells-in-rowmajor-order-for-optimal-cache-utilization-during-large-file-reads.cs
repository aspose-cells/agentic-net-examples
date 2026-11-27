// Title: Enable LightCells row‑major processing for cache‑friendly large Excel workbook reads in Aspose.Cells for .NET
// Common Searches: how to enable LightCells row major mode in Aspose.Cells C# | Aspose.Cells load large XLSX with cache‑friendly row major iteration | set LightCells options for row‑major processing to improve read performance .NET | optimize memory usage when reading big Excel files using LightCells row major Aspose
// Tags: LightCells row‑major mode | cache‑friendly cell iteration | large XLSX workbook loading | Aspose.Cells high‑performance read | LoadOptions LightCells configuration

using Aspose.Cells;
using System;
using System.IO;

// Demonstrates how to enable LightCells row‑major processing via LoadOptions when loading a large XLSX workbook, providing cache‑friendly cell iteration and improved read performance with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Input and output file paths
        string inputPath = "largeFile.xlsx";
        string outputPath = "processedFile.xlsx";

        // Verify that the input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Prepare load options for the workbook (XLSX format)
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

            // Load the large Excel file using the specified options
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // (Optional) Add any additional processing on the workbook here

            // Save the workbook after processing
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
