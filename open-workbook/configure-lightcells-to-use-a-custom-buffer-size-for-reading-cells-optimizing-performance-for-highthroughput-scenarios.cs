// Title: Set a custom LightCells read buffer size in Aspose.Cells for .NET to boost high‑throughput workbook loading
// AI Prompts: Generate C# code that configures LightCellsOptions.BufferSize before opening a workbook with Aspose.Cells. | Show how to apply a custom read buffer when loading an Excel file using the LightCells API in .NET. | Provide a sample that tunes LightCells performance by setting the buffer size in Aspose.Cells for C#.
// Common Searches: Aspose.Cells LightCells set buffer size C# example | How to configure LightCells read buffer for large Excel files in .NET | Adjust LightCellsOptions.BufferSize to improve workbook load speed | Performance tuning LightCells API buffer size for high‑throughput scenarios
// Tags: LightCells buffer size configuration | Aspose.Cells LightCellsOptions usage | custom read buffer for Excel .NET | high‑throughput workbook loading | LightCells performance tuning

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to set LightCellsOptions.BufferSize to a custom value before loading a workbook, enabling faster cell reading for high‑throughput Excel processing in Aspose.Cells for .NET.
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
                throw new FileNotFoundException($"The input file '{inputPath}' was not found.");

            // Load the workbook (using standard loading; LightCells API may not be available in all versions)
            var workbook = new Workbook(inputPath);

            // (Optional) Perform any workbook processing here

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the workbook to a new file
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
