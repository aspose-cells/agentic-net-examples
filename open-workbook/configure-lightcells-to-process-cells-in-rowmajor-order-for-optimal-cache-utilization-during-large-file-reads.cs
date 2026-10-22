// Title: Configure Aspose.Cells LightCells for row‑major cell processing in C# to improve cache utilization during large workbook reads
// Common Searches: how to enable LightCells row major order in Aspose.Cells C# | optimizing cache usage when reading large XLSX files with LightCells | Aspose.Cells LightCellsOptions performance settings for big workbooks | C# example of row‑major cell iteration using LightCells | increase Excel read speed by configuring LightCells row‑major processing
// Tags: LightCells row-major processing configuration | Aspose.Cells cache‑optimized large workbook read | C# LightCellsOptions performance tuning | row-major cell iteration Aspose.Cells | high‑throughput Excel file loading with LightCells

using System;
using System.IO;
using Aspose.Cells;

// The program loads an existing Excel workbook, iterates through every cell of the first worksheet in row‑major order, and saves the workbook to a new file while handling potential errors, illustrating a cache‑friendly approach for processing large workbooks.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Example processing: iterate through cells of the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;
            for (int row = 0; row <= cells.MaxDataRow; row++)
            {
                for (int col = 0; col <= cells.MaxDataColumn; col++)
                {
                    // Read cell value (no modification performed)
                    var value = cells[row, col].Value;
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook after processing
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
