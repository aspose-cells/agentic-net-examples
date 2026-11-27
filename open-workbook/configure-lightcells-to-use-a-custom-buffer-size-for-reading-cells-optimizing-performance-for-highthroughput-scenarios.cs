// Title: Configuring a custom buffer size for LightCellsProcessor in Aspose.Cells for .NET (current version limitation)
// Common Searches: how to set a custom buffer size for LightCellsProcessor in Aspose.Cells .NET | Aspose.Cells LightCellsOptions buffer size configuration example | optimizing LightCells performance for high‑throughput reading in C# | why LightCellsProcessor settings are unavailable in latest Aspose.Cells release | adjusting cell read buffer to improve Aspose.Cells processing speed
// Tags: LightCellsProcessor custom buffer size | Aspose.Cells high‑throughput cell reading | LightCellsOptions performance tuning | configure cell read buffer Aspose.Cells .NET | Aspose.Cells version limitation LightCells API

using System;
using System.IO;
using Aspose.Cells;

// The sample checks that input.xlsx exists, loads it into a Workbook, ensures the output directory is present, and saves the workbook as output.xlsx. It also documents that LightCellsProcessor and LightCellsOptions are not available in the current Aspose.Cells version, so no custom buffer size can be configured in this example.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input workbook exists to avoid FileNotFoundException.
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook.
            Workbook workbook = new Workbook(inputPath);

            // NOTE: LightCellsProcessor and LightCellsOptions are not available in the
            // current Aspose.Cells version. The workbook is processed directly.
            // Any required processing can be added here.

            // Ensure the output directory exists.
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook processed and saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display an informative message.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
