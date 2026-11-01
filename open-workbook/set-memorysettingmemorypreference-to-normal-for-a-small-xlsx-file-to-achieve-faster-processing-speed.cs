// Title: Set MemorySetting to Normal for faster processing of a small XLSX workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set workbook memory setting to Normal for small Excel files | How to improve load speed of a small XLSX using Aspose.Cells memory preference | Configure MemorySetting.Normal in Aspose.Cells .NET to speed up workbook processing | Best memory setting for small Excel workbook in Aspose.Cells C# | Increase performance when opening small XLSX files with Aspose.Cells
// Tags: MemorySetting.Normal Aspose.Cells .NET | fast workbook load small XLSX Aspose.Cells | configure workbook memory preference C# | optimize small Excel file processing Aspose.Cells | Aspose.Cells memory configuration for performance

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads a small input.xlsx file into an Aspose.Cells Workbook, sets Settings.MemorySetting to Normal to accelerate processing, saves the workbook to output.xlsx, and includes basic error handling.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing XLSX file
                Workbook workbook = new Workbook(inputPath);

                // Set memory preference to Normal for faster processing on small files
                workbook.Settings.MemorySetting = MemorySetting.Normal;

                // Save the workbook (overwrites or creates a new file)
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
