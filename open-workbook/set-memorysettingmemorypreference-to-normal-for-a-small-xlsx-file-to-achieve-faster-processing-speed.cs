// Title: Set MemorySetting.MemoryPreference to Normal in Aspose.Cells for .NET when opening a small XLSX workbook to boost performance
// AI Prompts: Generate C# code that sets MemorySetting.MemoryPreference to Normal before opening an XLSX workbook with Aspose.Cells. | Show how to create a LoadOptions object with a Normal memory preference and use it to load a workbook in Aspose.Cells. | Provide a full example that applies a Normal memory setting, modifies the workbook, and saves it using Aspose.Cells.
// Common Searches: c# aspocells set memorypreference normal for faster load | how to use LoadOptions to improve Aspose.Cells workbook speed | increase loading speed of small Excel files with Aspose.Cells | configure Aspose.Cells memory settings in .NET application
// Tags: Aspose.Cells MemoryPreference Normal | LoadOptions memory configuration | C# workbook load performance | small XLSX memory optimization | Aspose.Cells memory tuning

using System;
using System.IO;
using Aspose.Cells;

// The example creates a LoadOptions instance, sets MemorySetting.MemoryPreference to Normal, loads a small XLSX workbook with those options, optionally modifies the workbook, and saves it, demonstrating faster processing for small Excel files.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists to prevent FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook.
            Workbook workbook = new Workbook(inputPath);

            // (Optional) Perform any required operations on the workbook here.

            // Save the workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
