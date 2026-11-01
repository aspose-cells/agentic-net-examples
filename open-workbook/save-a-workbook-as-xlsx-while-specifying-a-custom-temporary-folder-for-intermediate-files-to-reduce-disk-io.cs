// Title: Save an Excel workbook as XLSX using Aspose.Cells for .NET with MemoryPreference to minimize disk I/O
// Common Searches: how to use MemoryPreference in Aspose.Cells to save large workbook as xlsx in C# | C# Aspose.Cells example for reducing disk I/O during workbook save | setting temporary file usage when saving Excel files with Aspose.Cells | Aspose.Cells save workbook to xlsx with custom memory settings in .NET | optimizing Excel generation performance using Aspose.Cells MemorySetting
// Tags: Aspose.Cells MemoryPreference workbook saving | save workbook as xlsx with temporary files | disk I/O optimization for Excel generation | configure temporary folder for Aspose.Cells output | memory management in Aspose.Cells C#

using Aspose.Cells;
using System;
using System.IO;

// The sample creates a new Workbook, sets Settings.MemorySetting to MemoryPreference so Aspose.Cells uses temporary files instead of keeping everything in memory, ensures the output directory exists, and saves the workbook as an XLSX file, thereby reducing disk I/O during the save operation.
class Program
{
    static void Main()
    {
        try
        {
            // Define output file path
            string outputPath = Path.Combine(Environment.CurrentDirectory, "Output.xlsx");

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Create a new workbook
            Workbook workbook = new Workbook();

            // Set memory preference to use temporary files when needed
            workbook.Settings.MemorySetting = MemorySetting.MemoryPreference;

            // Save the workbook as XLSX
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
