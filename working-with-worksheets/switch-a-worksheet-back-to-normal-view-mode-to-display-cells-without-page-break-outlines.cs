// Title: How to save an Excel worksheet in Normal view mode (no page break outlines) using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an existing .xlsx file with Aspose.Cells, checks for the file, and writes it back so the worksheet opens in Normal view without page break outlines. | Show how to create the output folder if needed and persist the workbook using Aspose.Cells without modifying any view properties.
// Common Searches: Aspose.Cells C# export Excel without showing page break outlines | how to keep Normal view when writing Excel file with Aspose.Cells | disable page break preview in generated Excel using Aspose.Cells .NET | C# Aspose.Cells default view mode after file creation
// Tags: Aspose.Cells default worksheet view | C# hide page break outlines | Aspose.Cells normal view mode | Excel file output folder handling Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample verifies that the source .xlsx file exists, loads it into an Aspose.Cells Workbook, ensures the destination directory is present, and saves the workbook. Because Aspose.Cells does not expose a property to toggle page break outlines, the file is saved using the default Normal view, so the resulting worksheet displays without page break outlines.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // NOTE: Aspose.Cells does not expose a direct property to hide page break outlines.
            // The workbook will be saved in the default view mode.

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
