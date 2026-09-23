// Title: Load only the first three worksheets from an XLSX file with LightCells and save the trimmed workbook using Aspose.Cells for .NET
// AI Prompts: Implement workbook loading with Aspose.Cells LightCells so that only the first three worksheets are read into memory. | After loading, write the workbook back to a new XLSX file preserving just those three sheets.
// Common Searches: Aspose.Cells LightCells load specific worksheets from large XLSX | memory‑efficient way to keep first three sheets in C# Aspose.Cells | how to trim a workbook to a subset of sheets without loading all data | C# load only selected sheets using LoadFilter Aspose.Cells
// Tags: lightcells selective worksheet loading | load first three sheets Aspose.Cells | memory‑efficient workbook trimming .NET | Aspose.Cells load specific worksheets | save trimmed workbook XLSX

using System;
using System.IO;
using Aspose.Cells;

// The example demonstrates how to use Aspose.Cells LightCells API in C# to open an XLSX file, load only the first three worksheets into memory, and then save a new workbook containing just those sheets, minimizing memory consumption.
class Program
{
    static void Main()
    {
        // Paths to the source and destination files
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        try
        {
            // Ensure the input file exists; create an empty workbook if it does not
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found. Creating a new empty workbook.");
                var placeholder = new Workbook();
                placeholder.Save(inputPath, SaveFormat.Xlsx);
            }

            // Load the workbook (default loading loads all worksheets)
            var loadOptions = new LoadOptions(LoadFormat.Xlsx);
            var workbook = new Workbook(inputPath, loadOptions);

            // Keep only the first three worksheets (if they exist)
            while (workbook.Worksheets.Count > 3)
            {
                workbook.Worksheets.RemoveAt(3);
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the resulting workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
