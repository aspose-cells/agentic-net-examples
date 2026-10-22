// Title: Load a massive XLSX workbook with Aspose.Cells for .NET using MemoryPreference to avoid OutOfMemoryException
// Common Searches: Aspose.Cells C# load large XLSX file without OutOfMemoryException | How to use MemoryPreference in LoadOptions for big Excel workbooks in .NET | Set MemorySetting to MemoryPreference when opening massive workbook with Aspose.Cells | Low‑memory loading of huge Excel files using Aspose.Cells LoadOptions C# | Prevent memory overflow when reading large Excel workbook with Aspose.Cells
// Tags: MemoryPreference LoadOptions Aspose.Cells | load massive XLSX workbook .NET | prevent OutOfMemoryException Aspose.Cells | low‑memory Excel loading C# | configure LoadOptions for large workbooks

using Aspose.Cells;
using System;
using System.IO;

// The example checks for the XLSX file, creates LoadOptions with MemorySetting set to MemoryPreference to reduce memory consumption, loads the massive workbook using these options, optionally processes it, and saves the result while handling any runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "MassiveWorkbook.xlsx";
            string outputPath = "MassiveWorkbook_Processed.xlsx";

            // Verify that the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                return;
            }

            // Configure load options for memory‑efficient loading of large workbooks
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                MemorySetting = MemorySetting.MemoryPreference
            };

            // Load the workbook using the configured options
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // (Optional) Perform any required operations on the workbook here
            // ...

            // Save the processed workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            // Catch and display any runtime errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
