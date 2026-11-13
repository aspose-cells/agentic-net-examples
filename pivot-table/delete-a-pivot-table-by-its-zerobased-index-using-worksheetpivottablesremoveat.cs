// Title: Remove a pivot table at a specific zero‑based index from a worksheet using Aspose.Cells for .NET
// Common Searches: how to remove a pivot table by index with Aspose.Cells C# | Aspose.Cells delete pivot table at position 0 in worksheet | C# code to check pivot table count before calling RemoveAt | saving workbook after removing pivot table using Aspose.Cells | verify input Excel file exists before loading with Aspose.Cells
// Tags: remove pivot table by index Aspose.Cells | Worksheet.PivotTables.RemoveAt method | pivot table count validation Aspose.Cells | output workbook path handling Aspose.Cells | exception handling for Excel file operations C#

using System;
using System.IO;
using Aspose.Cells;

// The example checks that the source Excel file exists, loads it into a Workbook, accesses the first worksheet, verifies that at least one pivot table is present, and then removes the pivot table at zero‑based index 0 using Worksheet.PivotTables.RemoveAt. After removal it ensures the output directory exists, saves the workbook, and catches any runtime exceptions.
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
            var workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust if needed)
            var worksheet = workbook.Worksheets[0];

            // Ensure there is at least one pivot table before attempting removal
            if (worksheet.PivotTables.Count > 0)
            {
                // Delete the pivot table at the specified zero‑based index (0 = first pivot table)
                worksheet.PivotTables.RemoveAt(0);
                Console.WriteLine("Pivot table removed successfully.");
            }
            else
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
            }

            // Ensure the output directory exists
            var outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
