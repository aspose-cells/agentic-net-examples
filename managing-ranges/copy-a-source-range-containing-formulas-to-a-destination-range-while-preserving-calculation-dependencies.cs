// Title: How to copy a formula‑containing range to a new area and keep Excel calculation dependencies with Aspose.Cells for .NET
// Common Searches: Aspose.Cells copy range with formulas and keep cell references in C# | C# copy Excel range preserving formula dependencies using Aspose.Cells | How to duplicate a block of cells with formulas to another location in a workbook with Aspose.Cells | Copy cells A1:C10 to E1 while maintaining formula links in Aspose.Cells .NET | Preserve Excel formula relationships when programmatically copying ranges in C#
// Tags: Aspose.Cells range.Copy preserving formulas | copy Excel range to new area C# | maintain calculation dependencies during range copy | duplicate cell block with formula links | C# workbook range copy example

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The example loads a workbook, defines a source range (A1:C10) that contains formulas, creates a destination range starting at E1, and uses Aspose.Cells' Range.Copy method to duplicate styles, values, and formulas while preserving all calculation dependencies, then saves the result to output.xlsx.
class Program
{
    static void Main()
    {
        const string sourcePath = "source.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(sourcePath);

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Define the source range that contains formulas (A1:C10)
            int srcFirstRow = 0;      // Row 1 (zero‑based)
            int srcFirstColumn = 0;   // Column A (zero‑based)
            int totalRows = 10;       // Number of rows to copy
            int totalColumns = 3;     // Number of columns to copy

            // Define the destination start cell (E1)
            int destFirstRow = 0;     // Row 1 (zero‑based)
            int destFirstColumn = 4;  // Column E (zero‑based)

            // Create source and destination ranges
            AsposeRange srcRange = sheet.Cells.CreateRange(srcFirstRow, srcFirstColumn, totalRows, totalColumns);
            AsposeRange destRange = sheet.Cells.CreateRange(destFirstRow, destFirstColumn, totalRows, totalColumns);

            // Copy the range (style, value, and formula are copied by default)
            srcRange.Copy(destRange);
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors during loading or copying
            Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
            return;
        }

        try
        {
            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? string.Empty;
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook with the copied range
            Workbook workbook = new Workbook(sourcePath); // Reload to ensure changes are saved
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any errors during saving
            Console.WriteLine($"An error occurred while saving the workbook: {ex.Message}");
        }
    }
}
