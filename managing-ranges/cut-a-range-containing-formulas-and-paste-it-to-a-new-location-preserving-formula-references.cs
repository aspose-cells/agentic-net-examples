// Title: Cut a range with formulas and paste it to another location while preserving references using Aspose.Cells for .NET
// Common Searches: how to cut a range of cells that contain formulas and move them with Aspose.Cells in C# | preserving formula references when moving a block of cells using Aspose.Cells .NET | cut and paste a range with formulas and clear the source in Aspose.Cells C# example | copy a formula range to a new area and adjust cell references automatically Aspose.Cells | move cells with formulas to a different worksheet location using Aspose.Cells API
// Tags: cut range with formulas Aspose.Cells | copy range preserving references Aspose.Cells | clear source range after copy Aspose.Cells | adjust formulas on range move Aspose.Cells | same-size destination range Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using CellsRange = Aspose.Cells.Range;

// The example loads a workbook, defines a source range containing formulas (e.g., B2:D5), creates an equally sized destination range (e.g., G10), copies the source to the destination so that formula references are automatically updated, clears the original cells to emulate a cut operation, and saves the modified workbook.
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
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // Define the source range that contains formulas
            // Example: cells B2:D5 (zero‑based indices)
            // -------------------------------------------------
            int srcFirstRow = 1;      // Row 2
            int srcFirstColumn = 1;   // Column B
            int srcRows = 4;          // 4 rows (2,3,4,5)
            int srcColumns = 3;       // 3 columns (B,C,D)

            // Create the source range object
            CellsRange sourceRange = sheet.Cells.CreateRange(srcFirstRow, srcFirstColumn, srcRows, srcColumns);

            // -------------------------------------------------
            // Define the destination top‑left cell for the paste
            // Example: cell G10
            // -------------------------------------------------
            int destFirstRow = 9;     // Row 10
            int destFirstColumn = 6;  // Column G

            // Create the destination range with the same size as the source
            CellsRange destRange = sheet.Cells.CreateRange(destFirstRow, destFirstColumn, srcRows, srcColumns);

            // Copy the range to the new location within the same worksheet.
            // Formulas are automatically adjusted to preserve correct references.
            sourceRange.Copy(destRange);

            // Clear the original range to achieve a "cut" operation.
            sourceRange.Clear();

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
