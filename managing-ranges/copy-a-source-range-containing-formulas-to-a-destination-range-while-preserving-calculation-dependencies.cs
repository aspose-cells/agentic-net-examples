// Title: Copy a formula‑containing range to another location while preserving calculation dependencies using Aspose.Cells for .NET
// Common Searches: Aspose.Cells copy range with formulas preserving cell references in C# | How to duplicate a range of formula cells and keep dependencies in .NET | Copy Excel range with formulas using PasteOptions in Aspose.Cells C# | Preserve formula links when moving cells to a new range with Aspose.Cells | Recalculate workbook after copying formula range with Aspose.Cells | Copy A1:C10 to E1 using Aspose.Cells while maintaining calculations
// Tags: Aspose.Cells PasteOptions All copy | range.Copy preserving dependencies | duplicate Excel range maintaining calculations | recalculate formulas after range copy | preserve cell reference integrity Aspose.Cells | copy formula range to another worksheet .NET

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The sample loads an existing workbook, defines a source range (A1:C10) that contains formulas, creates a destination range of identical size starting at E1, and copies the source to the destination using PasteOptions.PasteType.All to retain formulas, values, and formatting. It then recalculates all formulas to update dependent cells and saves the workbook to a new file, with basic file‑existence validation and exception handling.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the source range that contains formulas (e.g., A1:C10)
            AsposeRange sourceRange = sheet.Cells.CreateRange("A1", "C10");

            // Determine size of the source range
            int rowCount = sourceRange.RowCount;
            int columnCount = sourceRange.ColumnCount;

            // Define the destination range start cell (e.g., E1) with the same size as source
            AsposeRange destRange = sheet.Cells.CreateRange(0, 4, rowCount, columnCount);

            // Set paste options to preserve formulas, values, and styles
            PasteOptions pasteOptions = new PasteOptions
            {
                PasteType = PasteType.All // copies everything including formulas, values, formats
            };

            // Copy the source range to the destination range while preserving calculation dependencies
            sourceRange.Copy(destRange, pasteOptions);

            // Recalculate the workbook to ensure all dependent formulas are updated
            workbook.CalculateFormula();

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
