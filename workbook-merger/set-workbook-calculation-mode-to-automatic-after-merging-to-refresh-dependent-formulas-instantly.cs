// Title: Set workbook calculation mode to Automatic after merging cells with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set calculation mode to automatic after merging a cell range | How to force formula recalculation after merging cells using Aspose.Cells for .NET | C# merge cells and refresh dependent formulas in an Excel workbook with Aspose.Cells | Enable automatic workbook calculation in Aspose.Cells after a cell merge operation | Aspose.Cells calculate formulas programmatically after merging cells in C#
// Tags: Aspose.Cells set calculation mode automatic | merge cells recalculate formulas Aspose.Cells | Workbook.CalculateFormula after cell merge | C# Aspose.Cells workbook calculation settings | Excel cell merge refresh dependent formulas

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel file, merges a specified range of cells on the first worksheet, switches the workbook's calculation mode to Automatic, forces an immediate formula recalculation, and saves the updated workbook, while handling missing files and ensuring the output directory exists.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            var workbook = new Workbook(inputPath);

            // Get the first worksheet
            var sheet = workbook.Worksheets[0];

            // Merge a range of cells (example: B2:D4)
            // Parameters: firstRow, firstColumn, totalRows, totalColumns (zero‑based indexing)
            sheet.Cells.Merge(1, 1, 3, 3);

            // Force immediate recalculation of all formulas
            workbook.CalculateFormula();

            // Ensure the output directory exists
            var outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
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
