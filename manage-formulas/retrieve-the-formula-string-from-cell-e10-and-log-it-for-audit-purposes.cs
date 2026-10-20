// Title: Read and log the formula from cell E10 in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells get formula string from a specific cell | How to extract and print Excel cell formula with Aspose.Cells .NET | Logging cell formulas for audit using Aspose.Cells in C# | Retrieve formula of E10 using Aspose.Cells API | Read Excel formula from cell E10 without opening Excel
// Tags: Aspose.Cells read formula property C# | read Excel formula via Aspose.Cells API | log worksheet formula for compliance | retrieve formula from designated cell Aspose.Cells | output cell formula to console C#

using Aspose.Cells;
using System;
using System.IO;

// The example verifies the input.xlsx file, loads it with Aspose.Cells, accesses the first worksheet, reads the Formula property of cell E10, and writes the formula text to the console while handling CellsException and generic errors.
class Program
{
    static void Main()
    {
        const string filePath = "input.xlsx";

        try
        {
            // Verify that the input file exists before attempting to load it
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            // Load the workbook (no password). If the file is password‑protected,
            // a CellsException will be thrown and handled below.
            var workbook = new Workbook(filePath);

            // Access the first worksheet
            var worksheet = workbook.Worksheets[0];

            // Retrieve cell E10
            var cell = worksheet.Cells["E10"];

            // Get the formula string from the cell
            string formula = cell.Formula;

            // Log the formula for audit purposes
            Console.WriteLine($"Formula in E10: {formula}");
        }
        catch (CellsException ex)
        {
            // Handle password‑protected or other Cells-specific errors
            Console.WriteLine($"Aspose.Cells error: {ex.Message}");
        }
        catch (Exception ex)
        {
            // Handle any other unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
