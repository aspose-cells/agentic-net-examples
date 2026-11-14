// Title: Add a rectangle shape, link it to a cell with a circular reference, and handle the calculation error using Aspose.Cells for .NET
// Common Searches: asp.net add rectangle shape to worksheet and catch circular reference exception | calculate formulas with circular reference handling in Aspose.Cells C# | associate a shape with cell A1 that contains a self‑referencing formula in C# | save workbook after inserting shape and handling formula calculation errors using Aspose.Cells | use PlacementType.MoveAndSize to position a shape relative to a cell in Aspose.Cells
// Tags: add rectangle shape Aspose.Cells worksheet | link shape to cell circular reference | handle CellsException circular reference | calculate formulas with error handling Aspose.Cells | placementtype moveandsize shape positioning | save workbook after error handling Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, inserts a rectangle shape positioned at a specific cell, assigns a self‑referencing formula (=A1) to that cell, attempts to calculate all formulas while gracefully catching the CellsException for the circular reference, and finally saves the workbook with robust I/O error handling.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper left row, upper left column, top offset, left offset, height, width
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 0, 0, 80, 120);

            // Position the shape relative to cell A1 (row 0, column 0)
            shape.Placement = PlacementType.MoveAndSize;
            // SetPosition is not required; shape already positioned at (2,2). If needed, adjust using Top/Left.
            // shape.Top = 0;
            // shape.Left = 0;

            // Introduce a circular reference in cell A1 (formula refers to itself)
            Cell cell = sheet.Cells["A1"];
            cell.Formula = "=A1";

            try
            {
                // Force calculation of all formulas – this will raise an exception because of the circular reference
                workbook.CalculateFormula();
            }
            catch (CellsException)
            {
                // Handle circular reference error gracefully
                Console.WriteLine("Circular reference detected in cell A1. Calculation was aborted.");
            }
            catch (Exception ex)
            {
                // General fallback for any other errors
                Console.WriteLine($"An unexpected error occurred during calculation: {ex.Message}");
            }

            // Save the workbook to a file
            string outputPath = "ShapeWithCircularReference.xlsx";

            try
            {
                // Ensure the directory exists before saving
                string directory = Path.GetDirectoryName(outputPath);
                if (string.IsNullOrEmpty(directory))
                {
                    directory = Directory.GetCurrentDirectory();
                }

                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            // Top-level exception handling
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
