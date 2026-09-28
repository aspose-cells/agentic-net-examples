// Title: Create a rectangle shape linked to a cell with a nested SUM‑IF formula and verify the displayed value using Aspose.Cells for .NET (C#)
// AI Prompts: Add a rectangle shape to a worksheet, set its Text property to "=B2", call workbook.CalculateFormula(), then read shape.Text and compare it to sheet.Cells["B2"].Value to confirm they are identical. | Populate cells A1‑A3, assign a nested SUM and IF formula to B2, link a shape’s text to that cell, force formula evaluation, and output a message indicating whether the shape reflects the final calculated result.
// Common Searches: Aspose.Cells C# link rectangle shape text to a cell formula | display calculated cell value inside a shape using Aspose.Cells | verify shape text matches cell result after CalculateFormula in .NET | C# example adding shape linked to cell with SUM and IF formula Aspose.Cells
// Tags: add rectangle shape linked to cell formula Aspose.Cells | nested SUM IF formula evaluation Aspose.Cells | shape text reflects cell value after CalculateFormula | C# workbook calculation before reading shape text | verify shape displays evaluated cell result Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, fills cells A1‑A3, sets a nested SUM/IF formula in B2, adds a rectangle shape whose Text is "=B2", forces formula calculation, then compares the shape's displayed text with the evaluated cell value and saves the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate cells used in the nested formula
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["A2"].PutValue(10);
            sheet.Cells["A3"].PutValue(15);

            // Cell B2 contains a nested formula: sum of A1:A3 plus conditional double of A1
            sheet.Cells["B2"].Formula = "=SUM(A1:A3) + IF(A1>0, A1*2, 0)";

            // Add a rectangle shape and link its text to cell B2
            // Parameters: type, upper left row, upper left column, top, left, height, width
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 5, 1, 0, 0, 100, 30);
            shape.Text = "=B2";

            // Force calculation of all formulas
            workbook.CalculateFormula();

            // Retrieve the evaluated value from the cell
            var cellValue = sheet.Cells["B2"].Value;

            // Retrieve the displayed text from the shape (will be the evaluated value after calculation)
            var shapeText = shape.Text;

            // Output values for verification
            Console.WriteLine($"Cell B2 value: {cellValue}");
            Console.WriteLine($"Shape text: {shapeText}");

            // Simple verification that the shape reflects the cell's final result
            if (shapeText == cellValue?.ToString())
            {
                Console.WriteLine("Verification passed: Shape reflects the cell's final result.");
            }
            else
            {
                Console.WriteLine("Verification failed: Shape does not match the cell's result.");
            }

            // Save the workbook (optional)
            workbook.Save("ShapeLinkedToCell.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
