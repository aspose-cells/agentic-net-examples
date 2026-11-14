// Title: Create a rectangle shape linked to a SUM formula cell and validate the displayed total with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# add rectangle shape and link its text to a cell formula | How to bind a shape to the result of a SUM formula in an Excel workbook using Aspose.Cells | Verify that shape text matches the calculated cell value after calling workbook.CalculateFormula in C# | Set shape text to reference a cell and retrieve the evaluated value with Aspose.Cells .NET | Link an Excel shape to cell A1 formula and read the displayed total using Aspose.Cells
// Tags: rectangle shape addition Aspose.Cells | shape text formula binding C# | workbook formula calculation Aspose.Cells | retrieve evaluated shape text .NET | shape value verification against cell Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, fills cells B2‑B5 with numbers, assigns a SUM formula to A1, adds a rectangle shape, sets the shape's text to "=A1", forces formula calculation, reads the evaluated text from both the shape and the cell, compares them for verification, and saves the workbook as ShapeLinkedToSum.xlsx.
class ShapeLinkExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate numbers in B2:B5
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["B5"].PutValue(40);

            // Set A1 to contain a SUM formula of the numbers above
            sheet.Cells["A1"].Formula = "=SUM(B2:B5)";

            // Add a rectangle shape to the worksheet
            // Parameters: upper left row, upper left column, top offset, left offset, height, width
            int upperLeftRow = 2;      // Row index (0‑based) where shape starts
            int upperLeftColumn = 0;   // Column index (0‑based) where shape starts
            int top = 5;               // Pixels from the top of the cell
            int left = 5;              // Pixels from the left of the cell
            int height = 50;           // Height in pixels
            int width = 150;           // Width in pixels
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle,
                                                upperLeftRow, upperLeftColumn,
                                                top, left, height, width);

            // Link the shape's text to cell A1 by setting a formula as the shape's text
            shape.Text = "=A1";

            // Force calculation of all formulas (including the shape's linked formula)
            workbook.CalculateFormula();

            // Retrieve the evaluated value from the shape and the cell for verification
            string shapeValue = shape.Text;                     // Should contain the calculated sum
            string cellValue = sheet.Cells["A1"].StringValue;   // Should be the same as shapeValue

            // Simple verification output
            Console.WriteLine("Cell A1 value: " + cellValue);
            Console.WriteLine("Shape text value: " + shapeValue);
            Console.WriteLine("Verification: " + (shapeValue == cellValue ? "PASS" : "FAIL"));

            // Save the workbook to a file
            string outputFile = "ShapeLinkedToSum.xlsx";
            workbook.Save(outputFile);
            Console.WriteLine("Workbook saved to: " + Path.GetFullPath(outputFile));
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
