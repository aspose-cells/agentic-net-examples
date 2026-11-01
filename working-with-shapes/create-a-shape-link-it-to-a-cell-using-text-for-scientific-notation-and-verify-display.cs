// Title: Create a rectangle shape linked to a worksheet cell and display its value in scientific notation using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that inserts a rectangle shape, sets its Text property to a TEXT formula referencing cell A1, and formats the result in scientific notation. | Show how to trigger workbook formula calculation and then read the rendered text of a shape that is bound to a cell in Aspose.Cells.
// Common Searches: how to bind a shape's text to a cell using TEXT function in Aspose.Cells C# | display cell value in scientific notation inside a rectangle shape Aspose.Cells | force formula evaluation before reading shape text Aspose.Cells .NET | Aspose.Cells example linking shape to cell with custom number format | C# Aspose.Cells shape text formula not updating until CalculateFormula
// Tags: rectangle shape linked to worksheet cell Aspose.Cells | shape text using TEXT formula scientific notation | force workbook calculation before shape text retrieval | C# Aspose.Cells shape text binding to cell

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, writes a numeric value to cell A1, adds a rectangle shape, assigns its Text property to a TEXT formula that formats A1 in scientific notation, forces formula calculation, outputs the shape's displayed text, and saves the file as ShapeLinkedToCell.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Put a numeric value in cell A1
            Cell cellA1 = sheet.Cells["A1"];
            cellA1.PutValue(12345.6789);

            // Add a rectangle shape to the worksheet
            // Parameters: drawing type, upper left row, upper left column, top, left, height, width
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 0, 0, 100, 50, 200);

            // Link the shape's text to cell A1 using TEXT function with scientific notation
            shape.Text = "=TEXT(A1,\"0.00E+00\")";

            // Force calculation of all formulas in the workbook
            workbook.CalculateFormula();

            // Verify the displayed text of the shape
            Console.WriteLine("Shape displayed text: " + shape.Text); // Expected: 1.23E+04 (or similar)

            // Save the workbook (optional, just to demonstrate lifecycle usage)
            workbook.Save("ShapeLinkedToCell.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
