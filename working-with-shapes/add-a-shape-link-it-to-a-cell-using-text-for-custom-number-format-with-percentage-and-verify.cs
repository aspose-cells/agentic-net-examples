// Title: Add a rectangle shape linked to a cell with a custom 0.00% number format using Aspose.Cells for .NET
// AI Prompts: Generate C# code that inserts a rectangle shape, sets its Text property to a formula referencing cell B2, and ensures the shape displays the cell's custom percentage format. | Write a snippet that applies the custom number format "0.00%" to a cell, links a shape's text to that cell via a formula, triggers workbook calculation, and prints the shape's displayed text. | Create a program that places a shape at a specific row/column, binds its text to a formatted cell, recalculates all formulas, and saves the workbook.
// Common Searches: how to bind a shape's text to a cell value in Aspose.Cells .NET | display cell's custom percentage format inside a rectangle shape using Aspose.Cells | link shape to cell formula and recalculate workbook Aspose.Cells C# example | add rectangle shape at C3 and show B2 formatted as percentage Aspose.Cells
// Tags: rectangle shape insertion Aspose.Cells | shape text formula linking Aspose.Cells | cell custom percentage format Aspose.Cells | workbook formula calculation Aspose.Cells | shape text verification Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, writes a numeric value to B2, applies a custom 0.00% format, adds a rectangle shape at C3, links the shape's text to B2 using a formula, recalculates formulas so the shape shows "12.34%", prints the displayed text, and saves the file as LinkedShape.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Put a numeric value into cell B2 (12.34%)
            Cell targetCell = sheet.Cells["B2"];
            targetCell.PutValue(0.1234);

            // Apply a custom number format for percentage (e.g., 0.00%)
            Style style = targetCell.GetStyle();
            style.Custom = "0.00%";
            targetCell.SetStyle(style);

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper left row, upper left column, top, left, height, width
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                2,    // upper left row (zero‑based, row 3)
                2,    // upper left column (zero‑based, column C)
                0,    // top offset in pixels
                0,    // left offset in pixels
                100,  // height in points
                200   // width in points
            );

            // Link the shape's text to the cell using a formula.
            // The shape will display the formatted value of B2 (12.34%).
            shape.Text = "=B2";

            // Calculate all formulas so the linked shape reflects the cell's value
            workbook.CalculateFormula();

            // Verify the shape's displayed text
            string displayedText = shape.Text; // Should be "12.34%" after calculation
            Console.WriteLine($"Shape text: {displayedText}");

            // Save the workbook to verify the result in Excel
            workbook.Save("LinkedShape.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
