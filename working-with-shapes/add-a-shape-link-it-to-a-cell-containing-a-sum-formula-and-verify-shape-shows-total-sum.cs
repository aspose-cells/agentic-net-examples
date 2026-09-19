// Title: Create a rectangle shape linked to a SUM formula cell and validate its displayed value using Aspose.Cells for .NET
// AI Prompts: Generate C# code that adds a rectangle shape to a worksheet, assigns the shape's caption to the result of a SUM(A1:A4) formula, and confirms the caption matches the calculated total. | Write an Aspose.Cells example that populates cells, inserts a SUM formula, calculates it, links a shape's text to the formula result, verifies the link, and saves the workbook.
// Common Searches: Aspose.Cells C# set shape caption to result of SUM formula | how to bind a shape's displayed text to a calculated cell value in Aspose.Cells | validate that a rectangle shape shows the same total as a SUM cell in a .NET workbook | link shape text to cell containing SUM(A1:A4) using Aspose.Cells
// Tags: add rectangle shape Aspose.Cells | assign shape caption from cell formula C# | check shape text equals calculated sum Aspose.Cells | persist linked shape in .xlsx workbook

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;

// The example creates a workbook, fills cells A1‑A4 with numbers, inserts a SUM formula in A5, calculates the formula, adds a rectangle shape, sets the shape's caption to the computed sum, verifies the caption matches the cell value, and saves the file as ShapeLinkedToSum.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate cells A1:A4 with sample numeric data
            worksheet.Cells["A1"].PutValue(10);
            worksheet.Cells["A2"].PutValue(20);
            worksheet.Cells["A3"].PutValue(30);
            worksheet.Cells["A4"].PutValue(40);

            // Insert a SUM formula in cell A5 that adds the values from A1 to A4
            worksheet.Cells["A5"].Formula = "=SUM(A1:A4)";

            // Calculate all formulas so that A5 now contains the computed total
            workbook.CalculateFormula();

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper left row, upper left column, top offset, left offset, width, height
            Shape shape = worksheet.Shapes.AddShape(MsoDrawingType.Rectangle, 5, 0, 5, 0, 200, 50);

            // Link the shape's displayed text to the value of the SUM cell (A5)
            shape.Text = worksheet.Cells["A5"].StringValue; // After calculation, this holds the total sum

            // Verification: ensure the shape's text matches the computed sum in cell A5
            double cellSum = worksheet.Cells["A5"].DoubleValue;
            double shapeSum;
            bool isValid = double.TryParse(shape.Text, out shapeSum) && shapeSum == cellSum;

            if (isValid)
            {
                Console.WriteLine($"Verification passed: Shape displays the correct total sum ({shapeSum}).");
            }
            else
            {
                Console.WriteLine("Verification failed: Shape text does not match the cell sum.");
            }

            // Save the workbook (optional, demonstrates that the shape is persisted)
            workbook.Save("ShapeLinkedToSum.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
