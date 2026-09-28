// Title: How to bind a rectangle shape’s text to a cell formula for live updates in Aspose.Cells for .NET
// AI Prompts: Create a rectangle shape on a worksheet and set its Text property to a cell reference so the shape displays the current formula result. | Generate code that links a shape’s displayed text to a worksheet cell, ensuring the text updates automatically when the cell’s formula recalculates.
// Common Searches: Aspose.Cells .NET bind shape text to cell formula for live result | display cell formula result in a textbox shape using Aspose.Cells | update shape text automatically when worksheet cell changes Aspose.Cells | link rectangle shape to B2 formula Aspose.Cells example
// Tags: rectangle shape text binding Aspose.Cells .NET | dynamic shape content from worksheet formula | shape text linked to cell value Aspose.Cells | C# Aspose.Cells live formula display in shape | binding shape to cell calculation result

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, writes a SUM formula in cell B2, adds a rectangle shape, sets the shape's Text property to "=B2" so it shows the live calculation result, formats the text, and saves the file as LiveFormulaShape.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Put a formula in cell B2 (example: sum of numbers)
            sheet.Cells["B2"].Formula = "=SUM(1,2,3)"; // Result will be 6

            // Add a rectangle shape (acts as a textbox) to the worksheet
            // Parameters: drawing type, upper left row, upper left column,
            // top offset (points), left offset (points), height (points), width (points)
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                1,    // upper left row
                0,    // upper left column
                0,    // top offset
                100,  // left offset
                30,   // height
                200   // width
            );

            // Link the shape's text to the cell containing the formula.
            // The shape will display the live result of B2 and update automatically.
            shape.Text = "=B2";

            // Optional: format the text inside the shape
            shape.Font.Size = 12;
            shape.Font.IsBold = true;

            // Save the workbook to a file
            string outputPath = "LiveFormulaShape.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
