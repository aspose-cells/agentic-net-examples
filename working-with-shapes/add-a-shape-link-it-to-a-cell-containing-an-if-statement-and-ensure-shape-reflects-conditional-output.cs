// Title: Create a rectangle shape linked to an IF formula cell and set its text dynamically using Aspose.Cells for .NET (C#)
// Common Searches: how to set a shape's text from a cell formula result using Aspose.Cells C# | Aspose.Cells C# update rectangle shape after evaluating an IF formula | link Excel shape to a cell value and reflect conditional output with Aspose.Cells | populate shape text with the result of an IF statement in Aspose.Cells .NET | dynamic shape labeling based on worksheet formula evaluation in C#
// Tags: Aspose.Cells rectangle shape from cell value | C# set shape text using evaluated formula | conditional shape text binding Aspose.Cells | link Excel shape to IF formula result .NET | dynamic shape labeling based on worksheet cell

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, assigns a numeric value to B1, adds an IF formula in A1 that returns "High" or "Low", recalculates the worksheet, inserts a rectangle shape at row 2 column 2, sets the shape's text to the evaluated result of A1, formats the text, and saves the file as ShapeLinkedToCell.xlsx.
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

            // Set a value in B1 that will be used by the IF formula
            sheet.Cells["B1"].PutValue(15);

            // Place an IF formula in A1: if B1 > 10 then "High" else "Low"
            sheet.Cells["A1"].Formula = "=IF(B1>10,\"High\",\"Low\")";

            // Recalculate formulas so A1 contains the evaluated result
            workbook.CalculateFormula();

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper left row, upper left column, top offset, left offset, height, width
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 0, 0, 100, 200);

            // Set the shape's text to reflect the conditional output from cell A1
            shape.Text = sheet.Cells["A1"].StringValue;

            // Optional: format the shape's text
            shape.Font.Color = Color.Blue;
            shape.Font.Size = 12;

            // Define output file path
            string outputPath = "ShapeLinkedToCell.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
