// Title: Create a rectangle shape in Aspose.Cells for .NET and display the result of an IF formula from a linked cell
// AI Prompts: Add a rectangle shape to a worksheet, calculate the workbook formulas, and set the shape's text to the evaluated value of cell A1 that contains an IF formula using C# and Aspose.Cells. | Store the address of the source cell in the shape's AlternativeText property and refresh the shape's displayed text after calling workbook.CalculateFormula in a .NET application.
// Common Searches: asp.net how to link a shape's text to a cell formula result in Aspose.Cells | c# set rectangle shape text from IF formula cell Aspose.Cells | update shape after workbook.CalculateFormula Aspose.Cells example | store linked cell address in shape AlternativeText Aspose.Cells C#
// Tags: add rectangle shape linked to cell formula Aspose.Cells | shape text binding to IF formula result C# | alternativetext store cell address Aspose.Cells | calculate formulas before assigning shape text .NET | Aspose.Cells shape and cell interaction example

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// // Demonstrates creating a workbook, placing a numeric value in B1, defining an IF formula in A1, calculating formulas, adding a rectangle shape, setting the shape's text to the evaluated result of A1, recording the linked cell address in AlternativeText, and saving the file as ShapeLinkedToIfCell.xlsx.
class ShapeLinkExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set up cells:
        // B1 will hold the numeric value used in the IF statement
        sheet.Cells["B1"].PutValue(15);

        // A1 contains an IF formula that depends on B1
        sheet.Cells["A1"].Formula = "=IF(B1>10,\"High\",\"Low\")";

        // Calculate formulas so that A1 gets the evaluated result
        workbook.CalculateFormula();

        // Add a rectangle shape to the worksheet
        // Parameters: type, upper left row, upper left column, top offset, left offset, height, width
        Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 0, 0, 0, 30, 100);

        // Set shape's text to reflect the value of the linked cell (A1)
        shape.Text = sheet.Cells["A1"].StringValue;

        // Optionally store the linked cell address in AlternativeText for reference
        shape.AlternativeText = "LinkedCell=A1";

        // Save the workbook
        workbook.Save("ShapeLinkedToIfCell.xlsx");
    }
}
