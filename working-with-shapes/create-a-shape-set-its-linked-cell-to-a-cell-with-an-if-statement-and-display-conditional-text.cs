// Title: Create a rectangle shape linked to an IF‑formula cell to display conditional text with Aspose.Cells for .NET
// AI Prompts: Generate C# code that uses Aspose.Cells to insert a rectangle shape, set its LinkedCell to a cell containing an IF formula, and center the shape's text. | Show how to bind a shape's text to the result of a conditional formula in a worksheet using Aspose.Cells for .NET. | Write a C# example that creates a workbook, adds an IF formula, links a rectangle shape to that cell, aligns the text, and saves the file.
// Common Searches: how to link a shape to a cell with an IF statement using Aspose.Cells C# | Aspose.Cells .NET display result of IF formula inside a rectangle shape | center text in a shape linked to a formula cell Aspose.Cells | save workbook after linking shape to a formula cell with Aspose.Cells
// Tags: add rectangle shape linkedcell Aspose.Cells | display IF formula result in shape text | center shape text alignment Aspose.Cells | save workbook with linked shape Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program creates a new workbook, writes an IF formula in cell A1 based on B1, adds a rectangle shape, links the shape's text to cell A1 so it shows the formula result, centers the text within the shape, and saves the workbook as ShapeLinkedCell.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Set a value that will be used in the IF formula
            sheet.Cells["B1"].PutValue(15); // Change this value to test different outcomes

            // Place an IF formula in cell A1
            sheet.Cells["A1"].Formula = "=IF(B1>10,\"High\",\"Low\")";

            // Add a rectangle shape to the worksheet
            // Parameters: drawing type, upper left row, upper left row offset,
            // upper left column, upper left column offset, width, height
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                5,   // upper left row
                0,   // upper left row offset
                5,   // upper left column
                0,   // upper left column offset
                120, // width
                30   // height
            );

            // Link the shape's text to cell A1 so it displays the result of the IF formula
            shape.LinkedCell = "A1";

            // Center the text inside the shape
            shape.TextHorizontalAlignment = TextAlignmentType.Center;
            shape.TextVerticalAlignment = TextAlignmentType.Center;

            // Save the workbook to a file
            workbook.Save("ShapeLinkedCell.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
