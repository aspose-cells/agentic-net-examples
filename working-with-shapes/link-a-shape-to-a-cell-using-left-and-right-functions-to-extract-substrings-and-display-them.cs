// Title: Link a rectangle shape’s text to a cell using LEFT and RIGHT functions with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# assign an Excel formula to a shape’s text | How to use LEFT and RIGHT functions in a shape text formula with Aspose.Cells | Link a rectangle shape to cell A1 and show part of its value in .NET | Programmatically concatenate parts of a cell value in a shape using Aspose.Cells | Set shape text to =LEFT(A1,6)&RIGHT(A1,4) with Aspose.Cells for C#
// Tags: shape text formula linking Aspose.Cells | LEFT RIGHT functions in shape content C# | rectangle shape cell reference binding .NET | cell substring extraction for shape display Aspose.Cells | Aspose.Cells shape text concatenation example

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, writes "AsposeCellsDemo" to cell A1, adds a rectangle shape at row 2 column 2, and sets the shape's Text property to the formula "=LEFT(A1,6)&RIGHT(A1,4)" so the shape displays the first six and last four characters of A1. The workbook is saved as ShapeLinked.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Put sample text into cell A1
            sheet.Cells["A1"].PutValue("AsposeCellsDemo");

            // Add a rectangle shape to the worksheet
            // Parameters: type, upperLeftRow, upperLeftColumn, top, left, height, width
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 0, 0, 100, 200);

            // Link the shape's text to cell A1 using LEFT and RIGHT functions
            // This extracts the first 6 characters and the last 4 characters, then concatenates them
            shape.Text = "=LEFT(A1,6)&RIGHT(A1,4)";

            // Save the workbook
            workbook.Save("ShapeLinked.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
