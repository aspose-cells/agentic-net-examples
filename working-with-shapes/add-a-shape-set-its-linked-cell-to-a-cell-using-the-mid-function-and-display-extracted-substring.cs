// Title: Create a rectangle shape linked to a cell with a MID formula to display an extracted substring using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# link rectangle shape to cell formula result | How to display MID function output inside a shape with Aspose.Cells .NET | Set linked cell for a shape to show substring extracted from another cell in C# | Create shape with text bound to a cell formula in an Aspose.Cells workbook | Center text in a linked rectangle shape using Aspose.Cells for .NET
// Tags: add rectangle shape linked cell Aspose.Cells | mid formula shape text binding Aspose.Cells | link shape to cell result Aspose.Cells C# | centered shape text Aspose.Cells | extract substring with MID Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, writes "Aspose.CellsDemo" to cell A1, applies a MID formula in B1 to extract a substring, adds a rectangle shape, links the shape to cell B1 so the extracted text appears inside the shape, centers the shape's text, and saves the file as ShapeLinkedCell.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Put source text in cell A1
            sheet.Cells["A1"].PutValue("Aspose.CellsDemo");

            // Set formula in B1 to extract a substring using the MID function
            // Example: extract 4 characters starting from the 8th character ("Cell")
            sheet.Cells["B1"].Formula = "MID(A1,8,4)";

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper-left row, upper-left column,
            // upper-left row offset (pixels), upper-left column offset (pixels),
            // width (pixels), height (pixels)
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle,
                5,   // upper-left row
                5,   // upper-left column
                0,   // row offset
                0,   // column offset
                120, // width
                30); // height

            // Link the shape to cell B1 so the shape displays the extracted substring
            shape.LinkedCell = "B1";

            // Optional: center the text inside the shape
            shape.TextHorizontalAlignment = TextAlignmentType.Center;

            // Save the workbook to a file
            string outputPath = "ShapeLinkedCell.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
