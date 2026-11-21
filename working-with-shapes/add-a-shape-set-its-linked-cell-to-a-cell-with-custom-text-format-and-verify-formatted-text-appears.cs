// Title: C# example: add a rectangle shape linked to a custom‑formatted cell and read the formatted text with Aspose.Cells
// Common Searches: Aspose.Cells C# link a shape to a cell with a custom number format | display cell's custom formatted value inside a rectangle shape using Aspose.Cells | how to read the text shown by a shape linked to a formatted cell in .NET | verify that a linked shape shows the formatted value of cell A1 with Aspose.Cells | C# example for linking shape text to a formatted worksheet cell
// Tags: add rectangle shape linked cell Aspose.Cells | custom number format shape text .NET | retrieve linked shape displayed text Aspose.Cells | shape linked cell formatting verification C# | Aspose.Cells shape text reflects cell format

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a workbook, writes a numeric value to cell A1, applies a custom number format, adds a rectangle shape, links the shape's text to the formatted cell, reads the Shape.Text property to confirm the formatted output, and saves the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Put a numeric value into cell A1
            Cell targetCell = worksheet.Cells["A1"];
            targetCell.PutValue(1234.567);

            // Apply a custom text format to the cell
            Style cellStyle = targetCell.GetStyle();
            cellStyle.Custom = "\"Value: \"#,##0.00;\"Value: \"-#,##0.00";
            targetCell.SetStyle(cellStyle);

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper left row, upper left column, top offset, left offset, height, width
            Shape shape = worksheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 0, 0, 100, 50);

            // Link the shape's text to the formatted cell A1
            shape.LinkedCell = "A1";

            // Verify that the shape's displayed text reflects the cell's custom format
            string displayedText = shape.Text; // Expected: "Value: 1,234.57"
            Console.WriteLine("Shape linked text: " + displayedText);

            // Define output file path
            string outputPath = "LinkedShapeExample.xlsx";

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
