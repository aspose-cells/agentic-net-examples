// Title: Add a rectangle shape linked to a worksheet cell with a MID formula to display a substring using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# link rectangle shape to a cell that contains a MID formula | How to show the result of a MID function inside a shape with Aspose.Cells .NET | Create a shape whose LinkedCell property points to a formula cell in a C# workbook | Display extracted substring from a cell in a linked shape using Aspose.Cells
// Tags: shape linked cell property Aspose.Cells | MID function substring shape Aspose.Cells | add rectangle shape to worksheet C# | bind shape text to formula cell Aspose.Cells | create shape displaying extracted substring .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, writes "Aspose.CellsDemo" to cell A1, sets a MID formula in B1 to extract "Cell", adds a rectangle shape, links the shape to B1 so it automatically shows the extracted substring, clears any static text, and saves the file as ShapeLinkedCell.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Put source text in cell A1
            Cell sourceCell = sheet.Cells["A1"];
            sourceCell.PutValue("Aspose.CellsDemo");

            // Set a formula in B1 that extracts a substring using MID
            // Example: extract characters 8-11 ("Cell")
            Cell formulaCell = sheet.Cells["B1"];
            formulaCell.Formula = "MID(A1,8,4)";

            // Add a rectangle shape to the worksheet and obtain the Shape object directly
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                5,   // upper-left column
                5,   // upper-left row
                0,   // column offset (pixels)
                0,   // row offset (pixels)
                100, // width (pixels)
                30   // height (pixels)
            );

            // Link the shape to cell B1 so it displays the extracted substring
            shape.LinkedCell = "B1";

            // No static text needed; the linked cell value will be shown
            shape.Text = string.Empty;

            // Define output file name
            string outputPath = "ShapeLinkedCell.xlsx";

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
