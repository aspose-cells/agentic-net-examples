// Title: Configure a rectangle shape to follow cell H5 resizing by setting Placement to Move in Aspose.Cells for .NET
// AI Prompts: Generate C# code that adds a rectangle shape anchored to cell H5 and sets its Placement to Move so the shape resizes with the cell using Aspose.Cells. | Show how to link a shape to a specific worksheet cell and enable the move‑with‑cells behavior in Aspose.Cells for .NET. | Provide an example of setting the Placement property of a shape to Move for automatic alignment with cell changes in an Excel workbook.
// Common Searches: Aspose.Cells C# set shape placement to Move for cell H5 | How to make a shape move with its anchored cell in Aspose.Cells .NET | Anchor rectangle shape to a specific cell and enable move‑with‑cells in Excel using Aspose.Cells
// Tags: Aspose.Cells shape placement move | rectangle shape anchored H5 Aspose.Cells | C# shape move‑with‑cells Excel | link shape to worksheet cell Aspose.Cells | automatic shape alignment with cell resizing

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// // This program creates a new workbook, adds a rectangle shape anchored to cell H5, sets its Placement to Move so the shape moves with the cell when resized, and saves the workbook as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape anchored to cell H5 (row 4, column 7)
            // Parameters: type, upperLeftRow, upperLeftColumn, topOffset, leftOffset, height, width
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 4, 7, 0, 0, 60, 120);

            // Make the shape move with the cell when the cell is resized or moved
            shape.Placement = PlacementType.Move;

            // Save the workbook
            workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
