// Title: Link a rectangle shape to a TODAY() date cell and display it with a custom date format using Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a workbook, sets cell A1 to =TODAY() with a custom "dd-MMM-yyyy" format, adds a rectangle shape, links the shape's text to A1, and saves the file. | Show how to configure a shape's Placement property to MoveAndSize so the shape moves and resizes together with its linked date cell in Aspose.Cells. | Demonstrate applying blue, 12‑point bold font styling to a shape that displays a linked date value.
// Common Searches: Aspose.Cells example linking shape text to a date formula cell | C# how to set a custom date format for a cell and reflect it in a linked shape | MoveAndSize placement for shapes linked to cells in Aspose.Cells .NET | Formatting font of a shape that shows a linked cell value in Aspose.Cells | Saving a workbook with a rectangle shape linked to TODAY() cell using Aspose.Cells
// Tags: link shape text to cell formula Aspose.Cells | custom date format dd-MMM-yyyy Aspose.Cells | shape placement MoveAndSize .NET | rectangle shape font styling blue bold Aspose.Cells | date cell TODAY() linked shape example

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, assigns =TODAY() to cell A1 with the custom format "dd-MMM-yyyy", adds a rectangle shape, links the shape's text to the date cell, sets the shape's placement to MoveAndSize, styles the shape's font in blue, 12‑pt bold, and saves the result as ShapeLinkedToDate.xlsx.
class ShapeLinkToDateCell
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Set a date formula in cell A1 (e.g., TODAY())
            Cell dateCell = sheet.Cells["A1"];
            dateCell.Formula = "=TODAY()";

            // Apply a custom date format to the cell (e.g., "dd-MMM-yyyy")
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Custom = "dd-MMM-yyyy";
            dateCell.SetStyle(dateStyle);

            // Add a rectangle shape to the worksheet.
            // Parameters: type, upperLeftRow, upperLeftColumn, topOffset, leftOffset, height, width
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle,
                5,   // upper left row
                5,   // upper left column
                0,   // top offset (pixels)
                0,   // left offset (pixels)
                30,  // height (pixels)
                100  // width (pixels)
            );

            // Set the shape's placement so it moves and sizes with the linked cell
            shape.Placement = PlacementType.MoveAndSize;

            // Link the shape's text to the cell by using a formula reference
            shape.Text = "=A1";

            // Optional: format the shape's text appearance
            shape.Font.Color = Color.Blue;
            shape.Font.Size = 12;
            shape.Font.IsBold = true;

            // Save the workbook to a file
            workbook.Save("ShapeLinkedToDate.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
