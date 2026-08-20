// Title: C# – Link a Comment Box Shape to Cell G12 with Aspose.Cells for .NET
// Description: Demonstrates how to add a rectangle comment box to cell G12, set its text, assign the LinkedCell property ("$G$12") so the shape moves with the cell, and save the workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# shape LinkedCell | comment box anchor cell Aspose.Cells | AddShape rectangle Aspose.Cells | link shape to cell .NET | shape moves with cell Aspose | Aspose.Cells example LinkedCell property
// Common Searches: Aspose.Cells link shape to cell C# | How to anchor a comment box to a cell in Aspose.Cells | Set LinkedCell for a rectangle shape Aspose.Cells | Move shape with cell when rows are inserted Aspose | C# Aspose.Cells shape positioning example
// Developer Intent: The developer needs to attach a comment‑box shape to cell G12 so the shape stays aligned with the cell when rows or columns are added, removed, or resized.
// Use Cases: Create persistent annotations that remain aligned with data cells after worksheet layout changes. | Generate reports where each comment box is tied to a specific cell for dynamic positioning. | Build interactive Excel-like interfaces where shapes follow their reference cells during editing.
// AI Prompts: Show C# code to link any Aspose.Cells shape to a specific cell and keep it anchored after inserting rows or columns. | Provide an example that updates the LinkedCell property for multiple shapes based on a list of addresses. | Explain how to set top/left offsets for a shape while still using the LinkedCell property in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to add a rectangle comment box to cell G12, set its text, assign the LinkedCell property ("$G$12") so the shape moves with the cell, and save the workbook using Aspose.Cells for .NET.
class LinkCommentBoxToCell
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape positioned at cell G12 (row index 11, column index 6)
            // Signature: AddShape(MsoDrawingType shapeType, int upperLeftRow, int upperLeftColumn,
            //                     int top, int left, int height, int width)
            Shape commentBox = worksheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                11,                       // upper left row (G12)
                6,                        // upper left column (G12)
                0,                        // top offset
                0,                        // left offset
                100,                      // height
                50);                      // width

            // Set the text displayed inside the shape
            commentBox.Text = "Sample comment";

            // Link the comment box to cell G12 so it moves with the cell
            commentBox.LinkedCell = "$G$12";

            // Save the workbook
            workbook.Save("LinkedCommentBox.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
