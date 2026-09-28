// Title: How to anchor a comment box shape to cell G12 and keep it moving with the cell using Aspose.Cells for .NET
// AI Prompts: Create a Comment shape at cell G12, set its Placement to Move, and save the workbook with Aspose.Cells C#. | Add a linked comment box to a worksheet, configure PlacementType.MoveAndSize, and verify it follows inserted rows or columns.
// Common Searches: Aspose.Cells C# anchor comment shape to a specific cell | move comment box with cell when rows are inserted Aspose.Cells | set shape placement to Move in Aspose.Cells .NET example | link comment shape to G12 using Aspose.Cells API | how to keep a comment box aligned with a cell after column shift in C#
// Tags: comment shape placement move Aspose.Cells | comment box anchoring to worksheet cell C# | link shape to target cell Aspose.Cells | shape follows cell shifts Aspose.Cells .NET | programmatic addition of comment shape Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, adds a comment shape anchored to cell G12 on the first worksheet, sets its Placement to Move so the shape follows the cell when rows or columns shift, and saves the file as LinkedCommentBox.xlsx.
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

            // Target cell G12 (zero‑based indices)
            int targetRow = 11;   // row index for G12
            int targetColumn = 6; // column index for G12

            // Add a comment shape anchored to the target cell
            // Use MsoDrawingType.Comment for compatibility with older Aspose.Cells versions
            Shape commentBox = sheet.Shapes.AddShape(
                MsoDrawingType.Comment, // enum value for a comment box
                targetRow, targetColumn,
                0, 0,
                200, 100); // width and height in pixels

            // Set the comment text (optional)
            commentBox.Text = "This is a linked comment box.";

            // Ensure the shape moves when the cell shifts
            commentBox.Placement = PlacementType.Move; // or PlacementType.MoveAndSize

            // Define output file path
            string outputPath = "LinkedCommentBox.xlsx";

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
