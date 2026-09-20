// Title: Duplicate a textbox shape, attach the copy to a different cell, and verify matching text with Aspose.Cells for .NET (C#)
// AI Prompts: Create C# code that adds a textbox linked to cell B2, uses the AddCopy method to place a duplicate at cell D5, sets both shapes to MoveAndSize, and outputs whether their Text properties are identical. | Write a C# example that clones an existing shape on a worksheet, re‑anchors it to another cell, and compares the displayed content of the original and the clone using Aspose.Cells.
// Common Searches: aspnet cells how to copy a textbox shape to another cell in C# | c# aspose.cells compare Text property of original and copied shape | using AddCopy to duplicate a shape while preserving MoveAndSize placement | link duplicated shape to a different cell with Aspose.Cells for .NET
// Tags: AddCopy method shape copy Aspose.Cells | textbox shape anchored to cell B2 C# | MoveAndSize placement for copied shapes | compare Text property of two shapes Aspose.Cells | re‑anchor copied shape to another cell Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, adds a textbox linked to B2, duplicates it to D5 with AddCopy, sets both shapes to MoveAndSize, prints each shape's Text, checks equality, and saves the file as ShapeDuplicationResult.xlsx.
class ShapeDuplicationExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a textbox shape linked to cell B2 with some text
            int originalRow = 1;    // B2 (zero‑based index)
            int originalColumn = 1; // B2
            Shape originalShape = sheet.Shapes.AddTextBox(originalRow, originalColumn, 0, 0, 120, 30);
            originalShape.Text = "Original Shape";
            // Set the shape to move and size with cells
            originalShape.Placement = PlacementType.MoveAndSize;

            // Duplicate the original shape onto the same worksheet at a different location
            int targetRow = 4;    // D5 (zero‑based index)
            int targetColumn = 3; // D5
            Shape duplicateShape = sheet.Shapes.AddCopy(originalShape, targetRow, targetColumn, 0, 0);
            // Ensure the duplicate also moves and sizes with its linked cell
            duplicateShape.Placement = PlacementType.MoveAndSize;

            // Compare displayed contents (text) of the two shapes
            string originalText = originalShape.Text;
            string duplicateText = duplicateShape.Text;

            Console.WriteLine("Original shape text:  " + originalText);
            Console.WriteLine("Duplicate shape text: " + duplicateText);
            Console.WriteLine("Contents are equal:   " + (originalText == duplicateText));

            // Save the workbook to verify the result (optional)
            workbook.Save("ShapeDuplicationResult.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
