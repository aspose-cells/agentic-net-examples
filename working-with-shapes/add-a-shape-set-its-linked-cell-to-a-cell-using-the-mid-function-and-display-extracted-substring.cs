// Title: Link a Rectangle Shape to a MID Formula Cell in Aspose.Cells (C#)
// Description: Demonstrates how to create a workbook, place source text in A1, apply the MID function in B1, add a rectangle shape, link the shape to B1 so it displays the extracted substring, and save the file as ShapeLinkedCellMID.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# shape linked cell | MID function | rectangle shape | dynamic label | cell formula display | Excel automation | linked shape text
// Common Searches: Aspose.Cells link shape to cell with formula | C# add rectangle shape showing MID result | display cell value in shape Aspose.Cells | bind shape to calculated cell .NET | shape linked cell example Aspose
// Developer Intent: Create a rectangle shape whose displayed text updates automatically from a cell that uses the MID function.
// Use Cases: Dynamic dashboard labels that reflect parts of source strings. | Printable reports where shapes show extracted substrings without manual edits. | Interactive worksheets that use shapes as live text indicators for calculated values.
// AI Prompts: Generate C# code with Aspose.Cells to add a rectangle shape linked to a cell containing a MID formula and save the workbook. | Show how to link multiple shapes to cells using text functions (MID, LEFT, RIGHT) in Aspose.Cells for .NET. | Explain how to update a shape's displayed text when the linked cell formula result changes in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to create a workbook, place source text in A1, apply the MID function in B1, add a rectangle shape, link the shape to B1 so it displays the extracted substring, and save the file as ShapeLinkedCellMID.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Put source text in cell A1
        sheet.Cells["A1"].PutValue("Aspose.Cells Example");

        // In cell B1 use MID function to extract a substring (e.g., characters 9-14 -> "Cells")
        sheet.Cells["B1"].Formula = "=MID(A1,9,5)";

        // Add a rectangle shape to the worksheet
        // Parameters: upper left row, upper left column, upper left pixel offset X, upper left pixel offset Y, width, height
        Shape shape = sheet.Shapes.AddRectangle(2, 1, 0, 0, 200, 50);

        // Link the shape to the cell containing the MID formula (B1)
        shape.LinkedCell = "B1";

        // Optionally set some visual properties
        shape.Placement = PlacementType.FreeFloating;
        shape.Text = ""; // The shape will display the linked cell's value automatically

        // Save the workbook
        workbook.Save("ShapeLinkedCellMID.xlsx");
    }
}
