// Title: Link a Rectangle Shape to a MID Formula Cell in Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, place the text "Aspose.Cells" in A1, apply a MID formula in B1 to extract a substring, add a rectangle shape, link the shape to B1, clear the shape's own text so it shows the formula result, and save the file as an .xlsx document.
// Keywords: Aspose.Cells | C# | shape linked cell | MID function | rectangle shape | dynamic shape text | Excel automation | linked shape formula | Aspose.Cells example | cell formula to shape
// Common Searches: Aspose.Cells link shape to cell | C# add rectangle shape linked to formula | display MID formula result in shape using Aspose.Cells | set shape text from cell formula Aspose.Cells .NET | dynamic shape label with MID function
// Developer Intent: Create a rectangle shape whose displayed text automatically reflects the result of a MID formula applied to another cell.
// Use Cases: Show a label that updates when the source string changes, without manual edits. | Build a dashboard where a shape presents a specific portion of a longer text field. | Generate printable forms that display extracted substrings directly inside shapes.
// AI Prompts: Write C# code with Aspose.Cells to add a rectangle shape linked to a cell that uses the MID function and save the workbook. | Explain how to change the start position or length parameters of the MID formula after the shape is linked so the displayed text updates dynamically. | Provide best‑practice error handling for cases where the MID function arguments exceed the source string length when linking a shape.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to create a workbook, place the text "Aspose.Cells" in A1, apply a MID formula in B1 to extract a substring, add a rectangle shape, link the shape to B1, clear the shape's own text so it shows the formula result, and save the file as an .xlsx document.
class ShapeLinkedCellMidExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Put source text in cell A1
        sheet.Cells["A1"].PutValue("Aspose.Cells");

        // Set a formula in B1 that extracts a substring using MID
        // Example: extract 5 characters starting from the 2nd character -> "spose"
        sheet.Cells["B1"].Formula = "=MID(A1,2,5)";

        // Add a rectangle shape to the worksheet (positioned at row 2, column 2)
        // Parameters: upperRow, upperColumn, width, height, lowerRow, lowerColumn
        Shape shape = sheet.Shapes.AddRectangle(2, 2, 100, 30, 2, 2);

        // Link the shape to the cell containing the MID formula (B1)
        shape.LinkedCell = "B1";

        // Optionally, set the shape's text to be empty; it will display the linked cell's value
        shape.Text = "";

        // Save the workbook to a file
        workbook.Save("ShapeLinkedCellMidExample.xlsx");
    }
}
