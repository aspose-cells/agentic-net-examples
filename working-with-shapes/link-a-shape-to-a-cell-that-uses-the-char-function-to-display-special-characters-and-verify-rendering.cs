// Title: Link a Shape to a CHAR Formula Cell and Verify Unicode Rendering – Aspose.Cells for .NET Example
// Description: Shows how to place a =CHAR(9731) formula in cell A1, add a rectangle shape, bind the shape to the cell with SetLinkedCell, refresh the shape value, read shape.Text to confirm the snowman character, and save the workbook.
// Keywords: Aspose.Cells | .NET | C# shape linking | SetLinkedCell | CHAR function | Unicode character | shape.Text | linked shape verification | Excel shape API | dynamic shape text
// Common Searches: Aspose.Cells link shape to cell | SetLinkedCell CHAR formula | retrieve text from linked shape | display Unicode character in Excel shape using Aspose | update shape after cell formula change
// Developer Intent: The developer wants to bind a worksheet shape to a cell that returns a special character via the CHAR function and verify that the shape displays the correct character.
// Use Cases: Create dashboard icons that reflect Unicode symbols defined by formulas | Automatically synchronize shape captions with cell calculations | Extract shape captions for reporting or logging purposes
// AI Prompts: Write C# code with Aspose.Cells to link a rectangle shape to cell A1 containing =CHAR(9731) and output the shape's displayed text. | Explain how the three parameters of SetLinkedCell affect linking a shape to a cell that returns a Unicode character. | Provide debugging steps when shape.Text does not update after SetLinkedCell and UpdateSelectedValue.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Shows how to place a =CHAR(9731) formula in cell A1, add a rectangle shape, bind the shape to the cell with SetLinkedCell, refresh the shape value, read shape.Text to confirm the snowman character, and save the workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Put a formula in cell A1 that uses the CHAR function to display a special character (e.g., snowman)
            Cell linkedCell = sheet.Cells["A1"];
            linkedCell.Formula = "=CHAR(9731)";

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper‑left row, upper‑left column, top offset, left offset, height, width
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 0, 0, 0, 30, 100);

            // Link the shape to the cell containing the CHAR formula
            shape.SetLinkedCell("A1", false, false);

            // Update the shape so it reflects the value of the linked cell
            shape.UpdateSelectedValue();

            // Retrieve the text displayed by the shape to verify rendering
            string shapeText = shape.Text; // shape.Text returns the displayed string
            Console.WriteLine($"Shape displays: '{shapeText}'");

            // Save the workbook (optional visual verification)
            workbook.Save("LinkedShape.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
