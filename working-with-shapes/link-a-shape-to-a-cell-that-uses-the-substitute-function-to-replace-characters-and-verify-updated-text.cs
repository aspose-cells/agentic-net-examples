// Title: C# Aspose.Cells: Link TextBox Shape to Cell, Apply SUBSTITUTE Formula, Update Shape Text
// Description: Demonstrates how to create a workbook, add a TextBox shape, link it to cell B2, set a SUBSTITUTE formula that changes "Hello World" to "Hello Aspose", recalculate formulas, refresh the shape’s displayed text, verify the result, and save the file as LinkedShapeSubstitute.xlsx.
// Keywords: Aspose.Cells | C# shape linking | SetLinkedCell | UpdateSelectedValue | SUBSTITUTE function | Excel textbox shape | linked cell formula | recalculate formulas | .NET Aspose.Cells example | dynamic shape text
// Common Searches: Aspose.Cells link textbox to cell | SetLinkedCell C# example | Update shape text after formula Aspose.Cells | SUBSTITUTE formula in linked cell | Refresh linked shape value .NET | How to bind shape to cell in Aspose.Cells
// Developer Intent: Link a TextBox shape to a cell containing a SUBSTITUTE formula and refresh the shape’s text.
// Use Cases: Generate dynamic labels in Excel reports by linking shapes to cells with text‑manipulating formulas. | Build dashboards where changing a formula automatically updates chart or textbox annotations. | Create localized worksheets by substituting language strings in linked shapes. | Automate report generation where shape captions reflect calculated values without manual editing.
// AI Prompts: Write C# code using Aspose.Cells to add a TextBox shape, link it to a cell, assign a SUBSTITUTE formula, recalculate, and update the shape’s displayed text. | Explain the interaction between SetLinkedCell and UpdateSelectedValue for reflecting formula results in a shape. | Provide a verification method to confirm that a shape’s text matches the result of a SUBSTITUTE formula in the linked cell. | Show how to modify SUBSTITUTE parameters at runtime and automatically refresh linked shapes.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to create a workbook, add a TextBox shape, link it to cell B2, set a SUBSTITUTE formula that changes "Hello World" to "Hello Aspose", recalculate formulas, refresh the shape’s displayed text, verify the result, and save the file as LinkedShapeSubstitute.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a text box shape to the worksheet
        // Parameters: upperLeftRow, upperLeftColumn, lowerRightRow, lowerRightColumn, upperLeftPixel, upperLeftPixel2
        Shape shape = sheet.Shapes.AddTextBox(1, 1, 3, 3, 150, 40);
        shape.Text = "Initial Text";

        // Link the shape to cell B2 (row index 1, column index 1)
        shape.SetLinkedCell("B2", true, true);

        // Set a formula in the linked cell that uses SUBSTITUTE to replace "World" with "Aspose"
        Cell linkedCell = sheet.Cells["B2"];
        linkedCell.Formula = "SUBSTITUTE(\"Hello World\",\"World\",\"Aspose\")";

        // Recalculate formulas so the linked cell gets the new value
        workbook.CalculateFormula();

        // Update the shape's displayed text from the linked cell value
        shape.UpdateSelectedValue();

        // Verify the updated text (should be "Hello Aspose")
        Console.WriteLine("Shape text after update: " + shape.Text);

        // Save the workbook
        workbook.Save("LinkedShapeSubstitute.xlsx");
    }
}
