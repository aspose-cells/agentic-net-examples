// Title: Link a Shape to a CONCAT Formula Cell and Verify Displayed Text – Aspose.Cells for .NET (C#)
// Description: Creates a workbook, fills A1 and B1, sets C1 to =CONCAT(A1,B1), recalculates, adds a rectangle shape, links it to C1, forces an update, prints the concatenated result, and saves the file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# shape linked cell | CONCAT formula | rectangle shape | LinkedCell property | UpdateSelectedValue | dynamic shape text | Excel automation | cell formula display
// Common Searches: Aspose.Cells link shape to cell C# | display CONCAT result in a shape | refresh shape text after formula calculation | verify shape shows cell value Aspose | how to bind rectangle to cell in Aspose.Cells
// Developer Intent: Attach a rectangle shape to a cell that contains a CONCAT formula and confirm the shape displays the concatenated string.
// Use Cases: Dashboard labels that automatically combine values from multiple cells. | Printable reports where shapes show merged identifiers without manual editing. | Interactive worksheets that reflect real‑time product codes or names in shape captions.
// AI Prompts: Generate C# code to link a rectangle shape to a cell with a CONCAT formula and ensure the shape updates after workbook calculation using Aspose.Cells. | Explain how to programmatically verify that a shape's displayed text matches the result of its linked CONCAT formula. | Show how to automatically refresh a shape's text when source cells used in a formula are changed in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a workbook, fills A1 and B1, sets C1 to =CONCAT(A1,B1), recalculates, adds a rectangle shape, links it to C1, forces an update, prints the concatenated result, and saves the file using Aspose.Cells for .NET.
class ShapeLinkedCellConcatDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate cells A1 and B1 with sample text
        sheet.Cells["A1"].PutValue("Hello");
        sheet.Cells["B1"].PutValue("World");

        // Set a formula in C1 that concatenates A1 and B1 using CONCAT function
        sheet.Cells["C1"].Formula = "=CONCAT(A1,B1)";

        // Recalculate the workbook to evaluate the formula
        workbook.CalculateFormula();

        // Add a rectangle shape to the worksheet
        // Parameters: upper left row, upper left column, top, left, width, height
        Shape shape = sheet.Shapes.AddRectangle(2, 2, 0, 0, 100, 30);

        // Link the shape to cell C1 so it displays the concatenated result
        shape.LinkedCell = "C1";

        // Ensure the shape updates its displayed value from the linked cell
        shape.UpdateSelectedValue();

        // Retrieve the value from the linked cell for verification
        string linkedValue = sheet.Cells["C1"].StringValue;

        // Output the result
        Console.WriteLine("Cell C1 (concatenated): " + linkedValue);
        Console.WriteLine("Shape's linked cell address: " + shape.GetLinkedCell(false, false));
        Console.WriteLine("Verification: Shape should display \"" + linkedValue + "\".");

        // Save the workbook (optional, demonstrates lifecycle usage)
        workbook.Save("ShapeLinkedCellConcatDemo.xlsx");
    }
}
