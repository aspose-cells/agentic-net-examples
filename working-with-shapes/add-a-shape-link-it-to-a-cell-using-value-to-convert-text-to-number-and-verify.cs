// Title: Aspose.Cells .NET: Add Rectangle Shape, Link to Cell with Text‑to‑Number Conversion, and Verify
// Description: This example creates a workbook, inserts the string "123" into cell B2 with conversion enabled so it becomes numeric, adds a rectangle shape, links the shape to B2 using A1 notation and locale‑aware settings, updates the shape's displayed value, prints the linked cell address and numeric value, and saves the file.
// Keywords: Aspose.Cells C# shape linking | SetLinkedCell A1 notation | PutValue conversion flag | rectangle shape Excel .NET | verify linked cell value | locale aware linked shape | Aspose.Cells shape example
// Common Searches: how to link a shape to a cell in Aspose.Cells | Aspose.Cells convert text to number when putting value | SetLinkedCell parameters isR1C1 isLocal explanation | retrieve linked cell address from a shape Aspose.Cells | update shape value after linking to a cell
// Developer Intent: Create a shape, link it to a cell whose text is stored as a number, and confirm the link programmatically.
// Use Cases: Generate reports where a shape displays a calculated total from a linked cell. | Build interactive dashboards that automatically reflect numeric changes in linked shapes. | Automate Excel workbook creation with validation of text‑to‑number conversion before saving.
// AI Prompts: Write C# code using Aspose.Cells to add a rectangle shape, link it to cell B2 with text‑to‑number conversion, and output the linked cell address. | Explain the effect of the isR1C1 and isLocal parameters in SetLinkedCell for locale‑aware linking. | Show how to read back and verify the numeric value of a linked cell after saving the workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// This example creates a workbook, inserts the string "123" into cell B2 with conversion enabled so it becomes numeric, adds a rectangle shape, links the shape to B2 using A1 notation and locale‑aware settings, updates the shape's displayed value, prints the linked cell address and numeric value, and saves the file.
class ShapeLinkedCellDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Put a text value into cell B2 and let Aspose.Cells convert it to a number
        // The second parameter (true) enables conversion; the value becomes numeric 123
        cells["B2"].PutValue("123", true);

        // Add a rectangle shape to the worksheet (row, column, width, height, top, left)
        // Here we place it at row 3, column 1 with size 100x50 points
        Shape rectShape = worksheet.Shapes.AddRectangle(3, 1, 100, 50, 0, 0);

        // Link the shape's value to cell B2
        // isR1C1 = false (A1 style), isLocal = true (locale‑aware)
        rectShape.SetLinkedCell("$B$2", false, true);

        // Optionally update the shape's selected value from the linked cell
        rectShape.UpdateSelectedValue();

        // Verify: output the linked cell address and the numeric value stored in the cell
        Console.WriteLine("Shape's LinkedCell: " + rectShape.LinkedCell);
        Console.WriteLine("Linked cell value (numeric): " + cells["B2"].Value);

        // Save the workbook (required to persist the shape and link)
        workbook.Save("ShapeLinkedCellDemo.xlsx");
    }
}
