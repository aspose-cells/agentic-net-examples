// Title: Aspose.Cells for .NET – Add a Rectangle Shape Linked to a Cell with a Custom TEXT Number Format
// Description: Demonstrates how to create a workbook, write a numeric value to A1, apply a custom number format that prefixes the value (e.g., "Order #: 12345"), insert a rectangle shape, link the shape to the cell, refresh the shape to show the formatted text, and save the file.
// Keywords: Aspose.Cells C# shape linked cell | custom number format TEXT Aspose.Cells | update shape text from linked cell | AddRectangle Aspose.Cells example | LinkedCell property C#
// Common Searches: link shape to cell Aspose.Cells .NET | display custom formatted value in shape Aspose.Cells | Aspose.Cells rectangle shape linked cell example | how to use TEXT number format with linked shape
// Developer Intent: Create a rectangle shape, bind it to a worksheet cell that uses a custom TEXT number format, and verify that the shape reflects the formatted cell value.
// Use Cases: Dynamic order numbers on a dashboard shape that update automatically when the source cell changes. | Customer ID badges in a report where the shape shows a prefixed ID (e.g., "Customer #: 00123"). | Invoice templates that display a formatted invoice number inside a shape, keeping the visual layout in sync with cell data.
// AI Prompts: Generate C# code using Aspose.Cells to add a rectangle shape linked to cell A1 with a custom TEXT number format and ensure the shape displays the formatted value. | Explain the role of Shape.UpdateSelectedValue when a shape is linked to a cell that has a custom number format in Aspose.Cells. | Provide step‑by‑step instructions to verify the LinkedCell property of a shape and save the workbook after linking.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to create a workbook, write a numeric value to A1, apply a custom number format that prefixes the value (e.g., "Order #: 12345"), insert a rectangle shape, link the shape to the cell, refresh the shape to show the formatted text, and save the file.
class ShapeLinkedCellExample
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule: create)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Set a numeric value in cell A1
        Cell targetCell = sheet.Cells["A1"];
        targetCell.PutValue(12345);

        // Apply a custom number format that uses TEXT (e.g., display as "Order #: 12345")
        Style style = targetCell.GetStyle();
        style.Custom = "\"Order #: \"0";
        targetCell.SetStyle(style);

        // Add a rectangle shape to the worksheet
        // Parameters: upper left row, upper left column, upper left offset (pixels), upper left offset (pixels), width, height
        Shape shape = sheet.Shapes.AddRectangle(2, 1, 0, 0, 150, 50);

        // Link the shape to cell A1
        shape.LinkedCell = "$A$1";

        // Ensure the shape displays the linked cell's formatted value
        shape.UpdateSelectedValue();

        // Optional: set some placeholder text (will be replaced by linked cell value after UpdateSelectedValue)
        shape.Text = "Placeholder";

        // Verify the linked cell address
        Console.WriteLine("Shape is linked to cell: " + shape.LinkedCell);

        // Save the workbook (lifecycle rule: save)
        workbook.Save("ShapeLinkedCellExample.xlsx");
    }
}
