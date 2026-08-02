// Title: Aspose.Cells for .NET – Add a label shape, link it to a percentage‑formatted cell, and confirm the format
// Description: This example creates a workbook, writes 0.25 to cell A1, applies a custom percentage number format, inserts a label shape at row 2 column 2, links the shape to the formatted cell with SetLinkedCell, checks the cell's IsPercent flag, and saves the file as ShapeLinkedWithPercent.xlsx.
// Keywords: Aspose.Cells | C# | .NET | Excel shape linking | label shape | SetLinkedCell | percentage number format | custom number format | IsPercent property | verify cell format | Excel automation example | Add shape to worksheet
// Common Searches: Aspose.Cells link shape to cell with percentage format | How to use SetLinkedCell in C# | Check if a cell is formatted as percent in Aspose.Cells | Add label shape to Excel workbook using Aspose.Cells | Apply custom number format to a cell in Aspose.Cells
// Developer Intent: Create a label shape, bind it to a percent‑formatted cell, and validate the formatting programmatically.
// Use Cases: Bind a visual label to a data cell so clicking the shape jumps to the cell in the UI. | Apply a percentage format to numeric data and programmatically confirm the format before further processing. | Generate Excel reports where shapes act as interactive markers linked to formatted values.
// AI Prompts: Write C# code with Aspose.Cells that adds a label shape, links it to a cell formatted as a percentage, and prints whether the cell is a percent. | Explain the role of SetLinkedCell and the IsPercent property when linking shapes to formatted cells in Aspose.Cells. | Provide a step‑by‑step tutorial for adding a shape, applying a custom percentage format via TEXT or style, linking the shape, and verifying the format in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// This example creates a workbook, writes 0.25 to cell A1, applies a custom percentage number format, inserts a label shape at row 2 column 2, links the shape to the formatted cell with SetLinkedCell, checks the cell's IsPercent flag, and saves the file as ShapeLinkedWithPercent.xlsx.
class ShapeLinkWithPercentFormat
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Put a numeric value into cell A1
        Cell cell = worksheet.Cells["A1"];
        cell.PutValue(0.25); // 25%

        // Apply a percentage number format to the cell
        Style style = cell.GetStyle();
        style.Number = 10; // Built‑in percentage format
        cell.SetStyle(style);

        // Add a label shape to the worksheet
        // Parameters: upperLeftRow, upperLeftColumn, top, left, height, width
        Label label = (Label)worksheet.Shapes.AddLabel(2, 2, 0, 0, 100, 30);
        label.Text = "Linked Shape";

        // Link the shape to cell A1
        // formula, isR1C1, isLocal
        label.SetLinkedCell("$A$1", false, true);

        // Verify that the cell's style is a percent format
        Style verifyStyle = cell.GetStyle();
        Console.WriteLine("IsPercent after formatting: " + verifyStyle.IsPercent);

        // Save the workbook
        workbook.Save("ShapeLinkedWithPercent.xlsx");
    }
}
