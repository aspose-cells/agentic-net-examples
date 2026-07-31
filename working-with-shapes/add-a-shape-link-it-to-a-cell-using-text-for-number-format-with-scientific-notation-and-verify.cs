// Title: Aspose.Cells for .NET – Add a Rectangle Shape Linked to a Scientific‑Notation Cell
// Description: Demonstrates how to create a workbook, write a large number to A1, apply a custom scientific notation format (0.00E+00), insert a rectangle shape, link the shape to the formatted cell with SetLinkedCell, retrieve the address via GetLinkedCell for verification, and save the file as ShapeLinkedCellScientific.xlsx.
// Keywords: Aspose.Cells | C# | .NET | rectangle shape | linked cell | SetLinkedCell | GetLinkedCell | scientific notation format | custom number format | workbook automation
// Common Searches: Aspose.Cells link shape to cell C# | set scientific notation for a cell Aspose.Cells | retrieve linked cell address from shape Aspose.Cells | add rectangle shape programmatically Aspose.Cells | verify shape‑cell link Aspose.Cells .NET
// Developer Intent: Create a rectangle shape, bind it to a cell formatted in scientific notation, and confirm the binding programmatically.
// Use Cases: Design dashboards where shapes act as visual markers for cells displaying values in scientific notation. | Automate report generation that requires shapes to stay synchronized with formatted data cells. | Validate that shape‑to‑cell links persist after cell formatting or value changes.
// AI Prompts: Write C# code using Aspose.Cells to add a rectangle shape, link it to cell A1 formatted with scientific notation, and output the linked cell address. | Show how to read a shape's linked cell, extract the cell's scientific formatted value, and update the shape's text accordingly in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to create a workbook, write a large number to A1, apply a custom scientific notation format (0.00E+00), insert a rectangle shape, link the shape to the formatted cell with SetLinkedCell, retrieve the address via GetLinkedCell for verification, and save the file as ShapeLinkedCellScientific.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Put a numeric value into cell A1
        sheet.Cells["A1"].PutValue(123456789);

        // Apply scientific notation format to the cell
        Style sciStyle = sheet.Cells["A1"].GetStyle();
        sciStyle.Custom = "0.00E+00"; // scientific format
        sheet.Cells["A1"].SetStyle(sciStyle);

        // Add a rectangle shape (acts as a label)
        // Parameters: upper left row, upper left column, row offset, column offset, width, height
        Shape shape = sheet.Shapes.AddRectangle(2, 2, 0, 0, 120, 30);

        // Link the shape to cell A1
        shape.SetLinkedCell("$A$1", false, true);

        // Verify the linked cell
        string linkedCell = shape.GetLinkedCell(true, true);
        Console.WriteLine("Shape is linked to: " + linkedCell);

        // Save the workbook
        workbook.Save("ShapeLinkedCellScientific.xlsx");
    }
}
