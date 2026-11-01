// Title: How to bind a rectangle shape’s text to a CONCATENATE formula cell using Aspose.Cells for .NET
// AI Prompts: Create a rectangle shape on a worksheet and set its Text property to reference a cell containing a CONCATENATE formula with Aspose.Cells for .NET. | Link a shape’s displayed text to the result of a CONCATENATE function, apply fill color and line weight, then save the workbook. | Update a linked shape so its text automatically reflects changes in the underlying concatenated cell value.
// Common Searches: Aspose.Cells .NET link shape text to cell formula result | display CONCATENATE function output inside a rectangle shape using Aspose.Cells | bind Excel shape to cell with formula using C# Aspose.Cells example | how to make shape text update when CONCATENATE cell changes in Aspose.Cells
// Tags: shape text bound to formula cell Aspose.Cells | rectangle shape linked to CONCATENATE result .NET | Aspose.Cells set shape text from cell | dynamic shape content based on formula Aspose.Cells | formatting linked shape fill color Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, writes "Hello" and "World" to A1 and B1, sets C1 to a CONCATENATE formula, adds a rectangle shape at row 2 column 2, links the shape’s Text property to C1 so it displays the combined text, applies a light‑yellow fill and a thin border, and saves the file as LinkedShape.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Set values in A1 and B1
        cells["A1"].PutValue("Hello");
        cells["B1"].PutValue("World");

        // Set a formula in C1 that concatenates A1 and B1 with a space
        cells["C1"].Formula = "=CONCATENATE(A1, \" \", B1)";

        // Add a rectangle shape to the worksheet
        // Parameters: type, upper left row, upper left column, top, left, height, width
        Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 0, 0, 100, 30);

        // Link the shape's text to cell C1 so it displays the concatenated result
        shape.Text = "=C1";

        // Optionally format the shape (e.g., fill color, line)
        shape.FillFormat.ForeColor = System.Drawing.Color.LightYellow;
        shape.LineFormat.Weight = 1.0;

        // Save the workbook
        workbook.Save("LinkedShape.xlsx");
    }
}
