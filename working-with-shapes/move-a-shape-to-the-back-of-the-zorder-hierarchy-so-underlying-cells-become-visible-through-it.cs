// Title: Move a worksheet shape to the back (Z‑order) with Aspose.Cells for .NET (C#)
// Description: This C# example creates a new Workbook, adds a rectangle shape to the first worksheet, uses shape.ToFrontOrBack(1) to send the shape to the back of the Z‑order so underlying cells become visible, and saves the file as ShapeBackDemo.xlsx.
// Keywords: Aspose.Cells | C# | .NET | shape Z-order | ToFrontOrBack | send shape to back | move shape behind cells | worksheet shape ordering | Aspose.Cells example | shape layering
// Common Searches: Aspose.Cells move shape to back | How to send a shape behind cells in Aspose.Cells C# | Z‑order shape Aspose.Cells .NET | shape.ToFrontOrBack usage | place shape behind worksheet data Aspose
// Developer Intent: Place a worksheet shape behind the cell grid so the cells are visible through or above the shape.
// Use Cases: Add a semi‑transparent watermark rectangle behind data rows. | Insert a background image that stays behind all cell content. | Create layered graphics where a shape must appear behind charts, tables, or other shapes.
// AI Prompts: Generate C# code that moves any worksheet shape to the back using Aspose.Cells ToFrontOrBack method. | Explain the possible values for ToFrontOrBack and show how to bring a shape to the front after it has been sent to the back. | Provide a C# example that adds a picture shape, sends it to the back, and then adjusts its opacity for a watermark effect.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// This C# example creates a new Workbook, adds a rectangle shape to the first worksheet, uses shape.ToFrontOrBack(1) to send the shape to the back of the Z‑order so underlying cells become visible, and saves the file as ShapeBackDemo.xlsx.
class MoveShapeToBack
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            Shape shape = worksheet.Shapes.AddRectangle(5, 5, 200, 100, 0, 0);

            // Send the shape to the back of the Z‑order hierarchy (1 = back)
            shape.ToFrontOrBack(1);

            // Save the workbook
            workbook.Save("ShapeBackDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
