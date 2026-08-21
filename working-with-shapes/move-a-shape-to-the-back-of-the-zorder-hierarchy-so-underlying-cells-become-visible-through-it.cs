// Title: C# – Move a Shape to the Back (Z‑Order) in Aspose.Cells so Cells Appear Through It
// Description: Demonstrates how to create a workbook, add a rectangle shape, set its ZOrderPosition to 0 to send it to the back of the Z‑order hierarchy, and save the file, making the cells underneath visible.
// Keywords: Aspose.Cells C# shape ZOrderPosition | send shape to back Aspose.Cells | shape layering worksheet | reveal cells behind shape | Aspose.Cells shape order example
// Common Searches: Aspose.Cells move shape to back C# | how to set ZOrderPosition in Aspose.Cells | make cells visible through shape Aspose.Cells | change shape Z‑order worksheet Aspose.Cells
// Developer Intent: The developer needs to place a worksheet shape behind other objects so that the underlying cells become visible.
// Use Cases: Add a watermark shape behind data without obscuring it. | Insert a background image shape while keeping cell values readable. | Reorder overlapping shapes to control visual hierarchy on a sheet.
// AI Prompts: Write C# code using Aspose.Cells to add a rectangle and send it to the back of the Z‑order. | Explain the purpose of the ZOrderPosition property and how to read or modify it for worksheet shapes. | Provide a sample that adjusts Z‑order positions for multiple shapes in a single worksheet.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to create a workbook, add a rectangle shape, set its ZOrderPosition to 0 to send it to the back of the Z‑order hierarchy, and save the file, making the cells underneath visible.
class MoveShapeToBackExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape that will cover some cells
            // Parameters: upper left row, upper left column, upper left row offset (in pixels),
            // upper left column offset (in pixels), height (in pixels), width (in pixels)
            Shape shape = worksheet.Shapes.AddRectangle(5, 5, 0, 0, 100, 200);

            // Send the shape to the back of the Z‑order so cells underneath become visible
            // In Aspose.Cells the Z‑order is controlled by the ZOrderPosition property.
            // Setting it to 0 places the shape at the back.
            shape.ZOrderPosition = 0;

            // Save the workbook
            workbook.Save("ShapeMovedToBack.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
