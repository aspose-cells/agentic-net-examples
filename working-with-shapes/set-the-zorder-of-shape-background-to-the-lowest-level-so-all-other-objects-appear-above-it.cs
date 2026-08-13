// Title: Aspose.Cells .NET – Send Background Shape to Back (Lowest Z‑Order)
// Description: Demonstrates how to add two rectangles to a worksheet, then move the larger rectangle to the back using the ToFrontOrBack(-1) method so it becomes the lowest‑level shape before saving the workbook.
// Keywords: Aspose.Cells shape Z‑order | ToFrontOrBack method | send shape to back .NET | background shape layering | shape ordering Aspose.Cells
// Common Searches: Aspose.Cells move shape behind others | set shape lowest Z‑order .NET | background rectangle behind data Aspose.Cells | how to send shape to back in workbook
// Developer Intent: Place a shape at the bottom of the Z‑order stack so every other worksheet object appears above it.
// Use Cases: Create a full‑page background rectangle for a report template. | Add a watermark that must stay behind charts, tables, and images. | Layer multiple charts and graphics while keeping a background shape at the bottom.
// AI Prompts: Show code that sends a shape to the back using Aspose.Cells for .NET. | Explain the ToFrontOrBack method and how to set a shape's Z‑order to the lowest level. | Provide an example of layering shapes with a background shape at the lowest Z‑order.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to add two rectangles to a worksheet, then move the larger rectangle to the back using the ToFrontOrBack(-1) method so it becomes the lowest‑level shape before saving the workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a shape that will appear in front (foreground)
            Shape foreground = sheet.Shapes.AddRectangle(2, 2, 50, 50, 200, 150);
            foreground.Name = "Foreground";

            // Add a shape that will serve as the background
            // Adding it after the foreground allows us to move it behind later
            Shape background = sheet.Shapes.AddRectangle(0, 0, 0, 0, 800, 600);
            background.Name = "Background";

            // Send the background shape to the back (lowest Z‑order)
            // Since it is currently at index 1, moving it back by -1 places it at index 0
            background.ToFrontOrBack(-1);

            // Save the workbook
            workbook.Save("ZOrderDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
