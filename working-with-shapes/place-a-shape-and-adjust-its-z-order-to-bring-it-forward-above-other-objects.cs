// Title: Adjust Z‑Order of Overlapping Shapes in Aspose.Cells for .NET – Bring to Front or Send to Back
// Description: C# example that creates a workbook, adds two overlapping rectangle shapes, and uses the ToFrontOrBack method to change their layering (bring one shape forward, send the other backward) before saving the file.
// Keywords: Aspose.Cells | C# shape Z-order | BringToFront | SendToBack | ToFrontOrBack | overlapping shapes | Excel shape layering | Aspose.Cells Drawing API | rectangle shape | worksheet shapes
// Common Searches: Aspose.Cells bring shape to front C# | change Z-order of Excel shapes with Aspose.Cells | send shape to back Aspose.Cells .NET | ToFrontOrBack method example | layer overlapping shapes in Aspose.Cells workbook
// Developer Intent: Add overlapping shapes to a worksheet and control which shape appears on top by modifying their Z‑order.
// Use Cases: Design a diagram where foreground elements must overlay background graphics. | Place a company logo above data cells in an automated report. | Create interactive dashboards with layered shapes that need dynamic reordering.
// AI Prompts: Show C# code to move a shape two positions forward in the Z‑order using Aspose.Cells. | Generate an example that toggles shape layering based on a runtime flag. | Explain how ToFrontOrBack differs from the dedicated BringToFront and SendToBack methods.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeZOrderDemo
{
    // C# example that creates a workbook, adds two overlapping rectangle shapes, and uses the ToFrontOrBack method to change their layering (bring one shape forward, send the other backward) before saving the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add two overlapping rectangle shapes
                // shape1 added first (back initially)
                Shape shape1 = worksheet.Shapes.AddRectangle(10, 10, 100, 100, 0, 0);
                shape1.Text = "Back Shape";

                // shape2 added on top of shape1 (higher Z-order)
                Shape shape2 = worksheet.Shapes.AddRectangle(50, 50, 100, 100, 0, 0);
                shape2.Text = "Front Shape";

                // Bring shape2 one position forward (valid range)
                shape2.ToFrontOrBack(1); // equivalent to BringToFront

                // Send shape1 to back (optional)
                shape1.ToFrontOrBack(-1); // equivalent to SendToBack

                // Save the workbook
                string outputPath = "ShapeZOrderDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
