// Title: Send a Shape Named "Background" to the Back (Lowest Z‑Order) with Aspose.Cells for .NET
// Description: Load a workbook, locate the shape called "Background" on the first worksheet, set its ZOrderPosition to 0 (or use ToFrontOrBack(-1)) to place it behind all other objects, and save the file.
// Keywords: Aspose.Cells shape Z-order | move shape to back Aspose.Cells | background shape lowest Z order .NET | Aspose.Cells send shape to back | C# Excel shape layering | ZOrderPosition Aspose.Cells | ToFrontOrBack Aspose.Cells | Excel shape ordering .NET | Aspose.Cells background image behind charts
// Common Searches: Aspose.Cells send shape to back | set ZOrderPosition of a shape in C# | move background shape behind other objects Excel | how to change shape layering Aspose.Cells | C# Aspose.Cells shape Z-order example
// Developer Intent: Place the "Background" shape at the lowest Z‑order so every other worksheet object appears above it.
// Use Cases: Ensure a watermark or background image never obscures data tables, charts, or text boxes in generated reports. | Programmatically reorder decorative shapes in a template to maintain a clean visual hierarchy. | Adjust shape layering when adding multiple objects dynamically to keep important content on top.
// AI Prompts: Show C# code that sends a specific shape to the back of the Z‑order using Aspose.Cells. | Provide an example of setting a shape's ZOrderPosition to the lowest level and verifying the change. | Explain alternative methods such as ToFrontOrBack for moving shapes forward or backward in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Load a workbook, locate the shape called "Background" on the first worksheet, set its ZOrderPosition to 0 (or use ToFrontOrBack(-1)) to place it behind all other objects, and save the file.
class SetBackgroundZOrder
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("Input.xlsx");
        Worksheet sheet = workbook.Worksheets[0];

        // Find the shape named "Background"
        Shape backgroundShape = null;
        foreach (Shape shape in sheet.Shapes)
        {
            if (shape.Name == "Background")
            {
                backgroundShape = shape;
                break;
            }
        }

        if (backgroundShape != null)
        {
            // Set the shape to the lowest Z‑order position (backmost)
            backgroundShape.ZOrderPosition = 0;
            // Alternatively, you could use: backgroundShape.ToFrontOrBack(-1);
        }
        else
        {
            Console.WriteLine("Shape named 'Background' not found.");
        }

        // Save the workbook with the updated Z‑order
        workbook.Save("Output.xlsx");
    }
}
