// Title: Set 1.2‑point character spacing for all shape text in an Excel workbook with Aspose.Cells for .NET
// Description: Shows how to create a workbook, add text‑filled shapes, and use TextOptions.Spacing to apply a 1.2‑point character gap to each shape that contains text, then save the file as ShapesWithSpacing.xlsx.
// Keywords: Aspose.Cells | C# | Excel shape text spacing | TextOptions.Spacing | character spacing | shape formatting | Aspose.Cells .NET | Excel automation | adjust text spacing | shape text readability
// Common Searches: Aspose.Cells change character spacing in shape | C# set text spacing for Excel shapes | how to adjust spacing between characters in shape text using Aspose.Cells | apply uniform text spacing to all shapes in workbook | TextOptions.Spacing example C#
// Developer Intent: Apply a uniform 1.2‑point character spacing to every shape that contains text in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Enhance readability of text inside text boxes, rectangles, or other shapes in generated reports. | Maintain consistent typography across multiple shape types when programmatically building Excel files. | Prepare Excel documents for print or PDF export where precise character spacing improves visual layout.
// AI Prompts: Rewrite the sample to accept the spacing value as a method parameter instead of a hard‑coded constant. | Show how to assign different spacing values based on shape type (e.g., larger spacing for text boxes, smaller for rectangles). | Provide code that verifies the applied character spacing by reading TextOptions.Spacing after saving the workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// Shows how to create a workbook, add text‑filled shapes, and use TextOptions.Spacing to apply a 1.2‑point character gap to each shape that contains text, then save the file as ShapesWithSpacing.xlsx.
class ApplyCharacterSpacing
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a few sample shapes with text
        Shape shape1 = worksheet.Shapes.AddTextBox(0, 0, 0, 0, 200, 100);
        shape1.Text = "First shape text";

        Shape shape2 = worksheet.Shapes.AddRectangle(1, 0, 0, 150, 250, 0);
        shape2.Text = "Second shape text";

        // Apply character spacing of 1.2 points to all shapes that contain text
        foreach (Shape shape in worksheet.Shapes)
        {
            // Only process shapes that have a TextOptions object (i.e., contain text)
            if (!string.IsNullOrEmpty(shape.Text))
            {
                TextOptions textOptions = shape.TextOptions;
                textOptions.Spacing = 1.2; // Set spacing to 1.2 points
            }
        }

        // Save the workbook
        workbook.Save("ShapesWithSpacing.xlsx");
    }
}
