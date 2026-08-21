// Title: Set 1.2‑point character spacing for all shape text in Excel with Aspose.Cells for .NET
// Description: Creates a workbook, adds text box shapes, iterates over each shape on the first worksheet, and assigns Shape.TextOptions.Spacing = 1.2 points for shapes that contain text, then saves the file as ShapesWithSpacing.xlsx.
// Keywords: Aspose.Cells | C# | Excel shape text spacing | Shape.TextOptions.Spacing | character spacing | text box formatting | increase readability | programmatic Excel styling | apply spacing to all shapes | Aspose.Cells .NET API
// Common Searches: Aspose.Cells set shape text spacing | How to change character spacing in Excel shapes using C# | Apply uniform text spacing to text boxes with Aspose.Cells | Shape.TextOptions.Spacing example | Adjust spacing of shape text programmatically
// Developer Intent: Apply a consistent 1.2‑point character spacing to every shape that contains text in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Enhance readability of generated reports by standardizing spacing in all text boxes. | Create Excel templates with uniform text appearance across multiple shapes. | Prepare workbooks for printing where precise character spacing improves layout aesthetics.
// AI Prompts: Generate C# code with Aspose.Cells that sets Shape.TextOptions.Spacing to 1.5 points for all shapes containing text. | Show how to iterate through worksheet shapes and apply different character spacing based on shape type (e.g., text box vs. callout). | Explain how to reset Shape.TextOptions.Spacing to the default value for selected shapes using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// Creates a workbook, adds text box shapes, iterates over each shape on the first worksheet, and assigns Shape.TextOptions.Spacing = 1.2 points for shapes that contain text, then saves the file as ShapesWithSpacing.xlsx.
class ApplyCharacterSpacing
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add sample shapes with text to demonstrate the effect
        Shape shape1 = worksheet.Shapes.AddTextBox(0, 0, 0, 0, 200, 100);
        shape1.Text = "First shape text";

        Shape shape2 = worksheet.Shapes.AddTextBox(5, 0, 0, 0, 200, 100);
        shape2.Text = "Second shape text";

        // Iterate through all shapes in the worksheet
        foreach (Shape shape in worksheet.Shapes)
        {
            // Apply spacing only if the shape contains text
            if (!string.IsNullOrEmpty(shape.Text))
            {
                // Set character spacing to 1.2 points
                shape.TextOptions.Spacing = 1.2;
            }
        }

        // Save the workbook with the updated spacing
        workbook.Save("ShapesWithSpacing.xlsx");
    }
}
