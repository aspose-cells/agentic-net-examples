// Title: C# – Move a Watermark Shape Behind Cells Using Aspose.Cells for .NET
// Description: Loads an Excel file, finds the shape named "Watermark" on the first worksheet, sends it to the back of the Z‑order with ToFrontOrBack(-1) so cell data stays visible, and saves the result.
// Keywords: Aspose.Cells | C# shape Z-order | move shape to back | watermark behind cells | ToFrontOrBack | Excel shape layering | Aspose.Cells API | worksheet shape ordering
// Common Searches: Aspose.Cells send shape to back | C# move watermark behind worksheet cells | How to change shape Z-order in Excel with Aspose | Place watermark behind data using Aspose.Cells | ToFrontOrBack method example
// Developer Intent: Place the "Watermark" shape behind all other objects so worksheet cells remain unobstructed.
// Use Cases: Generating reports where a logo or text watermark must not hide cell values | Preparing templates that require background images while keeping data readable | Automating batch processing of workbooks to standardize shape layering | Adjusting visual hierarchy of multiple shapes before publishing an Excel file
// AI Prompts: Create C# code with Aspose.Cells that locates a shape called "Watermark" and moves it to the back layer of the first worksheet. | Show how to verify a shape exists, apply ToFrontOrBack(-1) to push it behind other objects, and then save the workbook. | Explain the effect of negative values in ToFrontOrBack on shape Z-order and best practices for watermark placement.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads an Excel file, finds the shape named "Watermark" on the first worksheet, sends it to the back of the Z‑order with ToFrontOrBack(-1) so cell data stays visible, and saves the result.
class MoveWatermarkToBack
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Assume the watermark shape is on the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Find the shape named "Watermark"
        Shape watermarkShape = null;
        foreach (Shape shape in sheet.Shapes)
        {
            if (shape.Name == "Watermark")
            {
                watermarkShape = shape;
                break;
            }
        }

        if (watermarkShape != null)
        {
            // Send the shape to the back (negative order moves it behind other objects)
            watermarkShape.ToFrontOrBack(-1);
        }
        else
        {
            Console.WriteLine("Shape named 'Watermark' was not found.");
        }

        // Save the modified workbook (replace with desired output path)
        workbook.Save("output.xlsx");
    }
}
