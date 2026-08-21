// Title: Move a Watermark Shape to the Back Layer with Aspose.Cells for .NET (C#)
// Description: Loads a workbook, finds the shape named "Watermark" on the first worksheet, sends it to the back using the ToFrontOrBack(-1) method so cells stay visible, and saves the file.
// Keywords: Aspose.Cells move shape to back | C# Aspose.Cells watermark layer | ToFrontOrBack method | Excel shape ordering Aspose | send shape behind cells
// Common Searches: Aspose.Cells send shape to back | C# move watermark behind cells Excel | ToFrontOrBack example Aspose.Cells | layering shapes in Excel with Aspose
// Developer Intent: Place the "Watermark" shape behind all other objects on the worksheet so the underlying cells remain unobscured.
// Use Cases: Adjust watermark visibility after insertion | Control the stacking order of multiple shapes in a report | Ensure decorative graphics do not cover data cells
// AI Prompts: Generate C# code that uses Aspose.Cells to locate a shape named "Watermark" and move it to the back layer of the first worksheet. | Show how to check for a shape's existence, apply ToFrontOrBack(-1), and save the workbook with Aspose.Cells. | Explain the effect of the ToFrontOrBack method on shape layering in an Excel file.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads a workbook, finds the shape named "Watermark" on the first worksheet, sends it to the back using the ToFrontOrBack(-1) method so cells stay visible, and saves the file.
class MoveWatermarkToBack
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("Input.xlsx");

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

        // Save the workbook (replace with your desired output path)
        workbook.Save("Output.xlsx");
    }
}
