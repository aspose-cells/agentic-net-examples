// Title: Apply an 8‑point blue glow to all shapes in an Aspose.Cells worksheet (C#)
// Description: Creates a new workbook, adds sample shapes, iterates through Worksheet.Shapes, sets each shape's Glow.Size to 8 points and Glow.Color to blue, then saves the file as AllShapesBlueGlow.xlsx.
// Keywords: Aspose.Cells shape glow | C# apply blue glow | worksheet shapes glow size | set glow effect Aspose.Cells | blue glow 8 points
// Common Searches: Aspose.Cells set glow effect on shapes C# | C# add blue glow to all worksheet shapes | How to apply uniform glow to shapes in Aspose.Cells | Set glow radius 8 points Aspose.Cells
// Developer Intent: Apply an 8‑point blue glow to every shape on a worksheet.
// Use Cases: Highlight all diagram elements in a generated report with a consistent blue glow. | Create a template where any newly added shape automatically receives the predefined glow styling. | Prepare presentation‑style spreadsheets that emphasize shapes before distribution.
// AI Prompts: Generate C# code that applies a red glow of 5 points to every auto shape in an Aspose.Cells worksheet. | Show how to change the glow color based on shape type while iterating through Worksheet.Shapes. | Provide an example that applies a gradient glow effect to shapes using Aspose.Cells in C#.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a new workbook, adds sample shapes, iterates through Worksheet.Shapes, sets each shape's Glow.Size to 8 points and Glow.Color to blue, then saves the file as AllShapesBlueGlow.xlsx.
class ApplyBlueGlowToAllShapes
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add sample shapes (optional, for demonstration)
            worksheet.Shapes.AddRectangle(1, 0, 1, 0, 100, 80);
            // Use AddOval instead of the non‑existent AddEllipse method
            worksheet.Shapes.AddOval(2, 0, 2, 0, 120, 90);
            worksheet.Shapes.AddAutoShape(AutoShapeType.RoundedRectangle, 3, 0, 3, 0, 150, 100);

            // Apply a blue glow effect with size 8 points to every shape in the worksheet
            foreach (Shape shape in worksheet.Shapes)
            {
                // Access the GlowEffect object of the shape
                GlowEffect glow = shape.Glow;

                // Set the radius (size) of the glow to 8 points
                glow.Size = 8;

                // Create a CellsColor instance for the glow color and set it to blue
                CellsColor glowColor = workbook.CreateCellsColor();
                glowColor.Color = Color.Blue;
                glow.Color = glowColor;
            }

            // Define output file path
            string outputPath = "AllShapesBlueGlow.xlsx";

            // Save the workbook with the applied glow effects
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
