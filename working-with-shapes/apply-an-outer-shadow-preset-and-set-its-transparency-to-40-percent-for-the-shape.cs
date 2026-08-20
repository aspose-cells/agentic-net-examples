// Title: Apply a 40% Transparent Outer Shadow Preset to an Excel Shape with Aspose.Cells for .NET (C#)
// Description: Shows how to create a workbook, insert a rectangle shape, retrieve its ShadowEffect, set the PresetShadowType to an outer shadow (OffsetBottom), adjust transparency to 0.4 (40%), and save the result as an .xlsx file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# shape shadow | outer shadow preset | PresetShadowType | OffsetBottom | shadow transparency | Excel shape formatting | Aspose.Cells .NET | apply shadow effect | Excel workbook styling
// Common Searches: Aspose.Cells add outer shadow to shape | C# set shape shadow transparency Aspose.Cells | How to use PresetShadowType in Aspose.Cells | Apply shadow effect to Excel rectangle using .NET | Set shadow transparency percentage Aspose.Cells
// Developer Intent: Add an outer shadow preset to a worksheet shape and configure its transparency to 40% using Aspose.Cells for .NET.
// Use Cases: Enhance Excel reports with subtle depth by applying outer shadows to diagrammatic shapes. | Generate template workbooks that automatically style all shapes with a consistent shadow before distribution. | Prepare workbooks for PDF conversion where shadowed shapes improve visual presentation.
// AI Prompts: Write C# code with Aspose.Cells that applies the OffsetBottom outer shadow preset and 40% transparency to every shape in a workbook. | Show how to change the shadow preset to OffsetTopRight and set transparency based on a user‑provided percentage. | Create a reusable method that accepts a PresetShadowType and a transparency value, then applies those settings to a given shape.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShadowDemo
{
    // Shows how to create a workbook, insert a rectangle shape, retrieve its ShadowEffect, set the PresetShadowType to an outer shadow (OffsetBottom), adjust transparency to 0.4 (40%), and save the result as an .xlsx file using Aspose.Cells for .NET.
    public class ApplyOuterShadow
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle shape to the worksheet
                // Parameters: upper left row, upper left column, upper left offset X, upper left offset Y, width, height
                Shape shape = worksheet.Shapes.AddRectangle(2, 2, 0, 0, 150, 100);

                // Get the shadow effect of the shape
                ShadowEffect shadow = shape.ShadowEffect;

                // Apply an outer shadow preset (e.g., OffsetBottom)
                shadow.PresetType = PresetShadowType.OffsetBottom;

                // Set the transparency of the shadow to 40% (0.4)
                shadow.Transparency = 0.4;

                // Save the workbook to a file
                string outputPath = "ShapeWithOuterShadow.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ApplyOuterShadow.Run();
        }
    }
}
