// Title: Apply a Uniform Shadow Offset to Every Shape in an Aspose.Cells Workbook (C#)
// Description: Creates a workbook, adds sample shapes, defines a preset shadow type and distance, then loops through all worksheets and shapes to apply the same ShadowEffect settings before saving the file.
// Keywords: Aspose.Cells shape shadow | C# uniform shadow offset | Excel shape ShadowEffect | preset shadow type Aspose.Cells | apply shadow to all shapes | iterate shapes workbook | Aspose.Cells visual styling
// Common Searches: set same shadow offset for all shapes Aspose.Cells C# | apply uniform shadow distance to Excel shapes using Aspose.Cells | C# code to change shadow preset for every shape in a workbook | Aspose.Cells bulk shape formatting shadow effect | how to iterate shapes and set shadow in Aspose.Cells
// Developer Intent: Apply an identical shadow preset and distance to every shape across all worksheets in a workbook.
// Use Cases: Standardize shape appearance in automatically generated reports. | Maintain consistent design in dashboards where all shapes share the same shadow style. | Enforce branding guidelines by applying a uniform shadow to every shape in a template workbook.
// AI Prompts: Generate C# code with Aspose.Cells that sets a custom shadow color and blur while keeping a uniform preset for all shapes. | Show how to vary shadow distance by shape type but retain the same preset across the workbook. | Provide an example that applies the uniform shadow settings and saves the workbook to a memory stream instead of a file.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShadowDemo
{
    // Creates a workbook, adds sample shapes, defines a preset shadow type and distance, then loops through all worksheets and shapes to apply the same ShadowEffect settings before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample shapes to demonstrate the effect
            sheet.Shapes.AddRectangle(2, 0, 2, 0, 150, 100);
            sheet.Shapes.AddOval(5, 0, 5, 0, 120, 80);
            sheet.Shapes.AddTextBox(8, 0, 8, 0, 100, 200);

            // Define the uniform shadow settings
            PresetShadowType uniformPreset = PresetShadowType.OffsetDiagonalBottomRight; // chosen offset type
            double uniformDistance = 15; // distance in points for all shapes

            // Apply the uniform shadow to every shape in every worksheet
            foreach (Worksheet ws in workbook.Worksheets)
            {
                foreach (Shape shape in ws.Shapes)
                {
                    ShadowEffect shadow = shape.ShadowEffect;
                    shadow.PresetType = uniformPreset;
                    shadow.Distance = uniformDistance;
                }
            }

            // Save the workbook with the applied shadow effects
            workbook.Save("UniformShadowDemo.xlsx");
        }
    }
}
