// Title: Batch apply identical shadow effect to all shapes on ten worksheets with Aspose.Cells for .NET
// Description: Creates a workbook containing ten worksheets, defines a single set of shadow parameters (angle, blur, distance, size, transparency, preset type) and loops through every worksheet and each shape to assign the same ShadowEffect, then saves the result as ShadowBatchResult.xlsx.
// Keywords: Aspose.Cells | C# | .NET | shape shadow effect | ShadowEffect API | PresetShadowType | batch update shapes | multiple worksheets | Excel automation | GitHub example
// Common Searches: apply same shadow to all shapes Aspose.Cells | C# batch update shape formatting across worksheets | set shadow properties for every shape in an Excel workbook | Aspose.Cells iterate worksheets and shapes | how to use ShadowEffect with PresetShadowType in C#
// Developer Intent: Apply a uniform shadow configuration to every shape on each of ten worksheets in a single operation.
// Use Cases: Ensure consistent visual styling of charts, pictures, and text boxes in a multi‑sheet report. | Create a template workbook where all existing and future shapes share a predefined shadow. | Retrofit an older workbook to corporate branding by updating all shape shadows in one pass.
// AI Prompts: Generate a reusable C# method that accepts shadow parameters and applies them to all shapes in every worksheet of an Aspose.Cells workbook. | Write code that skips worksheets without shapes while still processing those that contain shapes for shadow updates. | Show how to load an existing workbook, change the shadow preset for every shape, and save the modified file.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShadowBatch
{
    // Creates a workbook containing ten worksheets, defines a single set of shadow parameters (angle, blur, distance, size, transparency, preset type) and loops through every worksheet and each shape to assign the same ShadowEffect, then saves the result as ShadowBatchResult.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Ensure there are exactly ten worksheets
            // The workbook starts with one default sheet; add nine more.
            for (int i = 1; i < 10; i++)
            {
                workbook.Worksheets.Add($"Sheet{i + 1}");
            }

            // Define the shadow settings that will be applied to every shape
            // These settings are identical for all shapes across all worksheets.
            const double shadowAngle = 135.0;          // lighting angle in degrees
            const double shadowBlur = 20.0;            // blur amount (points)
            const double shadowDistance = 50.0;        // distance from the shape (points)
            const double shadowSize = 1.0;             // size multiplier
            const double shadowTransparency = 0.3;     // 30% transparent
            const PresetShadowType preset = PresetShadowType.OffsetDiagonalBottomRight; // preset type

            // Iterate through each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through each shape in the current worksheet
                for (int s = 0; s < sheet.Shapes.Count; s++)
                {
                    Shape shape = sheet.Shapes[s];

                    // Access the ShadowEffect object of the shape
                    ShadowEffect shadow = shape.ShadowEffect;

                    // Apply the predefined shadow settings
                    shadow.Angle = shadowAngle;
                    shadow.Blur = shadowBlur;
                    shadow.Distance = shadowDistance;
                    shadow.Size = shadowSize;
                    shadow.Transparency = shadowTransparency;
                    shadow.PresetType = preset;
                }
            }

            // Save the workbook with the updated shadow settings
            workbook.Save("ShadowBatchResult.xlsx");
        }
    }
}
