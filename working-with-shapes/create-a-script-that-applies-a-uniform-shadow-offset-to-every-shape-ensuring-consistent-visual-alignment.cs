// Title: Apply Uniform Shadow Offset to All Shapes in Aspose.Cells (C#)
// Description: Shows how to create a workbook, add rectangle, oval, and textbox shapes, define a single shadow preset (distance, blur, transparency, size) and apply it to every shape in a worksheet using Aspose.Cells for .NET, then save the file as UniformShadow.xlsx.
// Keywords: Aspose.Cells | C# | shape shadow | uniform shadow offset | ShadowEffect | preset shadow type | Excel shape styling | apply shadow to all shapes
// Common Searches: C# set same shadow for all shapes Aspose.Cells | How to apply uniform shadow effect to worksheet shapes | Aspose.Cells shadow distance blur transparency example | Loop through Worksheet.Shapes to set ShadowEffect | Apply preset shadow to multiple shapes in Excel using .NET
// Developer Intent: Set identical shadow properties for every shape in a worksheet to ensure consistent visual alignment across the document.
// Use Cases: Standardize visual style of icons and diagrams in automated reports. | Create a branding template where new shapes inherit a predefined shadow. | Batch‑style Excel dashboards by applying a uniform shadow to all existing shapes.
// AI Prompts: Generate C# code with Aspose.Cells that assigns a uniform ShadowEffect (distance, blur, transparency, size) to all worksheet shapes. | Show a loop over Worksheet.Shapes to set PresetShadowType.OffsetBottom for each shape. | Explain how to modify shadow properties of existing shapes in an Aspose.Cells workbook programmatically.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Shows how to create a workbook, add rectangle, oval, and textbox shapes, define a single shadow preset (distance, blur, transparency, size) and apply it to every shape in a worksheet using Aspose.Cells for .NET, then save the file as UniformShadow.xlsx.
class ApplyUniformShadow
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample shapes to demonstrate the shadow effect
        sheet.Shapes.AddRectangle(1, 0, 1, 0, 100, 150);
        sheet.Shapes.AddOval(3, 0, 3, 0, 120, 80);
        sheet.Shapes.AddTextBox(5, 0, 5, 0, 100, 200);

        // Uniform shadow settings
        double uniformDistance = 20;                     // offset distance in points
        PresetShadowType uniformPreset = PresetShadowType.OffsetBottom;
        double uniformBlur = 10;                         // blur amount
        double uniformTransparency = 0.3;                // 30% transparent
        double uniformSize = 1.0;                        // size factor

        // Apply the same shadow effect to every shape in the worksheet
        foreach (Shape shape in sheet.Shapes)
        {
            ShadowEffect shadow = shape.ShadowEffect;
            shadow.PresetType = uniformPreset;
            shadow.Distance = uniformDistance;
            shadow.Blur = uniformBlur;
            shadow.Transparency = uniformTransparency;
            shadow.Size = uniformSize;
        }

        // Save the workbook with the applied shadow effects
        workbook.Save("UniformShadow.xlsx");
    }
}
