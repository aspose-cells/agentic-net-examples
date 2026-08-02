// Title: Set PerspectiveDiagonalBottomLeft Shadow Preset on a Rectangle Shape with Aspose.Cells for .NET
// Description: This example creates a new workbook, adds two rectangle shapes, applies identical shadow settings (transparency, blur, distance) and assigns different presets—OffsetDiagonalBottomRight to the first shape and PerspectiveDiagonalBottomLeft to the second. The workbook is saved so you can see the visual impact of the PerspectiveDiagonalBottomLeft preset.
// Keywords: Aspose.Cells | C# | .NET | shape shadow preset | PerspectiveDiagonalBottomLeft | OffsetDiagonalBottomRight | Excel rectangle shadow | visual comparison | shadow effect | Excel shape formatting
// Common Searches: Aspose.Cells set shape shadow preset C# | PerspectiveDiagonalBottomLeft example Aspose.Cells | compare shadow presets Aspose.Cells | how to change shape shadow in Excel using Aspose.Cells | C# code for rectangle shadow effect Aspose.Cells
// Developer Intent: Show how to apply the PerspectiveDiagonalBottomLeft shadow preset to a rectangle shape and compare its appearance with another preset.
// Use Cases: Create a design reference sheet that illustrates different shadow styles for Excel shapes. | Generate a template where specific shadow presets are applied automatically for brand‑consistent reports. | Test visual impact of shadow presets before finalizing workbook styling. | Document shape‑formatting guidelines for team members using Aspose.Cells.
// AI Prompts: Generate C# code with Aspose.Cells that sets a rectangle's ShadowEffect.PresetType to PerspectiveDiagonalBottomLeft and saves the workbook. | Explain the visual differences between OffsetDiagonalBottomRight and PerspectiveDiagonalBottomLeft shadow presets in an Excel file created with Aspose.Cells. | Provide a C# snippet that loops through multiple shapes, applies various shadow presets—including PerspectiveDiagonalBottomLeft—and outputs a workbook for side‑by‑side comparison.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShadowDemo
{
    // This example creates a new workbook, adds two rectangle shapes, applies identical shadow settings (transparency, blur, distance) and assigns different presets—OffsetDiagonalBottomRight to the first shape and PerspectiveDiagonalBottomLeft to the second. The workbook is saved so you can see the visual impact of the PerspectiveDiagonalBottomLeft preset.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add first rectangle shape with an initial shadow preset
            Shape shape1 = sheet.Shapes.AddRectangle(1, 0, 1, 0, 150, 100);
            shape1.ShadowEffect.PresetType = PresetShadowType.OffsetDiagonalBottomRight; // initial preset
            shape1.ShadowEffect.Transparency = 0.3;
            shape1.ShadowEffect.Blur = 15;
            shape1.ShadowEffect.Distance = 10;

            // Add second rectangle shape to demonstrate the PerspectiveDiagonalLowerLeft preset
            Shape shape2 = sheet.Shapes.AddRectangle(4, 0, 4, 0, 150, 100);
            shape2.ShadowEffect.PresetType = PresetShadowType.PerspectiveDiagonalLowerLeft; // target preset
            shape2.ShadowEffect.Transparency = 0.3;
            shape2.ShadowEffect.Blur = 15;
            shape2.ShadowEffect.Distance = 10;

            // Save the workbook to view the visual difference between the two shadow presets
            workbook.Save("ShadowPresetComparison.xlsx");
        }
    }
}
