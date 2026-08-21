// Title: Apply the PerspectiveDiagonalBottomLeft Shadow Preset to a Shape using Aspose.Cells for .NET
// Description: Demonstrates how to create an Excel workbook, add a rectangle shape, apply the OffsetBottom shadow preset, then switch to the PerspectiveDiagonalBottomLeft preset (via PresetShadowType.PerspectiveDiagonalLowerLeft), save both versions, and read back the preset to confirm the change.
// Keywords: Aspose.Cells | C# | shape shadow preset | PerspectiveDiagonalBottomLeft | PresetShadowType | Excel shadow effect | compare shadow presets | OffsetBottom | Aspose.Cells API | Excel workbook styling
// Common Searches: Aspose.Cells set shape shadow to PerspectiveDiagonalBottomLeft | C# change shape shadow preset in Excel | How to use PresetShadowType in Aspose.Cells | Compare OffsetBottom and PerspectiveDiagonalBottomLeft shadows | Save Excel file after modifying shape shadow
// Developer Intent: Show how to assign the PerspectiveDiagonalBottomLeft shadow preset to a shape, persist the change, and verify the applied preset.
// Use Cases: Generate side‑by‑side screenshots for documentation that illustrate different shadow styles. | Create Excel reports where key shapes are highlighted with a PerspectiveDiagonalBottomLeft shadow. | Build an interactive UI that lets end‑users pick a shadow preset and instantly preview the result in an Excel file.
// AI Prompts: Write C# code with Aspose.Cells to set a rectangle's shadow preset to PerspectiveDiagonalBottomLeft and save the workbook. | Provide an example that switches a shape's shadow from OffsetBottom to PerspectiveDiagonalBottomLeft, then reads the preset type to confirm the update. | Explain how to compare multiple shadow presets on a shape in Aspose.Cells and output the resulting preset names.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShadowDemo
{
    // Demonstrates how to create an Excel workbook, add a rectangle shape, apply the OffsetBottom shadow preset, then switch to the PerspectiveDiagonalBottomLeft preset (via PresetShadowType.PerspectiveDiagonalLowerLeft), save both versions, and read back the preset to confirm the change.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape to demonstrate shadow effects
            Shape shape = sheet.Shapes.AddRectangle(2, 2, 2, 2, 150, 100);

            // Set an initial preset shadow type (e.g., OffsetBottom) and save the result
            shape.ShadowEffect.PresetType = PresetShadowType.OffsetBottom;
            workbook.Save("Shadow_Initial.xlsx");

            // Change the preset shadow type to PerspectiveDiagonalLowerLeft
            // (Corresponds to the requested "PerspectiveDiagonalBottomLeft")
            shape.ShadowEffect.PresetType = PresetShadowType.PerspectiveDiagonalLowerLeft;

            // Save the workbook again to observe the visual difference
            workbook.Save("Shadow_PerspectiveDiagonalLowerLeft.xlsx");

            // Optional: Load the saved file to verify the preset type
            Workbook loaded = new Workbook("Shadow_PerspectiveDiagonalLowerLeft.xlsx");
            Shape loadedShape = loaded.Worksheets[0].Shapes[0];
            Console.WriteLine("Current PresetType: " + loadedShape.ShadowEffect.PresetType);
        }
    }
}
