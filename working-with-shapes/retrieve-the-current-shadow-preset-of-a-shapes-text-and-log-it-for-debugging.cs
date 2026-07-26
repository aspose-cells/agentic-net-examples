// Title: Retrieve and Log a Shape's Shadow Preset (PresetShadowType) with Aspose.Cells for .NET
// Description: Demonstrates how to add a rectangle shape, assign a shadow preset, save and reload the workbook, then read the Shape.ShadowEffect.PresetType and output it to the console for debugging purposes.
// Keywords: Aspose.Cells C# shape shadow preset | PresetShadowType read | ShadowEffect debugging Aspose.Cells | retrieve shape shadow type .NET | Aspose.Cells get shadow preset
// Common Searches: how to read shape shadow preset in Aspose.Cells | Aspose.Cells get ShadowEffect.PresetType after saving | C# log shape shadow preset Aspose.Cells | retrieve current shadow preset of Excel shape
// Developer Intent: Read the current shadow preset of a shape and display it for troubleshooting.
// Use Cases: Confirm that a shape's shadow setting persists after workbook serialization. | Log shadow presets to identify visual issues in generated Excel reports. | Conditionally adjust a shape's shadow based on its existing preset.
// AI Prompts: Generate C# code using Aspose.Cells that returns the ShadowEffect.PresetType of a given shape as a string. | Show how to check a shape's current shadow preset and replace it with another preset in Aspose.Cells for .NET. | Create a method that iterates over all worksheet shapes, logs each shape's PresetShadowType, and optionally modifies it.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShadowPresetDemo
{
    // Demonstrates how to add a rectangle shape, assign a shadow preset, save and reload the workbook, then read the Shape.ShadowEffect.PresetType and output it to the console for debugging purposes.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 0, 150, 100);

            // Set a known shadow preset for demonstration
            shape.ShadowEffect.PresetType = PresetShadowType.OffsetBottom;

            // Save the workbook
            workbook.Save("ShadowPresetDemo.xlsx");

            // Load the saved workbook
            Workbook loadedWorkbook = new Workbook("ShadowPresetDemo.xlsx");
            Shape loadedShape = loadedWorkbook.Worksheets[0].Shapes[0];

            // Retrieve the current shadow preset type
            PresetShadowType currentPreset = loadedShape.ShadowEffect.PresetType;

            // Log the preset type for debugging
            Console.WriteLine("Current Shadow Preset: " + currentPreset);
        }
    }
}
