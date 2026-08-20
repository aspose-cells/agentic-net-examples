// Title: Get and log a shape's shadow preset using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, add a rectangle shape, assign a shadow preset, save and reload the file, then read the shape's ShadowEffect.PresetType and output it to the console for debugging purposes.
// Keywords: Aspose.Cells shape shadow preset | C# read ShadowEffect.PresetType | Aspose.Cells debug shadow effect | retrieve shape shadow type .NET | Aspose.Cells shadow enum
// Common Searches: how to read shape shadow preset Aspose.Cells C# | Aspose.Cells get ShadowEffect.PresetType after saving | log shape shadow preset in .NET | Aspose.Cells retrieve shadow effect enum value | debug shape shadow settings in Excel workbook
// Developer Intent: Extract the current shadow preset of a worksheet shape and display it for troubleshooting.
// Use Cases: Verify that a shape's shadow setting survives workbook serialization. | Log shadow presets to diagnose visual inconsistencies in generated reports. | Automated test that compares each shape's actual shadow preset with expected values.
// AI Prompts: Write C# code that reads a shape's ShadowEffect.PresetType with Aspose.Cells and writes the result to a log file. | Show how to iterate over all shapes in a worksheet and print each shape's name and shadow preset. | Explain how to compare a shape's shadow preset against a target enum value and trigger an alert when they differ.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShadowDemo
{
    // Demonstrates how to create a workbook, add a rectangle shape, assign a shadow preset, save and reload the file, then read the shape's ShadowEffect.PresetType and output it to the console for debugging purposes.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape
            Shape shape = sheet.Shapes.AddRectangle(1, 0, 1, 0, 100, 100);

            // Set a known shadow preset for demonstration
            shape.ShadowEffect.PresetType = PresetShadowType.OffsetBottom;

            // Save the workbook
            workbook.Save("ShadowDemo.xlsx");

            // Load the workbook back
            Workbook loadedWorkbook = new Workbook("ShadowDemo.xlsx");
            Shape loadedShape = loadedWorkbook.Worksheets[0].Shapes[0];

            // Retrieve the current shadow preset type
            PresetShadowType currentPreset = loadedShape.ShadowEffect.PresetType;

            // Log the preset type for debugging
            Console.WriteLine("Current Shadow Preset: " + currentPreset);
        }
    }
}
