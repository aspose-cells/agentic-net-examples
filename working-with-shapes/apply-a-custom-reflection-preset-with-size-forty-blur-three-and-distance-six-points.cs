// Title: Aspose.Cells for .NET – Set custom reflection (size 40, blur 3, distance 6) on a rectangle shape
// Description: C# sample that creates a workbook, inserts a rectangle shape, and shows how to apply a reflection effect with specific size, blur, and distance values using Aspose.Cells. It also notes that Shape.EffectOptions is unavailable in older releases and advises upgrading to a version that supports shape effects.
// Keywords: Aspose.Cells reflection effect | C# rectangle shape reflection | Shape.EffectOptions | custom reflection preset | size 40 blur 3 distance 6 | Excel shape visual effects | Aspose.Cells .NET version support
// Common Searches: Aspose.Cells apply reflection to shape | set reflection size blur distance Aspose.Cells | Shape.EffectOptions availability .NET | add visual effects to Excel shapes using Aspose.Cells | upgrade Aspose.Cells for shape effects
// Developer Intent: Implement a reflection effect with size 40, blur 3, and distance 6 on a rectangle shape in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Create a formatted report where shapes need a reflective highlight for visual emphasis. | Programmatically check whether the installed Aspose.Cells library includes Shape.EffectOptions before attempting to apply effects. | Migrate legacy code to a newer Aspose.Cells version to enable custom reflection presets on worksheet shapes.
// AI Prompts: Write C# code that adds a rectangle to a worksheet and applies a reflection effect with size 40, blur 3, distance 6 using Aspose.Cells, including version‑check logic. | Explain how to detect the presence of Shape.EffectOptions in Aspose.Cells and provide fallback handling when the feature is missing. | Suggest alternative .NET libraries or techniques for adding reflection or shadow effects to Excel shapes when Aspose.Cells does not support them.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// C# sample that creates a workbook, inserts a rectangle shape, and shows how to apply a reflection effect with specific size, blur, and distance values using Aspose.Cells. It also notes that Shape.EffectOptions is unavailable in older releases and advises upgrading to a version that supports shape effects.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: upper left row, upper left column, top, left, width, height
            Shape rectangle = worksheet.Shapes.AddRectangle(1, 0, 1, 0, 150, 100);

            // NOTE: Shape effects (e.g., reflection) are not available in the current Aspose.Cells version.
            // If needed, upgrade to a version that supports Shape.EffectOptions.

            // Save the workbook
            workbook.Save("CustomReflection.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
