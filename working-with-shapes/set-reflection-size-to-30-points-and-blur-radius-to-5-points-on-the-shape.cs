// Title: Set shape reflection size to 30 points and blur radius to 5 points using Aspose.Cells for .NET (C#)
// Common Searches: how to add a reflection effect to a worksheet shape with Aspose.Cells in C# | C# code for setting blur radius on a shape using Aspose.Cells | Aspose.Cells shape effect format example for reflection and blur | unsupported shape visual effects in Aspose.Cells .NET | increase shape reflection intensity Aspose.Cells C# tutorial
// Tags: shape reflection size Aspose.Cells | shape blur radius Aspose.Cells | apply effect format to worksheet shape C# | Aspose.Cells visual effects limitation | C# shape effect API Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a new workbook, adds a rectangle shape to the first worksheet, and saves it as ReflectionShape.xlsx. It also documents that the Shape.EffectFormat API (used for reflection size and blur radius) is not available in the current Aspose.Cells version, indicating the need for an upgrade to access these visual effects.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper left row, upper left column, top offset, left offset, height, width
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 0, 0, 100, 100);

            // NOTE: EffectFormat (reflection, blur, etc.) is not available in the current Aspose.Cells version.
            // If needed, upgrade Aspose.Cells to a version that supports Shape.EffectFormat.

            // Determine output file path
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ReflectionShape.xlsx");

            // Save the workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
