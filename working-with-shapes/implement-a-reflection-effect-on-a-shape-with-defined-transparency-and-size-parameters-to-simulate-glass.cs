// Title: Create a semi‑transparent rectangle shape in an Excel worksheet using Aspose.Cells for .NET to mimic a glass reflection
// AI Prompts: Write C# code that adds a rectangle shape to the first worksheet of a new workbook and applies a LightBlue fill with 40% transparency using Aspose.Cells. | Demonstrate how to set the FillFormat color and Transparency properties of a shape in Aspose.Cells for .NET to achieve a glass‑like appearance.
// Common Searches: Aspose.Cells C# set shape fill color and transparency in Excel | how to simulate glass effect on an Excel shape with Aspose.Cells .NET | add rectangle shape with 40% transparency to worksheet using Aspose.Cells | Aspose.Cells shape fill format transparency example
// Tags: Aspose.Cells shape fill transparency C# | add rectangle shape Aspose.Cells .NET | glass effect shape fill Aspose.Cells | semi transparent shape fill Excel Aspose.Cells | Aspose.Cells workbook shape styling

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// // Creates a new workbook, inserts a rectangle shape on the first worksheet, sets its fill color to LightBlue with 40% transparency to emulate a glass‑like look, and saves the file as GlassReflection.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper left row, upper left column, top offset, left offset, height, width
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 0, 0, 150, 80);

            // The current Aspose.Cells version does not expose EffectFormat for shapes.
            // As an alternative, set a semi‑transparent fill to simulate a glass‑like appearance.
            shape.FillFormat.ForeColor = System.Drawing.Color.LightBlue;
            shape.FillFormat.Transparency = 0.4; // 40% transparent

            // Save the workbook to a file
            string outputPath = "GlassReflection.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
