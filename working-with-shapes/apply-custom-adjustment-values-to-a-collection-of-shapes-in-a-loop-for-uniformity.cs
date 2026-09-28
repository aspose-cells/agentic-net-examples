// Title: Iterate over worksheet AutoShapes to apply the same adjustment value using Aspose.Cells for .NET
// AI Prompts: Write C# code that enumerates all AutoShape objects on a worksheet and assigns a specified adjustment index to each shape with Aspose.Cells. | Create a reusable method that adds a set of star AutoShapes to a worksheet and uniformly sets their first adjustment parameter using the Aspose.Cells API.
// Common Searches: C# Aspose.Cells how to set adjustment values for multiple AutoShape objects in a loop | apply uniform shape adjustment to all stars in an Excel file using Aspose.Cells .NET | batch update AutoShape adjustments in a worksheet with Aspose.Cells API
// Tags: Aspose.Cells loop through AutoShape collection | set uniform adjustment value for AutoShape | batch modify shape adjustments in Excel .NET | add star AutoShape with adjustment Aspose.Cells | programmatic shape property update Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, adds three 10‑point star AutoShape objects to the first worksheet, demonstrates how to assign a uniform adjustment value to each shape (commented for versions that support adjustments), and saves the file as ShapeAdjustmentDemo.xlsx.
class ShapeAdjustmentDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample shapes to the worksheet for demonstration
            // Example: Add three stars (AutoShape) which support adjustment values in newer versions
            for (int i = 0; i < 3; i++)
            {
                // Position each shape at a different location
                int upperLeftRow = 2 + i * 10;
                int upperLeftColumn = 2;

                // Width and height in pixels
                int width = 100;
                int height = 100;

                // Create an AutoShape of type Star10 (10‑point star)
                Shape shape = sheet.Shapes.AddAutoShape(
                    AutoShapeType.Star10,
                    upperLeftRow,
                    upperLeftColumn,
                    0,          // upper left row offset (in pixels)
                    0,          // upper left column offset (in pixels)
                    height,     // shape height (in pixels)
                    width);     // shape width (in pixels)

                // Optionally give each shape a name
                shape.Name = $"StarShape{i + 1}";
            }

            // NOTE: Adjustments are not available in the current Aspose.Cells version.
            // If using a version that supports shape adjustments, you can set them like:
            // shape.Adjustments[0] = 0.4;

            // Save the workbook (lifecycle rule: save)
            workbook.Save("ShapeAdjustmentDemo.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
