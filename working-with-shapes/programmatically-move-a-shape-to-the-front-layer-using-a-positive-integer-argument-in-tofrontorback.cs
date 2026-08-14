// Title: Aspose.Cells .NET: Move a worksheet shape to the front layer using ToFrontOrBack(int)
// Description: C# example that creates a workbook, adds two overlapping rectangle shapes, and uses shape2.ToFrontOrBack(1) to shift the second shape one level forward, placing it above the first shape before saving as ShapeFrontDemo.xlsx.
// Keywords: Aspose.Cells ToFrontOrBack | C# move shape to front | shape z-order Aspose.Cells | bring shape forward .NET | Excel shape layering | Aspose.Cells shape ordering | worksheet shape front layer | Aspose.Cells C# example
// Common Searches: Aspose.Cells move shape to front | ToFrontOrBack method C# | change shape z-order Aspose.Cells | bring shape forward Excel API | shape layering Aspose.Cells example
// Developer Intent: Programmatically set a specific worksheet shape as the topmost object.
// Use Cases: Ensure a title textbox appears above data bars in a financial report. | Overlay a company logo on top of decorative shapes in a dashboard worksheet. | Adjust the visual stacking order of chart annotations and shapes for clearer presentation.
// AI Prompts: Generate C# code that sends a shape to the back layer in Aspose.Cells by passing a negative integer to ToFrontOrBack. | Show how to reorder multiple worksheet shapes in a loop using ToFrontOrBack to achieve a custom z‑order. | Explain how different integer values affect shape layering when using ToFrontOrBack in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // C# example that creates a workbook, adds two overlapping rectangle shapes, and uses shape2.ToFrontOrBack(1) to shift the second shape one level forward, placing it above the first shape before saving as ShapeFrontDemo.xlsx.
    public class MoveShapeToFrontDemo
    {
        // Entry point for the application
        public static void Main(string[] args)
        {
            try
            {
                Run();
                Console.WriteLine("Workbook saved successfully as ShapeFrontDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add two overlapping rectangle shapes
            Shape shape1 = worksheet.Shapes.AddRectangle(10, 10, 100, 100, 0, 0);
            Shape shape2 = worksheet.Shapes.AddRectangle(30, 30, 100, 100, 0, 0);

            // Bring shape2 to the front layer (positive integer moves forward)
            shape2.ToFrontOrBack(1);

            // Save the workbook with the updated shape order
            workbook.Save("ShapeFrontDemo.xlsx");
        }
    }
}
