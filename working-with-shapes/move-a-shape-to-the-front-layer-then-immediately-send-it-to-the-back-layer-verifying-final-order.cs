// Title: C# – Change Shape Z‑Order with Aspose.Cells: Bring to Front, Send to Back, Verify Order
// Description: Demonstrates how to add overlapping rectangle shapes to a workbook, use the ToFrontOrBack method to move a shape to the front layer, then send it to the back layer, and confirm the final Z‑order by comparing ZOrderPosition values. The workbook is saved after verification.
// Keywords: Aspose.Cells C# shape Z-order | ToFrontOrBack method | bring shape to front Aspose.Cells | send shape to back Aspose.Cells | verify shape stacking order | Excel shape layering .NET | ZOrderPosition property
// Common Searches: Aspose.Cells change shape Z order C# | how to bring a shape to front in Aspose.Cells | send shape to back Aspose.Cells .NET | check shape stacking order Aspose.Cells | move Excel shapes between layers programmatically
// Developer Intent: Programmatically move a shape to the front layer, then immediately send it to the back layer, and confirm that the final stacking order matches the expected hierarchy.
// Use Cases: Adjust overlapping annotations in generated reports so that specific notes appear above or below others at runtime. | Create interactive diagrams where shapes are reordered based on user actions or data-driven conditions. | Validate visual layout before exporting an Excel file to ensure that shape layering conforms to design specifications.
// AI Prompts: Write C# code using Aspose.Cells to move a given shape to the front, then to the back, and return a boolean indicating whether it ends up behind another specified shape. | Explain the meaning of positive and negative arguments in the ToFrontOrBack method and how they affect the ZOrderPosition of shapes. | Refactor the sample into a reusable utility method that accepts a shape, a front offset, and a back offset, adjusts the Z‑order accordingly, and logs the before/after positions.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Demonstrates how to add overlapping rectangle shapes to a workbook, use the ToFrontOrBack method to move a shape to the front layer, then send it to the back layer, and confirm the final Z‑order by comparing ZOrderPosition values. The workbook is saved after verification.
    public class ShapeFrontBackDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add two overlapping rectangle shapes
                Shape shape1 = sheet.Shapes.AddRectangle(10, 10, 100, 100, 0, 0);
                Shape shape2 = sheet.Shapes.AddRectangle(50, 50, 100, 100, 0, 0);

                // Initial Z-order positions
                Console.WriteLine($"Initial ZOrder - shape1: {shape1.ZOrderPosition}, shape2: {shape2.ZOrderPosition}");

                // Bring shape2 to the front (positive order)
                shape2.ToFrontOrBack(1);
                Console.WriteLine($"After ToFrontOrBack(1) - shape1: {shape1.ZOrderPosition}, shape2: {shape2.ZOrderPosition}");

                // Send shape2 to the back (negative order)
                shape2.ToFrontOrBack(-2);
                Console.WriteLine($"After ToFrontOrBack(-2) - shape1: {shape1.ZOrderPosition}, shape2: {shape2.ZOrderPosition}");

                // Verify final order: shape2 should be behind shape1
                bool isBehind = shape2.ZOrderPosition < shape1.ZOrderPosition;
                Console.WriteLine($"Shape2 is behind Shape1: {isBehind}");

                // Save the workbook
                workbook.Save("ShapeFrontBackDemo.xlsx");
                Console.WriteLine("Workbook saved as ShapeFrontBackDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ShapeFrontBackDemo.Run();
        }
    }
}
