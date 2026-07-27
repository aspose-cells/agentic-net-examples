// Title: C# Example: Move a Shape to the Front Layer with Aspose.Cells ToFrontOrBack
// Description: Demonstrates how to create a workbook, add overlapping rectangle shapes, and change their z‑order using Shape.ToFrontOrBack. A positive integer moves a shape forward (front layer), while a negative integer sends it backward, then the workbook is saved.
// Keywords: Aspose.Cells C# shape layering | ToFrontOrBack method | bring shape to front Aspose.Cells | move shape forward Excel | z‑order shapes Aspose.Cells | shape front back example
// Common Searches: Aspose.Cells move shape to front | C# ToFrontOrBack example | change shape z‑order Aspose.Cells | bring overlapping shape forward in Excel using Aspose | how to send shape to back Aspose.Cells
// Developer Intent: Adjust the z‑order of a specific Shape object so it appears above or below other worksheet objects.
// Use Cases: Display a newly added annotation on top of existing graphics by calling shape.ToFrontOrBack(1). | Hide a background watermark by sending it to the back with shape.ToFrontOrBack(-1). | Create interactive reports where users can reorder chart elements dynamically via positive or negative offsets.
// AI Prompts: Generate C# code that uses Aspose.Cells to bring a given shape to the front layer with ToFrontOrBack(1). | Write a reusable function that accepts a Shape and an integer offset to modify its z‑order using ToFrontOrBack. | Explain the effect of positive and negative values passed to Shape.ToFrontOrBack in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add overlapping rectangle shapes, and change their z‑order using Shape.ToFrontOrBack. A positive integer moves a shape forward (front layer), while a negative integer sends it backward, then the workbook is saved.
    public class MoveShapeToFrontDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add two overlapping shapes to demonstrate z‑order
                Shape shape1 = worksheet.Shapes.AddRectangle(10, 10, 100, 100, 0, 0);
                Shape shape2 = worksheet.Shapes.AddRectangle(30, 30, 100, 100, 0, 0);

                // Bring shape2 to the front by moving it forward 1 position
                shape2.ToFrontOrBack(1);   // positive integer => front

                // Optionally, send shape1 to the back
                shape1.ToFrontOrBack(-1);  // negative integer => back

                // Save the workbook
                workbook.Save("ShapeFrontDemo.xlsx");
                Console.WriteLine("Workbook saved as ShapeFrontDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            MoveShapeToFrontDemo.Run();
        }
    }
}
