// Title: Aspose.Cells .NET: Create Overlapping Shapes, Set Z‑Order, and Verify Rendering Order
// Description: Demonstrates how to add three overlapping rectangle shapes to a worksheet, assign explicit ZOrderPosition values, adjust stacking with ToFrontOrBack, output the before‑and‑after order, and save the workbook to confirm the visual rendering sequence.
// Keywords: Aspose.Cells shape Z-order | C# set shape stacking order | ToFrontOrBack Aspose.Cells | overlapping Excel shapes .NET | ZOrderPosition programmatic | shape rendering order verification
// Common Searches: change shape Z-order Aspose.Cells | bring shape to front C# Aspose | move shape backward Aspose.Cells .NET | verify Excel shape layering | adjust overlapping shapes order programmatically
// Developer Intent: Create overlapping shapes, control their Z‑order programmatically, and confirm that the rendered order matches the expected stacking.
// Use Cases: Define background, middle, and foreground layers for a custom report layout. | Reorder annotation or comment shapes on a chart so important notes appear on top. | Validate that dynamic shape ordering is reflected in the generated Excel file for dashboards.
// AI Prompts: Write C# code that adds five different shapes to a worksheet and sets a custom Z‑order sequence using Aspose.Cells. | Explain how ToFrontOrBack works and how to move a shape multiple steps forward or backward. | Provide a method to programmatically check the visual stacking order of shapes after reordering in an Excel workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsZOrderDemo
{
    // Demonstrates how to add three overlapping rectangle shapes to a worksheet, assign explicit ZOrderPosition values, adjust stacking with ToFrontOrBack, output the before‑and‑after order, and save the workbook to confirm the visual rendering sequence.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add three overlapping rectangle shapes
            // Parameters: upperLeftRow, top, upperLeftColumn, left, height, width
            Shape shape1 = worksheet.Shapes.AddRectangle(5, 0, 5, 0, 120, 120);
            shape1.Name = "Shape1";
            shape1.FillFormat.ForeColor = System.Drawing.Color.FromArgb(255, 200, 200, 255); // Light blue

            Shape shape2 = worksheet.Shapes.AddRectangle(8, 0, 8, 0, 120, 120);
            shape2.Name = "Shape2";
            shape2.FillFormat.ForeColor = System.Drawing.Color.FromArgb(255, 200, 255, 200); // Light green

            Shape shape3 = worksheet.Shapes.AddRectangle(11, 0, 11, 0, 120, 120);
            shape3.Name = "Shape3";
            shape3.FillFormat.ForeColor = System.Drawing.Color.FromArgb(255, 255, 200, 200); // Light red

            // Initial Z-order: shape1 at back (0), shape2 in middle (1), shape3 at front (2)
            shape1.ZOrderPosition = 0;
            shape2.ZOrderPosition = 1;
            shape3.ZOrderPosition = 2;

            Console.WriteLine("Initial Z-order positions:");
            Console.WriteLine($"{shape1.Name}: {shape1.ZOrderPosition}");
            Console.WriteLine($"{shape2.Name}: {shape2.ZOrderPosition}");
            Console.WriteLine($"{shape3.Name}: {shape3.ZOrderPosition}");

            // Bring shape1 to the front using ToFrontOrBack (positive value)
            shape1.ToFrontOrBack(1); // Moves shape1 one step forward

            // Send shape3 to the back using ToFrontOrBack (negative value)
            shape3.ToFrontOrBack(-2); // Moves shape3 two steps backward

            // After adjustments, output the new Z-order positions
            Console.WriteLine("\nAdjusted Z-order positions:");
            Console.WriteLine($"{shape1.Name}: {shape1.ZOrderPosition}");
            Console.WriteLine($"{shape2.Name}: {shape2.ZOrderPosition}");
            Console.WriteLine($"{shape3.Name}: {shape3.ZOrderPosition}");

            // Save the workbook to visualize the rendering order
            workbook.Save("ShapeZOrderDemo.xlsx");
        }
    }
}
