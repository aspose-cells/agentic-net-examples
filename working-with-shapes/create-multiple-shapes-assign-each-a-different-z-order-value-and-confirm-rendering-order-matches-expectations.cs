// Title: Aspose.Cells .NET: Create Overlapping Shapes, Set Z‑Order, and Verify Stacking
// Description: Demonstrates how to add three overlapping rectangle shapes to a worksheet, assign explicit ZOrderPosition values, move shapes to the front or back with ToFrontOrBack, output the order before and after changes, and save the workbook to confirm the visual stacking in Excel.
// Keywords: Aspose.Cells shape Z-order | C# Aspose.Cells ToFrontOrBack | change shape stacking order | overlapping shapes Excel | .NET workbook shape rendering | Aspose.Cells ZOrderPosition example
// Common Searches: Aspose.Cells set shape Z order .NET | bring shape to front Aspose.Cells | move shape to back using ToFrontOrBack | verify shape stacking order in Excel | C# Aspose.Cells overlapping rectangles
// Developer Intent: Create overlapping shapes, control their Z‑order programmatically, and confirm that the rendered order matches the intended hierarchy.
// Use Cases: Build layered diagrams in automated Excel reports where specific graphics must appear on top. | Dynamically reorder shapes to highlight key data points before publishing a workbook. | Adjust visual hierarchy after applying conditional formatting or data‑driven updates.
// AI Prompts: Generate C# code with Aspose.Cells that adds five shapes, assigns custom ZOrderPosition values, then moves the third shape to the front using ToFrontOrBack. | Explain the algorithm Aspose.Cells uses to calculate step counts for ToFrontOrBack based on a shape's current ZOrderPosition. | Write a unit test in .NET that asserts ZOrderPosition values before and after calling ToFrontOrBack to ensure correct stacking.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsZOrderDemo
{
    // Demonstrates how to add three overlapping rectangle shapes to a worksheet, assign explicit ZOrderPosition values, move shapes to the front or back with ToFrontOrBack, output the order before and after changes, and save the workbook to confirm the visual stacking in Excel.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add three overlapping rectangle shapes
                // Parameters: upperLeftRow, top, upperLeftColumn, left, height, width
                Shape shape1 = worksheet.Shapes.AddRectangle(5, 0, 5, 0, 120, 120);
                Shape shape2 = worksheet.Shapes.AddRectangle(8, 0, 8, 0, 120, 120);
                Shape shape3 = worksheet.Shapes.AddRectangle(11, 0, 11, 0, 120, 120);

                // Assign distinct Z-order positions (lower value = back)
                shape1.ZOrderPosition = 0; // backmost
                shape2.ZOrderPosition = 1;
                shape3.ZOrderPosition = 2; // frontmost

                // Output initial Z-order positions
                Console.WriteLine("Initial Z-order positions:");
                Console.WriteLine($"Shape1 (ID {shape1.Id}) ZOrderPosition = {shape1.ZOrderPosition}");
                Console.WriteLine($"Shape2 (ID {shape2.Id}) ZOrderPosition = {shape2.ZOrderPosition}");
                Console.WriteLine($"Shape3 (ID {shape3.Id}) ZOrderPosition = {shape3.ZOrderPosition}");

                // Bring shape1 to the front safely
                int forwardSteps = worksheet.Shapes.Count - 1 - shape1.ZOrderPosition;
                shape1.ToFrontOrBack(forwardSteps); // move to the front

                // Send shape3 to the back safely
                int backwardSteps = -shape3.ZOrderPosition;
                shape3.ToFrontOrBack(backwardSteps); // move to the back

                // Output Z-order positions after adjustments
                Console.WriteLine("\nZ-order positions after ToFrontOrBack adjustments:");
                Console.WriteLine($"Shape1 (ID {shape1.Id}) ZOrderPosition = {shape1.ZOrderPosition}");
                Console.WriteLine($"Shape2 (ID {shape2.Id}) ZOrderPosition = {shape2.ZOrderPosition}");
                Console.WriteLine($"Shape3 (ID {shape3.Id}) ZOrderPosition = {shape3.ZOrderPosition}");

                // Save the workbook to verify visual rendering order in Excel
                workbook.Save("ShapeZOrderDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
