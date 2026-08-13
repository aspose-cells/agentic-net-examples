// Title: Aspose.Cells for .NET: Move a Shape to Front/Back, Capture Z‑Order Index, and Compare in C#
// Description: This C# example creates a workbook, adds two overlapping rectangle shapes, records their initial ZOrderPosition values, brings the second shape to the front with ToFrontOrBack(1), captures the new index, sends it to the back with ToFrontOrBack(-1), captures the final index, prints comparison results, and saves the file as ShapeZOrderDemo.xlsx.
// Keywords: Aspose.Cells | C# | Excel shape Z-order | ToFrontOrBack | ZOrderPosition | bring shape to front | send shape to back | shape layering | Aspose.Cells .NET | Excel automation | programmatic shape ordering
// Common Searches: Aspose.Cells change shape Z-order C# | How to bring a shape to front in Aspose.Cells .NET | Get ZOrderPosition after moving shape Aspose.Cells | Send shape to back Excel using Aspose.Cells | Compare shape Z-order before and after ToFrontOrBack | C# code to reorder overlapping shapes in Excel
// Developer Intent: The developer needs to programmatically reorder overlapping shapes in an Excel worksheet, retrieve their Z-order indices, and verify that moving a shape to the front or back updates the indices as expected.
// Use Cases: Generate Excel reports with dynamic diagrams where certain shapes must appear on top. | Implement interactive Excel dashboards that reorder shapes based on user input. | Automated testing of shape layering to ensure visual consistency after workbook generation. | Create layered graphics such as flowcharts where Z-order is controlled via code.
// AI Prompts: Write C# code using Aspose.Cells to move a shape to the front, read its ZOrderPosition, then move it to the back and output both indices. | Explain the effect of ToFrontOrBack(1) and ToFrontOrBack(-1) on ZOrderPosition for overlapping shapes in Aspose.Cells. | Generate a unit test in C# that asserts ZOrderPosition increases after ToFrontOrBack(1) and decreases after ToFrontOrBack(-1). | Provide a step‑by‑step guide for capturing and comparing shape Z-order before and after reordering with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // This C# example creates a workbook, adds two overlapping rectangle shapes, records their initial ZOrderPosition values, brings the second shape to the front with ToFrontOrBack(1), captures the new index, sends it to the back with ToFrontOrBack(-1), captures the final index, prints comparison results, and saves the file as ShapeZOrderDemo.xlsx.
    public class ShapeZOrderDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add two overlapping rectangle shapes
            Shape shape1 = sheet.Shapes.AddRectangle(5, 5, 100, 100, 0, 0);
            Shape shape2 = sheet.Shapes.AddRectangle(10, 10, 100, 100, 0, 0);

            // Capture initial Z-order positions
            int initialPos1 = shape1.ZOrderPosition;
            int initialPos2 = shape2.ZOrderPosition;

            // Bring shape2 to the front
            shape2.ToFrontOrBack(1);
            int afterFrontPos2 = shape2.ZOrderPosition;

            // Send shape2 to the back
            shape2.ToFrontOrBack(-1);
            int afterBackPos2 = shape2.ZOrderPosition;

            // Output comparison results
            Console.WriteLine($"Shape1 initial Z-order: {initialPos1}");
            Console.WriteLine($"Shape2 initial Z-order: {initialPos2}");
            Console.WriteLine($"Shape2 after ToFrontOrBack(1): {afterFrontPos2}");
            Console.WriteLine($"Shape2 after ToFrontOrBack(-1): {afterBackPos2}");
            Console.WriteLine($"Front move successful: {afterFrontPos2 > initialPos2}");
            Console.WriteLine($"Back move successful: {afterBackPos2 < afterFrontPos2}");

            // Save the workbook
            string outputPath = "ShapeZOrderDemo.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
    }
}
