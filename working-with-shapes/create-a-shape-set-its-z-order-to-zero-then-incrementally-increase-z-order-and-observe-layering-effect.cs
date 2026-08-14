// Title: Aspose.Cells C# – Set Shape Z‑Order and Reorder Overlapping Rectangles
// Description: Demonstrates how to create three overlapping rectangle shapes in a workbook, assign explicit ZOrderPosition values, and use ToFrontOrBack to move shapes forward or backward, revealing the layering effect when the file is opened in Excel.
// Keywords: Aspose.Cells shape Z-order | C# Excel shape layering | ToFrontOrBack Aspose.Cells | ZOrderPosition rectangle | reorder overlapping shapes .NET | Excel shape front back programmatically | Aspose.Cells drawing API
// Common Searches: Aspose.Cells change shape Z order C# | bring shape to front Aspose.Cells | send shape to back Aspose.Cells .NET | how to reorder overlapping shapes in Excel using Aspose | ZOrderPosition example Aspose.Cells
// Developer Intent: The developer needs to programmatically control the stacking order of multiple overlapping shapes in an Excel workbook created with Aspose.Cells, and verify the visual result after reordering.
// Use Cases: Create diagrams where background graphics stay behind foreground annotations. | Adjust chart callouts or data markers so key information appears on top of other elements. | Generate Excel templates with layered images that can be reordered at runtime based on user input.
// AI Prompts: Show C# code to move a specific shape one level forward using Aspose.Cells. | Provide a method to reset all shapes to their original ZOrderPosition after changes. | Explain how positive and negative values affect ToFrontOrBack in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsZOrderDemo
{
    // Demonstrates how to create three overlapping rectangle shapes in a workbook, assign explicit ZOrderPosition values, and use ToFrontOrBack to move shapes forward or backward, revealing the layering effect when the file is opened in Excel.
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
                Shape shape1 = worksheet.Shapes.AddRectangle(10, 10, 100, 100, 0, 0);
                Shape shape2 = worksheet.Shapes.AddRectangle(30, 30, 100, 100, 0, 0);
                Shape shape3 = worksheet.Shapes.AddRectangle(50, 50, 100, 100, 0, 0);

                // Set initial Z-order positions: back (0), middle (1), front (2)
                shape1.ZOrderPosition = 0;
                shape2.ZOrderPosition = 1;
                shape3.ZOrderPosition = 2;

                // Display initial Z-order positions
                Console.WriteLine($"Initial ZOrder - Shape1: {shape1.ZOrderPosition}, Shape2: {shape2.ZOrderPosition}, Shape3: {shape3.ZOrderPosition}");

                // Bring shape1 to the front (move forward by 2 positions, the maximum possible)
                shape1.ToFrontOrBack(2);
                Console.WriteLine($"After bringing Shape1 forward: Shape1 ZOrder = {shape1.ZOrderPosition}");

                // Send shape3 to the back (move backward by 2 positions, the maximum possible)
                shape3.ToFrontOrBack(-2);
                Console.WriteLine($"After sending Shape3 backward: Shape3 ZOrder = {shape3.ZOrderPosition}");

                // Ensure the output directory exists
                string outputPath = "ShapeZOrderDemo.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook to observe the layering effect in Excel
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
