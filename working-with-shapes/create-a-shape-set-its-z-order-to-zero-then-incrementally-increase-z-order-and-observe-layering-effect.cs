// Title: Aspose.Cells .NET: Create Overlapping Shapes and Control Z‑Order with ToFrontOrBack
// Description: Sample C# program that creates three overlapping rectangle shapes in a new Excel workbook, reads their initial ZOrderPosition values, moves one shape to the front and another one step back using the ToFrontOrBack method, and saves the file (ZOrderDemo.xlsx) so you can see the layering effect.
// Keywords: Aspose.Cells | C# | .NET | Excel shape Z-order | ToFrontOrBack | ZOrderPosition | overlapping rectangles | shape layering | sample code | GitHub example | programmatic shape ordering
// Common Searches: Aspose.Cells bring shape to front C# | change Z-order of shapes Aspose.Cells .NET | ToFrontOrBack method example | how to layer overlapping shapes in Excel using Aspose | read ZOrderPosition Aspose.Cells
// Developer Intent: Programmatically adjust the stacking order of overlapping shapes in an Excel worksheet.
// Use Cases: Ensure a key rectangle appears above all other graphics by moving it to the front. | Hide a shape behind others by sending it one position back without dropping below zero. | Log ZOrderPosition before and after changes to verify correct layering.
// AI Prompts: Generate C# code that adds multiple shapes to an Aspose.Cells worksheet and cycles each shape to the front using ToFrontOrBack. | Explain the ZOrderPosition property in Aspose.Cells and demonstrate moving shapes forward and backward. | Provide a snippet that checks a shape's current Z-order and prevents it from going below zero when moving it back.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Sample C# program that creates three overlapping rectangle shapes in a new Excel workbook, reads their initial ZOrderPosition values, moves one shape to the front and another one step back using the ToFrontOrBack method, and saves the file (ZOrderDemo.xlsx) so you can see the layering effect.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add three overlapping rectangle shapes
            // Parameters: upperLeftRow, upperLeftColumn, width, height, upperLeftPixelRowOffset, upperLeftPixelColumnOffset
            Shape shape1 = worksheet.Shapes.AddRectangle(10, 10, 100, 100, 0, 0);
            Shape shape2 = worksheet.Shapes.AddRectangle(30, 30, 100, 100, 0, 0);
            Shape shape3 = worksheet.Shapes.AddRectangle(50, 50, 100, 100, 0, 0);

            // Initial Z-order positions (0 = backmost)
            Console.WriteLine($"Initial Z-order positions: shape1={shape1.ZOrderPosition}, shape2={shape2.ZOrderPosition}, shape3={shape3.ZOrderPosition}");

            // Bring shape1 to the front safely using ToFrontOrBack
            // Current max Z-order is 2, so move forward by 2 positions
            shape1.ToFrontOrBack(2);
            Console.WriteLine($"After moving shape1 to front: shape1 ZOrderPosition={shape1.ZOrderPosition}");

            // Send shape2 one position back (cannot go below 0)
            shape2.ToFrontOrBack(-1);
            Console.WriteLine($"After sending shape2 back: shape2 ZOrderPosition={shape2.ZOrderPosition}");

            // Save the workbook to observe the layering effect
            workbook.Save("ZOrderDemo.xlsx");
            Console.WriteLine("Workbook saved as ZOrderDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
