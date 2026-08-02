// Title: Move a Shape to Front, Retrieve Z‑Order Position, and Log It with Aspose.Cells for .NET
// Description: C# sample that creates a workbook, adds overlapping rectangles, calls Shape.ToFrontOrBack(1) to bring a shape forward, reads its ZOrderPosition, and writes the index with a timestamp to a log file before saving the workbook.
// Keywords: Aspose.Cells | C# | shape ZOrderPosition | ToFrontOrBack | move shape to front | log shape order | worksheet shapes | Aspose.Cells API | Z‑order index | shape layering
// Common Searches: Aspose.Cells bring shape to front | Get ZOrderPosition of a shape in C# | Log shape Z‑order index Aspose.Cells | ToFrontOrBack method example | How to record shape order in Excel using Aspose
// Developer Intent: The developer needs to promote a specific shape to the top layer, capture its Z‑order value, and persist that information in a log file.
// Use Cases: Validate that a watermark appears above all other objects by moving it forward and recording its Z‑order. | Audit the rendering sequence of overlapping charts in an automated report by logging each shape’s Z‑order after reordering. | Create a troubleshooting log that captures shape layer changes during dynamic workbook generation.
// AI Prompts: Generate C# code that moves a shape backward, reads its ZOrderPosition, and appends the value to a CSV file using Aspose.Cells. | Show how to iterate through all worksheet shapes, bring each to the front, and log their Z‑order indices with timestamps. | Explain the effect of positive and negative arguments in Shape.ToFrontOrBack and how to retrieve the resulting ZOrderPosition.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeZOrderDemo
{
    // C# sample that creates a workbook, adds overlapping rectangles, calls Shape.ToFrontOrBack(1) to bring a shape forward, reads its ZOrderPosition, and writes the index with a timestamp to a log file before saving the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add two overlapping rectangles to demonstrate Z‑order manipulation
            Shape shape1 = worksheet.Shapes.AddRectangle(10, 10, 100, 100, 0, 0);
            Shape shape2 = worksheet.Shapes.AddRectangle(50, 50, 100, 100, 0, 0);

            // Bring shape2 to the front (positive value moves forward)
            shape2.ToFrontOrBack(1);

            // Capture the Z‑order position of the shape that was moved to front
            int zOrderIndex = shape2.ZOrderPosition;

            // Log the Z‑order index to a text file
            string logPath = "ShapeZOrderLog.txt";
            using (StreamWriter writer = new StreamWriter(logPath, true))
            {
                writer.WriteLine($"{DateTime.Now}: Shape2 ZOrderPosition = {zOrderIndex}");
            }

            // Save the workbook (optional, just to persist the shapes)
            workbook.Save("ShapeZOrderDemo.xlsx");
        }
    }
}
