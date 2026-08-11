// Title: C# – Log a warning when a worksheet’s Shape count exceeds a threshold using Aspose.Cells
// Description: Creates a workbook, adds several rectangle shapes, iterates through each worksheet, compares Shapes.Count to a configurable limit, writes a console warning for sheets that exceed the limit, and saves the workbook.
// Keywords: Aspose.Cells | .NET | C# | shape count | threshold warning | heavy graphics | worksheet shapes | Excel performance | Shapes.Count | logging
// Common Searches: Aspose.Cells check shape count | log warning for many shapes in Excel C# | detect heavy graphics worksheet Aspose.Cells | set shape threshold Aspose.Cells | count shapes in worksheet .NET
// Developer Intent: Identify worksheets that contain more shapes than a defined limit and output a warning message.
// Use Cases: Validate worksheets before saving to avoid excessive graphics that inflate file size. | Generate a report of sheets exceeding the shape limit for performance monitoring. | Automate warning logs during batch processing of multiple workbooks.
// AI Prompts: Write C# code with Aspose.Cells that scans all worksheets and logs a warning when Shapes.Count is greater than a configurable threshold. | Show an example that logs the worksheet name and shape count when the number of shapes exceeds 10, then saves the workbook. | Explain how to replace Console.WriteLine with a structured logging framework and customize the warning format in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds several rectangle shapes, iterates through each worksheet, compares Shapes.Count to a configurable limit, writes a console warning for sheets that exceed the limit, and saves the workbook.
    public class ShapeCountWarningDemo
    {
        // Define the threshold for heavy graphics
        private const int ShapeCountThreshold = 5;

        public static void Main()
        {
            // Create a new workbook (lifecycle create rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "GraphicsSheet";

            // Add shapes to the worksheet to simulate heavy graphics
            // Adding more than the threshold number of shapes
            for (int i = 0; i < 8; i++)
            {
                // Add a simple rectangle shape
                // Parameters: upperLeftRow, top, upperLeftColumn, left, height, width
                sheet.Shapes.AddRectangle(i, 0, i, 0, 50, 100);
            }

            // Check each worksheet for shape count exceeding the threshold
            foreach (Worksheet ws in workbook.Worksheets)
            {
                int shapeCount = ws.Shapes.Count;
                if (shapeCount > ShapeCountThreshold)
                {
                    // Log a warning indicating heavy graphics
                    Console.WriteLine($"Warning: Worksheet \"{ws.Name}\" contains {shapeCount} shapes, which exceeds the threshold of {ShapeCountThreshold}.");
                }
            }

            // Save the workbook (lifecycle save rule)
            workbook.Save("ShapeCountWarningDemo.xlsx");
        }
    }
}
