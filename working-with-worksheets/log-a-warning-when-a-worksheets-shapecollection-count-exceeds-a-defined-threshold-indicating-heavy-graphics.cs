// Title: Warn When Worksheet Shape Count Exceeds a Threshold Using Aspose.Cells for .NET
// Description: This example creates a workbook, adds several rectangle shapes, then scans each worksheet's ShapeCollection. If the number of shapes surpasses a configurable limit (e.g., 5), a warning is written to the console before the file is saved as HeavyGraphics.xlsx. Useful for spotting worksheets that may degrade performance due to heavy graphics.
// Keywords: Aspose.Cells .NET shape count | worksheet graphics threshold | detect heavy graphics Excel | log warning for excessive shapes | shape collection performance | Excel shape limit Aspose | monitor drawing objects | shape count warning code | Aspose.Cells performance tip | shape threshold example
// Common Searches: how to check shape count in Aspose.Cells worksheet | log warning when worksheet has more than X shapes Aspose.Cells | detect heavy graphics in Excel with Aspose.Cells .NET | set shape collection limit Aspose.Cells | performance impact of many shapes in Aspose.Cells
// Developer Intent: Find worksheets that contain more drawing objects than allowed and generate a warning.
// Use Cases: Flag worksheets that could slow down opening or saving due to a large number of shapes. | Create a console or log report of worksheets exceeding a graphics complexity threshold. | Prevent further processing or trigger cleanup when a workbook contains overly complex graphics.
// AI Prompts: Generate a reusable method that scans all worksheets in an Aspose.Cells workbook and logs a warning if ws.Shapes.Count exceeds a configurable threshold. | Show how to integrate Serilog (or NLog) to record shape‑count warnings for each worksheet in Aspose.Cells. | Provide code that throws a custom ShapeLimitExceededException instead of a console message when the shape count is too high. | Explain how to read the shape threshold from an appsettings.json file and apply it in the warning logic.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// This example creates a workbook, adds several rectangle shapes, then scans each worksheet's ShapeCollection. If the number of shapes surpasses a configurable limit (e.g., 5), a warning is written to the console before the file is saved as HeavyGraphics.xlsx. Useful for spotting worksheets that may degrade performance due to heavy graphics.
class Program
{
    // Define the shape count threshold
    const int ShapeThreshold = 5;

    static void Main()
    {
        // Create a new workbook (lifecycle: create)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add several shapes to the worksheet (demo purpose)
        for (int i = 0; i < 7; i++)
        {
            // Add a rectangle shape at different positions
            sheet.Shapes.AddRectangle(i, 0, i, 0, 50, 50);
        }

        // Iterate through all worksheets and log a warning if shape count exceeds the threshold
        foreach (Worksheet ws in workbook.Worksheets)
        {
            int shapeCount = ws.Shapes.Count;
            if (shapeCount > ShapeThreshold)
            {
                Console.WriteLine($"Warning: Worksheet \"{ws.Name}\" contains {shapeCount} shapes, which exceeds the threshold of {ShapeThreshold}.");
            }
        }

        // Save the workbook (lifecycle: save)
        workbook.Save("HeavyGraphics.xlsx");
    }
}
