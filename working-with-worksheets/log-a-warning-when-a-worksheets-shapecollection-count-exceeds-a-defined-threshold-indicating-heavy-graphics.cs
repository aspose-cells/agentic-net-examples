// Title: Log a warning in C# with Aspose.Cells when a worksheet’s shape collection exceeds a specified threshold
// AI Prompts: Write C# code using Aspose.Cells that scans every worksheet in a workbook and prints a console warning if the worksheet’s Shapes.Count is greater than a given limit. | Modify the example to read the shape count limit from an appsettings.json file and write warning messages to a log file instead of the console. | Create a reusable C# method that returns a list of worksheet names whose shape collection size surpasses a configurable threshold using Aspose.Cells.
// Common Searches: how to iterate worksheets and check shape count with Aspose.Cells C# | detect worksheets with many shapes in Excel using Aspose.Cells and output warnings | log heavy graphics warning for Excel sheets when shape count exceeds 100 in C# | aspnet core read shape count threshold from config and log warning using Aspose.Cells
// Tags: Aspose.Cells worksheet shape count monitoring | C# log warning for excessive shapes in Excel | threshold-based graphics detection with Aspose.Cells | iterate workbook worksheets shapes collection C# | configure shape count limit Aspose.Cells

using System;
using Aspose.Cells;

namespace ShapeCollectionMonitor
{
    // The program loads an Excel workbook, iterates each worksheet, checks the number of shapes via sheet.Shapes.Count, and writes a console warning when the count exceeds the defined threshold (default 100).
    class Program
    {
        // Define the threshold for heavy graphics.
        private const int ShapeCountThreshold = 100;

        static void Main(string[] args)
        {
            // Load an existing workbook. Replace the path with your actual file.
            Workbook workbook = new Workbook("InputWorkbook.xlsx");

            // Iterate through each worksheet in the workbook.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the number of shapes in the current worksheet.
                int shapeCount = sheet.Shapes.Count;

                // Check if the shape count exceeds the defined threshold.
                if (shapeCount > ShapeCountThreshold)
                {
                    // Log a warning indicating heavy graphics on this worksheet.
                    Console.WriteLine($"Warning: Worksheet \"{sheet.Name}\" contains {shapeCount} shapes, which exceeds the threshold of {ShapeCountThreshold}.");
                }
            }

            // Optionally, save the workbook after processing.
            // workbook.Save("OutputWorkbook.xlsx");
        }
    }
}
