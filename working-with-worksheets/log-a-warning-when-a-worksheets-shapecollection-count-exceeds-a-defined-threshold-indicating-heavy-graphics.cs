// Title: Detect and log a warning in C# with Aspose.Cells when a worksheet's shape collection exceeds a performance threshold
// Common Searches: Aspose.Cells C# check worksheet shape count and log warning | how to monitor number of shapes in an Excel worksheet using Aspose.Cells | log performance warning for worksheets with many shapes in C# Aspose.Cells | C# example for shape collection threshold validation with Aspose.Cells | detect heavy graphics in Excel files using Aspose.Cells C#
// Tags: worksheet shape collection count check Aspose.Cells | log warning for excessive shapes C# | performance monitoring of Excel graphics Aspose.Cells | shape count threshold validation Aspose.Cells | detect heavy graphics in Excel using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeCollectionMonitor
{
    // The sample loads an Excel workbook, iterates through each worksheet, retrieves the Shapes collection, and compares its Count to a predefined threshold (100). It writes a warning to the console when the shape count exceeds the threshold, indicating potential performance issues due to heavy graphics, otherwise it logs an informational message. Errors during processing are caught and displayed.
    class Program
    {
        // Threshold for heavy graphics (number of shapes in a worksheet)
        private const int ShapeCountThreshold = 100;

        static void Main(string[] args)
        {
            // Determine the input workbook path
            string workbookPath;
            if (args.Length > 0 && !string.IsNullOrWhiteSpace(args[0]))
            {
                workbookPath = args[0];
            }
            else
            {
                Console.WriteLine("Please provide the path to an Excel workbook as the first argument.");
                return;
            }

            // Validate the file existence
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Error: The file \"{workbookPath}\" does not exist.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(workbookPath);

                // Iterate through each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Access the ShapeCollection of the worksheet
                    ShapeCollection shapes = sheet.Shapes;

                    // Get the count of shapes
                    int shapeCount = shapes.Count;

                    // Log a warning if the count exceeds the defined threshold
                    if (shapeCount > ShapeCountThreshold)
                    {
                        Console.WriteLine($"WARNING: Worksheet \"{sheet.Name}\" contains {shapeCount} shapes, which exceeds the threshold of {ShapeCountThreshold}. This may indicate heavy graphics and could affect performance.");
                    }
                    else
                    {
                        Console.WriteLine($"INFO: Worksheet \"{sheet.Name}\" contains {shapeCount} shapes (within acceptable limits).");
                    }
                }
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
                Console.WriteLine($"Stack Trace: {ex.StackTrace}");
            }
        }
    }
}
