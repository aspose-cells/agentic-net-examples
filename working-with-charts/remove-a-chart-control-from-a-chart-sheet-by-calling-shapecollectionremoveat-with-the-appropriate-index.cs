// Title: Remove a chart shape from a worksheet in Aspose.Cells for .NET using ShapeCollection.RemoveAt
// Common Searches: how to remove a chart shape from an Excel worksheet with Aspose.Cells C# | Aspose.Cells ShapeCollection.RemoveAt example for deleting charts | C# code to delete the first chart on a sheet using Aspose.Cells | remove chart control from chart sheet Aspose.Cells .NET | delete chart shape from workbook programmatically Aspose.Cells
// Tags: ShapeCollection.RemoveAt chart removal | Aspose.Cells chart shape deletion | C# Excel chart removal API | Worksheet shape collection manipulation | Aspose.Cells workbook chart cleanup

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// The sample loads or creates an Excel workbook, accesses the first worksheet, retrieves its Shapes collection, removes the first chart shape with ShapeCollection.RemoveAt(0) when present, and saves the modified file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            Workbook workbook;

            // Load existing workbook or create a new one if the file is missing
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                Worksheet ws = workbook.Worksheets[0];

                // Add a sample chart so that there is at least one shape to remove
                int chartIdx = ws.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                ws.Charts[chartIdx].Name = "SampleChart";

                // Save the newly created workbook as the input template
                workbook.Save(inputPath);
            }

            // Access the first worksheet (assumed to contain the chart)
            Worksheet chartSheet = workbook.Worksheets[0];

            // Get the collection of shapes on the worksheet
            ShapeCollection shapes = chartSheet.Shapes;

            // Remove the first shape if any exist
            if (shapes.Count > 0)
            {
                shapes.RemoveAt(0);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
