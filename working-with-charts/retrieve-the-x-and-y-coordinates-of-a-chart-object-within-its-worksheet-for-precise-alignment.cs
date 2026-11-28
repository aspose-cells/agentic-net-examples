// Title: How to read and modify the X/Y pixel position of a chart object in an Excel worksheet using Aspose.Cells for .NET
// Common Searches: c# get chartobject X coordinate from worksheet with Aspose.Cells | how to obtain chart Y position in pixels using Aspose.Cells .NET | move an Excel chart programmatically by pixels with Aspose.Cells | read chart shape location and shift it for alignment in C# | Aspose.Cells example for retrieving and setting chart coordinates
// Tags: chartobject x property aspnet | chartobject y property aspnet | retrieve chart shape coordinates c# | adjust excel chart pixel offset aspnet | excel chart alignment using Aspose.Cells | chart position manipulation c#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace ChartCoordinateRetriever
{
    // The sample loads or creates a workbook with a column chart, reads the chart's X and Y pixel coordinates via the ChartObject, prints them, moves the chart by a specified offset, and saves the workbook with the updated position.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the workbook. If a path is supplied as a command‑line argument it will be used,
            // otherwise a temporary workbook with a sample chart will be created.
            string workbookPath = args.Length > 0 ? args[0] : "SampleChartWorkbook.xlsx";

            Workbook workbook = null;

            try
            {
                if (File.Exists(workbookPath))
                {
                    // Load existing workbook
                    workbook = new Workbook(workbookPath);
                    Console.WriteLine($"Loaded workbook from \"{workbookPath}\".");
                }
                else
                {
                    // Create a new workbook with sample data and a chart
                    workbook = new Workbook();
                    Worksheet sheet = workbook.Worksheets[0];

                    // Populate sample data
                    sheet.Cells["A1"].PutValue("Category");
                    sheet.Cells["A2"].PutValue("A");
                    sheet.Cells["A3"].PutValue("B");
                    sheet.Cells["A4"].PutValue("C");
                    sheet.Cells["B1"].PutValue("Value");
                    sheet.Cells["B2"].PutValue(10);
                    sheet.Cells["B3"].PutValue(20);
                    sheet.Cells["B4"].PutValue(30);

                    // Add a chart (Column type) positioned at rows 5‑15 and columns 0‑5
                    int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                    Chart chart = sheet.Charts[chartIndex];
                    chart.NSeries.Add("B2:B4", true);
                    chart.NSeries.CategoryData = "A2:A4";

                    // Explicitly set known coordinates so the demo is deterministic
                    chart.ChartObject.X = 120;   // pixels from the left edge of the worksheet
                    chart.ChartObject.Y = 250;   // pixels from the top edge of the worksheet

                    // Save the temporary workbook for inspection
                    workbook.Save(workbookPath, SaveFormat.Xlsx);
                    Console.WriteLine($"Created temporary workbook with a chart at \"{workbookPath}\".");
                }

                // Retrieve the first worksheet and its first chart
                Worksheet worksheet = workbook.Worksheets[0];
                if (worksheet.Charts.Count == 0)
                {
                    Console.WriteLine("No charts found in the first worksheet.");
                    return;
                }

                Chart firstChart = worksheet.Charts[0];

                // Access the ChartObject (ChartShape) to get X and Y coordinates
                ChartShape chartShape = firstChart.ChartObject;

                // X and Y are expressed in points (1 point = 1/72 inch) but Aspose.Cells treats them as pixels.
                int chartX = chartShape.X;
                int chartY = chartShape.Y;

                Console.WriteLine("Chart Coordinates:");
                Console.WriteLine($"  X (horizontal offset) = {chartX} pixels");
                Console.WriteLine($"  Y (vertical offset)   = {chartY} pixels");

                // Demonstrate precise alignment by moving the chart 20 pixels right and 10 pixels down
                chartShape.X = chartX + 20;
                chartShape.Y = chartY + 10;

                Console.WriteLine("Chart moved for alignment:");
                Console.WriteLine($"  New X = {chartShape.X} pixels");
                Console.WriteLine($"  New Y = {chartShape.Y} pixels");

                // Save the workbook with updated coordinates
                string outputPath = "ChartWithUpdatedCoordinates.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved with updated chart position to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while processing the workbook:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
            finally
            {
                // Dispose the workbook if it implements IDisposable
                workbook?.Dispose();
            }
        }
    }
}
