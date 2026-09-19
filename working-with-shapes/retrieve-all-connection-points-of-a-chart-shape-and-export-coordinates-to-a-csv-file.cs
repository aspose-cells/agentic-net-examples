// Title: Export chart shape connection points to a CSV file with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an Excel workbook using Aspose.Cells, iterates over each chart, calls Shape.GetConnectionPoints, and writes the point index with X/Y coordinates to a CSV file. | Enhance the example to also record the chart's top‑left cell address and the worksheet name in each CSV row. | Add comprehensive error handling that skips charts lacking a shape or connection points and logs descriptive warnings to the console.
// Common Searches: Aspose.Cells C# how to extract chart connection points and save to CSV | GetConnectionPoints method example for Excel charts in .NET | Export chart shape coordinates from an .xlsx file using Aspose.Cells | C# code to list chart connection point coordinates in an Excel workbook
// Tags: Aspose.Cells GetConnectionPoints API | export chart connection points as CSV | chart shape coordinates extraction .NET | Excel chart connection points Aspose.Cells | C# write point data to CSV file

using System;
using System.Drawing;
using System.IO;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// The sample loads 'input.xlsx', accesses the first worksheet, iterates through all charts, obtains each chart's Shape object, retrieves its connection points via GetConnectionPoints, converts them to PointF values, and writes the chart name, point index, and X/Y coordinates to 'ChartConnectionPoints.csv'. It includes file‑existence checks, skips charts without shapes or points, and logs warnings for missing data.
class ExportChartConnectionPoints
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputCsv = "ChartConnectionPoints.csv";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Prepare a CSV file to write the connection points
            using (StreamWriter csvWriter = new StreamWriter(outputCsv))
            {
                // Write CSV header
                csvWriter.WriteLine("ChartName,PointIndex,X,Y");

                // Iterate through all charts in the worksheet
                foreach (Chart chart in sheet.Charts)
                {
                    // Obtain the shape that represents the chart
                    Shape chartShape = chart.ChartObject;
                    if (chartShape == null)
                    {
                        Console.WriteLine($"Warning: Chart \"{chart.Name}\" does not have an associated shape.");
                        continue;
                    }

                    // Retrieve the connection points (float[][]) and convert to PointF[]
                    PointF[] connectionPoints;
                    try
                    {
                        float[][] rawPoints = chartShape.GetConnectionPoints();
                        if (rawPoints == null || rawPoints.Length == 0)
                        {
                            Console.WriteLine($"Info: No connection points found for chart \"{chart.Name}\".");
                            continue;
                        }

                        connectionPoints = rawPoints
                            .Where(p => p != null && p.Length >= 2)
                            .Select(p => new PointF(p[0], p[1]))
                            .ToArray();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Warning: Unable to get connection points for chart \"{chart.Name}\": {ex.Message}");
                        continue;
                    }

                    // Export each point to the CSV file
                    for (int i = 0; i < connectionPoints.Length; i++)
                    {
                        PointF pt = connectionPoints[i];
                        // X and Y are in points (1/72 inch). Adjust if needed.
                        csvWriter.WriteLine($"{chart.Name},{i},{pt.X},{pt.Y}");
                    }
                }
            }

            Console.WriteLine($"Chart connection points have been exported to \"{outputCsv}\"");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine("An error occurred while exporting chart connection points:");
            Console.WriteLine(ex.Message);
        }
    }
}
