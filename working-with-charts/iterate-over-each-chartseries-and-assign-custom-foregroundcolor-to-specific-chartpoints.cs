// Title: Set custom foreground colors for each chart point in an Excel workbook with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code using Aspose.Cells to loop through all worksheets, locate every chart, and assign Red to the first point, Blue to the second, and Green to all remaining points by setting point.Area.ForegroundColor. | Update an existing Aspose.Cells workbook so that each chart series' points are colored based on their index, handling multiple worksheets and charts, then save the modified file.
// Common Searches: asp.net change color of individual data points in Excel chart using Aspose.Cells | c# set foreground color for chart point Aspose.Cells example | iterate over chart series and points to apply custom colors Aspose.Cells | how to color first two points differently in Excel chart with Aspose.Cells | Aspose.Cells programmatically customize chart point colors in .xlsx file
// Tags: Aspose.Cells chart point foreground color | C# iterate chart series Aspose.Cells | set individual chart point color Excel | dynamic series points Aspose.Cells | customize chart colors .xlsx Aspose.Cells

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, iterates through each worksheet and its charts, walks every series and its points, and sets the Area.ForegroundColor of the first point to red, the second to blue, and all other points to green, then saves the updated workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through worksheets and their charts
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    foreach (Chart chart in sheet.Charts)
                    {
                        try
                        {
                            // Iterate through each series in the chart
                            for (int s = 0; s < chart.NSeries.Count; s++)
                            {
                                // Use dynamic to avoid version‑specific type issues
                                dynamic series = chart.NSeries[s];

                                // Process each point in the series
                                for (int i = 0; i < series.Points.Count; i++)
                                {
                                    dynamic point = series.Points[i];

                                    if (i == 0)
                                    {
                                        // First point – Red
                                        point.Area.ForegroundColor = Color.Red;
                                    }
                                    else if (i == 1)
                                    {
                                        // Second point – Blue
                                        point.Area.ForegroundColor = Color.Blue;
                                    }
                                    else
                                    {
                                        // Other points – Green
                                        point.Area.ForegroundColor = Color.Green;
                                    }
                                }
                            }
                        }
                        catch (Exception exChart)
                        {
                            Console.WriteLine($"Error processing chart on sheet '{sheet.Name}': {exChart.Message}");
                        }
                    }
                }

                // Save the modified workbook
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
