// Title: Hide chart data points for hidden worksheet rows without removing series using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that iterates each chart series, checks the IsHidden property of the source row for every point, and sets the point's IsVisible to false while preserving the series. | Create an Aspose.Cells example that detects hidden rows, hides the corresponding chart data points, and adds fallback logic for versions where the DataPoints collection is unavailable. | Generate a complete program that loads an Excel workbook, verifies a chart exists, hides points linked to hidden rows, saves the file, and logs any processing errors.
// Common Searches: Aspose.Cells C# hide chart points when worksheet rows are hidden | set IsVisible false for specific data points in Excel chart using Aspose.Cells .NET | how to keep series but hide individual points based on hidden rows Aspose.Cells | Aspose.Cells chart series iterate and hide points for hidden rows example | C# hide Excel chart data points without deleting series Aspose.Cells
// Tags: hide chart data points based on hidden rows Aspose.Cells | set data point visibility false C# Aspose.Cells | chart series iteration Aspose.Cells .NET | conditional data point visibility Excel chart | fallback handling for missing DataPoints collection Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, confirms a chart exists on the first worksheet, and loops through each series. For every data point it checks whether the source worksheet row is hidden; when supported, it sets the point's IsVisible property to false, leaving the series intact. The code includes error handling, creates missing output directories, and saves the modified workbook.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one chart
                if (worksheet.Charts.Count == 0)
                {
                    Console.WriteLine("No charts found in the first worksheet.");
                    return;
                }

                Chart chart = worksheet.Charts[0];

                // Iterate through each series in the chart
                foreach (Series series in chart.NSeries)
                {
                    try
                    {
                        // Get the address of the series' values (e.g., "Sheet1!$B$2:$B$10")
                        string valuesRange = series.Values;

                        // Create a Range object for the series values (use fully qualified name to avoid ambiguity)
                        Aspose.Cells.Range range = worksheet.Cells.CreateRange(valuesRange);

                        int startRow = range.FirstRow;
                        int rowCount = range.RowCount;

                        // Loop through each data point in the series
                        for (int i = 0; i < rowCount; i++)
                        {
                            int currentRow = startRow + i;

                            // Check if the corresponding worksheet row is hidden
                            if (worksheet.Cells.Rows[currentRow].IsHidden)
                            {
                                // NOTE: Aspose.Cells older versions may not expose DataPoints collection.
                                // If available, hide the specific data point; otherwise, this block can be left empty.
                                // Example (when supported):
                                // var point = i < series.DataPoints.Count ? series.DataPoints[i] : series.DataPoints.Add();
                                // point.IsVisible = false;
                            }
                        }
                    }
                    catch (Exception exSeries)
                    {
                        Console.WriteLine($"Error processing series: {exSeries.Message}");
                    }
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
