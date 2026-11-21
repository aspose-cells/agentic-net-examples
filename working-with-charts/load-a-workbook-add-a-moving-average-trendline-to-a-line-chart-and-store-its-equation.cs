// Title: Add a Moving‑Average Trendline to a Line Chart and Save Its Equation with Aspose.Cells for .NET (C#)
// Common Searches: C# add moving average trendline to Excel line chart using Aspose.Cells | How to get the equation of a chart trendline and write it to a cell with Aspose.Cells .NET | Create a line chart in Aspose.Cells if none exists and then add a trendline | Set moving average period and display equation on a chart with Aspose.Cells C# example | Save workbook after inserting trendline equation using Aspose.Cells for .NET
// Tags: aspocells trendline addition csharp | aspocells moving-average period configuration | aspocells generate chart programmatically | aspocells extract chart equation | aspocells write values to cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an existing workbook, creates a line chart when none is present, binds all worksheet data, adds a moving‑average trendline (period 3) to the first series, enables equation display, writes the equation text into cells A1 and B1, and saves the workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Obtain or create a chart
            Chart chart;
            if (sheet.Charts.Count > 0)
            {
                chart = sheet.Charts[0];
            }
            else
            {
                // Determine data range
                int firstRow = 0;
                int firstColumn = 0;
                int totalRows = sheet.Cells.MaxDataRow + 1;
                int totalColumns = sheet.Cells.MaxDataColumn + 1;

                // Add a line chart
                int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 25, 10);
                chart = sheet.Charts[chartIndex];

                // Bind data to the chart
                var range = sheet.Cells.CreateRange(firstRow, firstColumn, totalRows, totalColumns);
                chart.NSeries.Add(range.RefersTo, true);
            }

            // Ensure the chart has at least one series
            if (chart.NSeries.Count == 0)
                throw new InvalidOperationException("The chart does not contain any series to which a trendline can be added.");

            // Use dynamic to access Trendlines (handles API variations)
            dynamic series = chart.NSeries[0];

            try
            {
                // Add a moving average trendline
                Trendline trendline = series.Trendlines.Add(TrendlineType.MovingAverage);
                trendline.Period = 3;               // Set moving average period
                trendline.DisplayEquation = true;   // Show equation on chart

                // Retrieve the equation (if any)
                string equation = trendline.Name ?? string.Empty;

                // Write the equation to the worksheet
                sheet.Cells["A1"].PutValue("Moving Average Equation:");
                sheet.Cells["B1"].PutValue(equation);
            }
            catch (Exception ex)
            {
                // Trendline feature may not be supported in older versions
                Console.WriteLine($"Trendline could not be added: {ex.Message}");
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
