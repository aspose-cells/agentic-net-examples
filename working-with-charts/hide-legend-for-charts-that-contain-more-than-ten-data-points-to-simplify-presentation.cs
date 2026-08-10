// Title: Hide Chart Legend in Aspose.Cells (C#) When a Series Has More Than 10 Data Points
// Description: C# example that opens an Excel workbook, scans every worksheet and chart, counts non‑empty cells in the first series range, disables the legend if the count exceeds ten, and saves the modified file. Ideal for batch‑processing reports where large charts need a cleaner look.
// Keywords: Aspose.Cells | C# chart legend | hide legend conditionally | Excel chart data point count | series values range | disable legend Aspose.Cells | .NET Excel automation | chart formatting | batch workbook processing | Excel reporting
// Common Searches: Aspose.Cells hide legend for large chart C# | how to hide Excel chart legend when data points > 10 using .NET | count non‑empty cells in chart series Aspose.Cells | conditional legend visibility in Excel with Aspose.Cells | C# example to remove chart legend based on series size
// Developer Intent: Automatically suppress the legend of any chart whose first series contains more than ten data points.
// Use Cases: Generate financial or scientific reports where dense charts are displayed without legends to avoid clutter. | Create dashboards that show legends only for simple charts, improving readability for end users. | Process multiple workbooks in a scheduled job, adjusting chart appearance based on data volume.
// AI Prompts: Write Aspose.Cells C# code that hides the legend for every chart whose first series has over ten non‑empty cells across all worksheets. | Suggest a faster way to determine the number of data points in a chart series without looping through each cell. | Explain how to extend the logic to evaluate all series in a chart and hide the legend when the total data points exceed ten.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using AsposeRange = Aspose.Cells.Range;

namespace HideLegendForLargeCharts
{
    // C# example that opens an Excel workbook, scans every worksheet and chart, counts non‑empty cells in the first series range, disables the legend if the count exceeds ten, and saves the modified file. Ideal for batch‑processing reports where large charts need a cleaner look.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                // Ensure the input file exists; create a blank workbook if it does not.
                Workbook workbook;
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook();
                    workbook.Worksheets[0].Name = "Sheet1";
                }

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all charts on the worksheet
                    foreach (Chart chart in sheet.Charts)
                    {
                        // Ensure the chart has at least one series
                        if (chart.NSeries.Count > 0)
                        {
                            // Get the first series (extend as needed)
                            Series series = chart.NSeries[0];

                            // The data range used for the series values (e.g., "B2:B15")
                            string valuesRange = series.Values;

                            // Obtain the range object from the worksheet
                            AsposeRange range = sheet.Cells.CreateRange(valuesRange);

                            // Count non‑empty data points in the range
                            int dataPointCount = 0;
                            foreach (Cell cell in range)
                            {
                                // Consider a cell as a data point if it contains a value
                                if (cell.Value != null && !string.IsNullOrEmpty(cell.StringValue))
                                {
                                    dataPointCount++;
                                }
                            }

                            // Hide the legend if the chart contains more than ten data points
                            if (dataPointCount > 10)
                            {
                                chart.ShowLegend = false;
                            }
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
