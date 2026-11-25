// Title: How to link a chart series' data label number format to its source column using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set data label number format from source cell format | programmatically copy custom number format from a column to chart series labels in .NET | retrieve number format of a range and apply it to Excel chart data labels with Aspose.Cells | link chart series values to data label formatting using Aspose.Cells API | C# example for syncing chart data label format with source cells in Aspose.Cells
// Tags: chart series data label formatting Aspose.Cells | apply source cell number format to chart labels C# | sync data label format with column values Aspose.Cells | retrieve custom number format from range Aspose.Cells | disable data label number format linking Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSeriesDataLabelNumberFormat
{
    // The example creates a workbook with a custom number format on a data column, adds a column chart, then loads the workbook, iterates through worksheets, charts, and series, extracts each series' source range, reads the first cell's style to obtain its custom or built‑in number format, assigns that format to the series' data labels (turning off automatic linking), and saves the updated workbook.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Ensure the input file exists; if not, create a simple workbook for demonstration.
                if (!File.Exists(inputPath))
                {
                    var demoWb = new Workbook();
                    var demoSheet = demoWb.Worksheets[0];
                    demoSheet.Name = "Demo";

                    // Sample data
                    demoSheet.Cells["A1"].PutValue("Category");
                    demoSheet.Cells["A2"].PutValue("A");
                    demoSheet.Cells["A3"].PutValue("B");
                    demoSheet.Cells["A4"].PutValue("C");

                    demoSheet.Cells["B1"].PutValue("Value");
                    demoSheet.Cells["B2"].PutValue(10);
                    demoSheet.Cells["B3"].PutValue(20);
                    demoSheet.Cells["B4"].PutValue(30);

                    // Apply a custom number format to the value column
                    var style = demoWb.CreateStyle();
                    style.Custom = "#,##0.00";
                    var flag = new StyleFlag { NumberFormat = true };
                    demoSheet.Cells.CreateRange("B2:B4").ApplyStyle(style, flag);

                    // Add a chart
                    int chartIdx = demoSheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                    Chart chart = demoSheet.Charts[chartIdx];
                    chart.NSeries.Add("B2:B4", true);
                    chart.NSeries.CategoryData = "A2:A4";

                    demoWb.Save(inputPath);
                    Console.WriteLine($"Demo workbook created at \"{inputPath}\".");
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all charts on the worksheet
                    foreach (Chart chart in sheet.Charts)
                    {
                        // Iterate through each series in the chart
                        foreach (Series series in chart.NSeries)
                        {
                            // Obtain the source range for the series values.
                            // The Series.Values property holds the source range string (e.g., "B2:B10").
                            string valuesRange = series.Values;

                            if (string.IsNullOrEmpty(valuesRange))
                            {
                                Console.WriteLine($"[Warning] Could not determine source range for series \"{series.Name}\" in chart \"{chart.Name}\". Skipping.");
                                continue;
                            }

                            // Create a range object for the source cells
                            Aspose.Cells.Range sourceRange;
                            try
                            {
                                sourceRange = sheet.Cells.CreateRange(valuesRange);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"[Error] Invalid range \"{valuesRange}\" for series \"{series.Name}\": {ex.Message}");
                                continue;
                            }

                            // Retrieve the number format from the first cell of the range.
                            Cell firstCell = sourceRange[0, 0];
                            Style cellStyle = firstCell.GetStyle();

                            // Prefer custom format; fall back to built‑in format number.
                            string numberFormat = cellStyle.Custom;
                            if (string.IsNullOrEmpty(numberFormat))
                            {
                                numberFormat = cellStyle.Number.ToString();
                            }

                            // Apply the retrieved format to the series data labels.
                            DataLabels dataLabels = series.DataLabels;
                            dataLabels.NumberFormat = numberFormat;
                            dataLabels.NumberFormatLinked = false;

                            Console.WriteLine($"Series \"{series.Name}\" in chart \"{chart.Name}\" on sheet \"{sheet.Name}\" linked to number format \"{numberFormat}\".");
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[Fatal] An unexpected error occurred: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
