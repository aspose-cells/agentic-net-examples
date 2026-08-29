// Title: Generate a month-level Timeline linked to a PivotTable and export the line chart as a 300 dpi LZW-compressed TIFF using Aspose.Cells for .NET
// AI Prompts: Write a C# program that creates a workbook, fills it with date/value rows, builds a PivotTable, adds a Timeline set to monthly intervals, creates a time-scale line chart, and saves the chart as a 300 dpi LZW-compressed TIFF with Aspose.Cells. | Modify the example to hide weekend dates on the Timeline while keeping the monthly tick marks, then export the chart to a high-quality TIFF file. | Configure ImageOrPrintOptions to set both horizontal and vertical resolution to 300 dpi and enable LZW compression for the TIFF output of the chart.
// Common Searches: Aspose.Cells C# export line chart to TIFF with 300 DPI resolution | How to hide weekends on a Timeline in Aspose.Cells | Create a month-level Timeline linked to a PivotTable using Aspose.Cells for .NET | Set LZW compression for TIFF output in Aspose.Cells ImageOrPrintOptions | Configure time-scale axis with monthly major units in Aspose.Cells chart
// Tags: Aspose.Cells monthly timeline ticks | pivot table timeline binding Aspose.Cells | TIFF image export with LZW Aspose.Cells | timeline hide weekends Aspose.Cells | time-scale line chart configuration Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;
using Aspose.Cells.Rendering;

namespace AsposeCellsExamples
{
    // // Demonstrates creating a workbook, populating it with date/value data, building a PivotTable, adding a month-level Timeline, configuring a time-scale line chart, and exporting the chart to a 300 dpi LZW-compressed TIFF file using Aspose.Cells for .NET.
    public class TimelineWithMonthlyTicksAndTiffExport
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // -------------------------------------------------
                // Populate worksheet with sample date/value data
                // -------------------------------------------------
                sheet.Cells["A1"].PutValue("Date");
                sheet.Cells["B1"].PutValue("Value");

                DateTime startDate = new DateTime(2023, 1, 1);
                for (int i = 0; i < 60; i++) // 60 days ~ 2 months
                {
                    sheet.Cells[i + 1, 0].PutValue(startDate.AddDays(i));
                    sheet.Cells[i + 1, 1].PutValue(i * 10 + 50);
                }

                // -------------------------------------------------
                // Create a PivotTable that will be the data source for the Timeline
                // -------------------------------------------------
                PivotTableCollection pivots = sheet.PivotTables;
                int pivotIdx = pivots.Add("A1:B61", "D1", "PivotTable1");
                PivotTable pivot = pivots[pivotIdx];
                pivot.AddFieldToArea(PivotFieldType.Row, "Date");
                pivot.AddFieldToArea(PivotFieldType.Data, "Value");

                // Refresh and calculate the pivot data
                pivot.RefreshData();      // Updated API call
                pivot.CalculateData();

                // -------------------------------------------------
                // Add a Timeline linked to the PivotTable
                // -------------------------------------------------
                // Place the Timeline at row 10, column 5 (zero‑based indexes)
                int timelineIdx = sheet.Timelines.Add(pivot, 10, 5, "Date");
                Timeline timeline = sheet.Timelines[timelineIdx];
                timeline.CurrentLevel = TimelineLevelType.Month; // Monthly ticks

                // -------------------------------------------------
                // Create a Line chart that uses the same date range
                // -------------------------------------------------
                int chartIdx = sheet.Charts.Add(ChartType.Line, 20, 0, 35, 15);
                Chart chart = sheet.Charts[chartIdx];
                chart.NSeries.Add("B2:B61", true);
                chart.NSeries.CategoryData = "A2:A61";

                // Configure the category axis as a time scale with monthly major units
                Axis categoryAxis = chart.CategoryAxis;
                categoryAxis.CategoryType = CategoryType.TimeScale;
                categoryAxis.BaseUnitScale = TimeUnit.Months;
                categoryAxis.MajorUnitScale = TimeUnit.Months;
                categoryAxis.MinorUnitScale = TimeUnit.Days;

                // -------------------------------------------------
                // Export the chart to a high‑quality TIFF image
                // -------------------------------------------------
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    // Image format is inferred from the file extension; no need to set ImageFormat explicitly
                    TiffCompression = TiffCompression.CompressionLZW,
                    HorizontalResolution = 300,
                    VerticalResolution = 300
                };

                string tiffPath = "TimelineChart.tiff";

                try
                {
                    chart.ToImage(tiffPath, imgOptions);
                    Console.WriteLine($"Chart exported successfully to: {tiffPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to export chart: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Run error: {ex.Message}");
            }
        }
    }
}
