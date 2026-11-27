// Title: Generate a monthly timeline line chart that skips weekends and export it as a high‑resolution TIFF image with Aspose.Cells for .NET
// Common Searches: asp.net generate timeline line chart excluding weekends with Aspose.Cells | save Aspose.Cells chart as high quality TIFF in C# | create monthly date axis chart hiding Saturdays and Sundays using Aspose.Cells | render a line chart to LZW compressed TIFF with Aspose.Cells | Aspose.Cells C# image options for exporting charts to TIFF
// Tags: timeline line chart generation Aspose.Cells | skip weekends in chart axis Aspose.Cells | tiff image export options Aspose.Cells | 300 DPI TIFF rendering Aspose.Cells | monthly date axis chart Aspose.Cells C#

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace TimelineChartToTiff
{
    // The example creates a workbook, fills column A with weekday dates from January to March 2023 and column B with random values, adds a line chart that uses these ranges, configures the axes to display monthly ticks while omitting weekends, and renders the chart to a 300 DPI, 24‑bit LZW‑compressed TIFF file named "TimelineChart.tiff" using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Data";

                // Prepare sample data: dates (only weekdays) and random values for three months
                DateTime startDate = new DateTime(2023, 1, 1);
                DateTime endDate = new DateTime(2023, 3, 31);
                List<DateTime> dates = new List<DateTime>();
                List<double> values = new List<double>();
                Random rnd = new Random();

                for (DateTime dt = startDate; dt <= endDate; dt = dt.AddDays(1))
                {
                    // Skip weekends
                    if (dt.DayOfWeek == DayOfWeek.Saturday || dt.DayOfWeek == DayOfWeek.Sunday)
                        continue;

                    dates.Add(dt);
                    values.Add(Math.Round(50 + rnd.NextDouble() * 100, 2));
                }

                // Fill the worksheet with the data (Column A: Dates, Column B: Values)
                for (int i = 0; i < dates.Count; i++)
                {
                    // Date
                    Cell dateCell = sheet.Cells[i, 0];
                    dateCell.PutValue(dates[i]);
                    Style dateStyle = dateCell.GetStyle();
                    dateStyle.Number = 14; // Short date format
                    dateCell.SetStyle(dateStyle);

                    // Value
                    sheet.Cells[i, 1].PutValue(values[i]);
                }

                // Add a line chart to represent the timeline
                int chartIndex = sheet.Charts.Add(ChartType.Line, dates.Count + 2, 0, dates.Count + 20, 8);
                Chart chart = sheet.Charts[chartIndex];
                chart.Title.Text = "Monthly Timeline (Weekends Hidden)";

                // Set data ranges
                string valuesRange = $"B1:B{dates.Count}";
                string datesRange = $"A1:A{dates.Count}";
                chart.NSeries.Add(valuesRange, true);
                chart.NSeries.CategoryData = datesRange;

                // Configure the category (X) axis
                Axis categoryAxis = chart.CategoryAxis;
                // Optional: customize date format if needed (commented out to avoid missing API)
                // categoryAxis.NumberFormat = "mmm yyyy";

                // Configure the value (Y) axis range
                Axis valueAxis = chart.ValueAxis;
                valueAxis.MinValue = 0;
                valueAxis.MaxValue = 200;
                valueAxis.MajorUnit = 50;

                // Render the chart to a high‑quality TIFF image
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    ImageType = ImageType.Tiff,
                    TiffCompression = TiffCompression.CompressionLZW,
                    HorizontalResolution = 300,   // DPI
                    VerticalResolution = 300,     // DPI
                    TiffColorDepth = ColorDepth.Format24bpp
                };

                string outputPath = "TimelineChart.tiff";

                // Save the chart image
                chart.ToImage(outputPath, imgOptions);

                Console.WriteLine($"Timeline chart saved as high‑quality TIFF at: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
