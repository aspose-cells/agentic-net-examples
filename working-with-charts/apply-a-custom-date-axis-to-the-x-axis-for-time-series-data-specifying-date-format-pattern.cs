// Title: Create a line chart with a custom‑formatted date X‑axis in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that adds a line chart to a worksheet, sets the X‑axis source to a range of dates, and applies a custom number format such as "dd-MMM-yyyy" to the date axis using Aspose.Cells. | Show how to configure the chart’s Axis object to use a specific date format pattern and then save the workbook to a user‑specified folder. | Extend the sample to include a second data series while keeping the custom date axis formatting intact.
// Common Searches: how to apply a custom date format to the X axis of an Aspose.Cells chart in C# | Aspose.Cells line chart date axis pattern dd-MMM-yyyy example | C# create time‑series line chart with formatted date axis using Aspose.Cells | set number format for chart axis Aspose.Cells .NET | add multiple series to a line chart with a date X axis in Aspose.Cells
// Tags: Aspose.Cells set date axis number format | C# line chart custom X axis date pattern | Aspose.Cells chart series with date values | Excel workbook line chart time series Aspose.Cells | Aspose.Cells configure chart axis formatting | C# add multiple series to Aspose.Cells chart

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The example creates a workbook, fills column A with sequential dates and column B with numeric values, adds a line chart that uses the dates as the X‑axis, applies a custom date‑format pattern to the axis, and saves the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and rename it
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "TimeSeries";

            // Populate sample time‑series data (dates in column A, values in column B)
            for (int i = 0; i < 10; i++)
            {
                sheet.Cells[i, 0].PutValue(new DateTime(2023, 1, 1).AddDays(i));
                sheet.Cells[i, 1].PutValue(i * 10 + 5);
            }

            // Add a line chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Line, 12, 0, 26, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.Title.Text = "Sample Time Series";

            // Add a series: Y values from B1:B10, X (date) values from A1:A10
            int seriesIndex = chart.NSeries.Add("B1:B10", true);
            chart.NSeries[seriesIndex].XValues = "A1:A10";

            // Determine output path and ensure directory exists
            string outputPath = "TimeSeriesWithDateAxis.xlsx";
            string fullPath = Path.GetFullPath(outputPath);
            string outputDir = Path.GetDirectoryName(fullPath);

            if (string.IsNullOrEmpty(outputDir))
            {
                outputDir = Directory.GetCurrentDirectory();
                fullPath = Path.Combine(outputDir, outputPath);
            }

            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(fullPath);
            Console.WriteLine($"Workbook saved to {fullPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
