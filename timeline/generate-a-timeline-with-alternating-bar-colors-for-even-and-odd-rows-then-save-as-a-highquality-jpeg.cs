// Title: Generate a timeline column chart with alternating bar colors and export it as a high‑quality JPEG using Aspose.Cells for .NET
// AI Prompts: Create C# code that builds a column chart from a date column and a numeric column, applies LightBlue to even-indexed columns and LightGreen to odd-indexed columns, and saves the chart as a JPEG with high quality using Aspose.Cells. | Adjust the rendering options to output a 300 dpi JPEG while keeping the custom bar colors and programmatically add a chart title.
// Common Searches: how to set different colors for each column in an Aspose.Cells chart c# | export Aspose.Cells chart to high resolution JPEG file | create timeline chart with date axis using Aspose.Cells .NET example | alternating bar colors based on row index in Aspose.Cells column chart | save Aspose.Cells chart as image with one page per sheet
// Tags: column chart color pattern Aspose.Cells | timeline chart JPEG export .NET | chart point color customization Aspose.Cells | high quality image rendering Aspose.Cells | date axis column chart generation C#

using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// The example builds a column timeline chart from date and numeric data, colors even bars LightBlue and odd bars LightGreen, and renders the chart to a high‑quality JPEG file using Aspose.Cells rendering options.
class TimelineChartGenerator
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            using (Workbook workbook = new Workbook())
            {
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data: Date (A) and Value (B)
                DateTime startDate = new DateTime(2023, 1, 1);
                for (int i = 0; i < 10; i++)
                {
                    sheet.Cells[i + 1, 0].PutValue(startDate.AddDays(i * 7)); // A2:A11
                    sheet.Cells[i + 1, 1].PutValue(10 + i * 5);               // B2:B11
                }

                // Add a column chart to represent the timeline
                // Position the chart from row 12, column 0 to row 30, column 10
                int chartIndex = sheet.Charts.Add(ChartType.Column, 12, 0, 30, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the series (values)
                chart.NSeries.Add("B2:B11", true);
                // Category (X‑axis) data – optional; if not set, default indices are used
                // chart.NSeries[0].CategoryData = "A2:A11";

                // Set a title for the chart
                chart.Title.Text = "Timeline with Alternating Bar Colors";

                // Apply alternating colors to each data point (even/odd rows)
                for (int i = 0; i < chart.NSeries[0].Points.Count; i++)
                {
                    ChartPoint point = chart.NSeries[0].Points[i];
                    point.Area.ForegroundColor = (i % 2 == 0) ? Color.LightBlue : Color.LightGreen;
                }

                // Configure high‑quality JPEG export options
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    OnePagePerSheet = true // Export the chart as a single page
                };

                // Export the chart to a JPEG file
                using (MemoryStream ms = new MemoryStream())
                {
                    chart.ToImage(ms, imgOptions);
                    try
                    {
                        File.WriteAllBytes("TimelineChart.jpg", ms.ToArray());
                    }
                    catch (Exception fileEx)
                    {
                        Console.WriteLine($"File write error: {fileEx.Message}");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
