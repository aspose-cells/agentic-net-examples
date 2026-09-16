// Title: Create a timeline chart with a logarithmic date axis, custom tick intervals, and export it as PNG using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that uses Aspose.Cells to build a line chart where the X‑axis displays dates on a logarithmic scale, set the log base to 10, define major and minor tick units, and save the chart as a PNG image. | Show how to populate an Excel worksheet with exponentially increasing dates, bind them to a chart series, and configure the category axis for logarithmic scaling in Aspose.Cells. | Demonstrate changing the logarithmic base to 2 and exporting the resulting chart to a JPEG file using Aspose.Cells in a .NET console application.
// Common Searches: aspnet aspose.cells create chart with logarithmic X axis and export to png | c# set major unit on logarithmic date axis in aspose.cells chart | how to render aspose.cells chart as image file png | timeline chart with exponential dates using aspose.cells .net | custom tick intervals on logarithmic axis aspose.cells example
// Tags: logarithmic axis setup Aspose.Cells | chart image export PNG Aspose.Cells | custom tick intervals chart axis C# | exponential date series worksheet Aspose.Cells | line chart with date series Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // The example creates a workbook, fills column A with exponentially spaced dates and column B with sequential values, adds a line chart, configures the category axis for logarithmic scaling (base 10) with custom major and minor tick intervals, titles both axes, and renders the chart directly to a PNG file.
    class TimelineLogScaleExample
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data: dates in column A and values in column B
                // Dates are spaced exponentially to illustrate logarithmic scaling
                DateTime startDate = new DateTime(2020, 1, 1);
                for (int i = 0; i < 10; i++)
                {
                    int daysOffset = (int)Math.Pow(2, i); // 1,2,4,8,...
                    sheet.Cells[i, 0].PutValue(startDate.AddDays(daysOffset));
                    sheet.Cells[i, 1].PutValue(i + 1);
                }

                // Add a Line chart (Timeline chart type is not available in Aspose.Cells)
                // Parameters: upper left row, upper left column, lower right row, lower right column, chart type
                int chartIndex = sheet.Charts.Add(0, 2, 15, 10, (int)ChartType.Line);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the series (X = dates, Y = values)
                chart.NSeries.Add("A1:A10", true);
                chart.NSeries[0].Values = "B1:B10";

                // Configure the X (category) axis for logarithmic scale
                Axis categoryAxis = chart.CategoryAxis;
                categoryAxis.IsLogarithmic = true;          // Enable logarithmic scaling
                categoryAxis.LogBase = 10;                  // Base of the logarithm
                categoryAxis.MajorUnit = 2;                 // Tick interval (every 2 log units)
                categoryAxis.MinorUnit = 1;                 // Minor tick interval
                categoryAxis.Title.IsVisible = true;
                categoryAxis.Title.Text = "Date (Log Scale)";
                // Number format for dates (optional – remove if not supported by the used version)
                // categoryAxis.NumberFormat = "yyyy-MM-dd";

                // Configure the Y (value) axis (linear by default)
                Axis valueAxis = chart.ValueAxis;
                valueAxis.Title.IsVisible = true;
                valueAxis.Title.Text = "Sample Value";

                // Render the chart directly to a PNG file
                chart.ToImage("TimelineLogScale.png", ImageType.Png);

                // Optional: Save the workbook if you want to keep the Excel file as well
                // workbook.Save("TimelineLogScale.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }
}
