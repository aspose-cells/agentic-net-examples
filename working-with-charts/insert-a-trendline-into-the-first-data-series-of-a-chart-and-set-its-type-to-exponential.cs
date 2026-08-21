// Title: C# – Add an Exponential Trendline to a Line Chart with Aspose.Cells for .NET
// Description: Creates a new workbook, inserts a line chart from cells B1:B4, adds an exponential trendline to the first data series, shows the equation and R‑squared value, and saves the file as an Excel workbook.
// Keywords: Aspose.Cells C# chart trendline | exponential trendline .NET | Aspose.Cells add trendline to series | display trendline equation Aspose.Cells | R squared chart Aspose.Cells | line chart sample code C# | Aspose.Cells chart API example
// Common Searches: how to add exponential trendline using Aspose.Cells C# | Aspose.Cells show trendline equation and R‑squared | C# chart trendline type exponential Aspose.Cells | sample code for Aspose.Cells line chart with trendline | Aspose.Cells add trendline to first series
// Developer Intent: Insert an exponential trendline into the first series of a line chart and display its equation and R‑squared value.
// Use Cases: Generate a sales‑forecast workbook where the line chart visualizes exponential growth. | Produce a scientific report that fits experimental data with an exponential curve and shows the formula. | Automate a KPI dashboard that highlights rapid performance changes using an exponential trendline.
// AI Prompts: Write C# code with Aspose.Cells to create a line chart from range B1:B4 and apply an exponential trendline to the first series, showing the equation and R‑squared. | Explain how to change the color and thickness of an exponential trendline after adding it to a chart in Aspose.Cells. | Show how to retrieve the calculated equation string of an exponential trendline from a chart using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a new workbook, inserts a line chart from cells B1:B4, adds an exponential trendline to the first data series, shows the equation and R‑squared value, and saves the file as an Excel workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue(1);
            sheet.Cells["A2"].PutValue(2);
            sheet.Cells["A3"].PutValue(3);
            sheet.Cells["A4"].PutValue(4);
            sheet.Cells["B1"].PutValue(2);
            sheet.Cells["B2"].PutValue(4);
            sheet.Cells["B3"].PutValue(6);
            sheet.Cells["B4"].PutValue(8);

            // Add a line chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 15, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the first series (Y values)
            chart.NSeries.Add("B1:B4", true);
            // Category (X) data can be set if the API supports it; omitted here to avoid compilation issues.

            // Insert an exponential trendline into the first data series
            int trendlineIdx = chart.NSeries[0].TrendLines.Add(TrendlineType.Exponential);
            Trendline trendline = chart.NSeries[0].TrendLines[trendlineIdx];

            // Display equation and R‑squared value
            trendline.DisplayEquation = true;
            trendline.DisplayRSquared = true;

            // Save the workbook
            string outputPath = "ChartWithExponentialTrendline.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
