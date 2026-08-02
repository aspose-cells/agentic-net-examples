// Title: Customize Line Series Markers and Highlight Top Profit Point in a Combo Chart – Aspose.Cells for .NET
// Description: Creates an Excel workbook, adds a column‑line combo chart, applies diamond markers with red outline and yellow fill to the entire profit line series, detects the maximum profit value, and marks that point with a larger blue circle on a white background before saving the file.
// Keywords: Aspose.Cells .NET | combo chart marker style | line series custom marker | highlight max data point | ChartMarkerType C# | Excel chart customization | marker size color shape | Aspose.Cells example
// Common Searches: Aspose.Cells change line series marker shape | C# set custom marker colors in combo chart | highlight specific point in Aspose.Cells chart | how to use ChartMarkerType in Aspose.Cells | customize marker size for Excel line series .NET
// Developer Intent: Apply custom marker shapes, sizes, and colors to a line series in a combo chart and emphasize a particular data point such as the highest profit.
// Use Cases: Display sales as columns and profit as a line with distinctive diamond markers. | Identify the month with the highest profit and draw attention to it using a larger, differently colored marker. | Generate an Excel report where key performance indicators are visually highlighted through marker styling.
// AI Prompts: Write C# code with Aspose.Cells that builds a column‑line combo chart and sets diamond markers (red outline, yellow fill) for the line series. | Provide a function that finds the maximum value in an array and applies a blue circle marker to that point in an Aspose.Cells line series. | Explain how to modify marker style, size, and colors for an entire series and for individual chart points using Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsDemo
{
    // Creates an Excel workbook, adds a column‑line combo chart, applies diamond markers with red outline and yellow fill to the entire profit line series, detects the maximum profit value, and marks that point with a larger blue circle on a white background before saving the file.
    class ComboChartMarkerDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for a combo chart (column + line)
                sheet.Cells["A1"].PutValue("Month");
                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["C1"].PutValue("Profit");

                string[] months = { "Jan", "Feb", "Mar", "Apr", "May" };
                double[] sales = { 5000, 7000, 6000, 8000, 7500 };
                double[] profit = { 1500, 2000, 1800, 2200, 2100 };

                for (int i = 0; i < months.Length; i++)
                {
                    sheet.Cells[i + 1, 0].PutValue(months[i]);   // Column A
                    sheet.Cells[i + 1, 1].PutValue(sales[i]);   // Column B
                    sheet.Cells[i + 1, 2].PutValue(profit[i]);  // Column C
                }

                // Add a combo chart: column series for Sales, line series for Profit
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
                Chart chart = sheet.Charts[chartIndex];
                chart.Title.Text = "Sales and Profit";

                // Column series (Sales)
                chart.NSeries.Add("B2:B6", true);
                chart.NSeries[0].Name = "Sales";

                // Line series (Profit)
                chart.NSeries.Add("C2:C6", true);
                chart.NSeries[1].Name = "Profit";
                chart.NSeries[1].Type = ChartType.Line;
                // Note: Setting series on secondary axis is not supported in this version; omitted.

                // Customize marker style for the entire line series
                Series lineSeries = chart.NSeries[1];
                lineSeries.Marker.MarkerStyle = ChartMarkerType.Diamond;
                lineSeries.Marker.MarkerSize = 12;
                lineSeries.Marker.ForegroundColor = Color.Red;
                lineSeries.Marker.BackgroundColor = Color.Yellow;

                // Highlight the data point with the highest profit using a distinct marker
                int maxIndex = 0;
                double maxProfit = profit[0];
                for (int i = 1; i < profit.Length; i++)
                {
                    if (profit[i] > maxProfit)
                    {
                        maxProfit = profit[i];
                        maxIndex = i;
                    }
                }

                ChartPoint maxPoint = lineSeries.Points[maxIndex];
                maxPoint.Marker.MarkerStyle = ChartMarkerType.Circle;
                maxPoint.Marker.MarkerSize = 16;
                maxPoint.Marker.ForegroundColor = Color.Blue;
                maxPoint.Marker.BackgroundColor = Color.White;

                // Save the workbook with the customized combo chart
                string outputPath = "ComboChartWithCustomMarkers.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            ComboChartMarkerDemo.Run();
        }
    }
}
