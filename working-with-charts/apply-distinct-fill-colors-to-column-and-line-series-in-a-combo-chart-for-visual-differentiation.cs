// Title: Set different fill colors for column and line series in an Aspose.Cells combo chart using C#
// AI Prompts: Generate C# code with Aspose.Cells that creates a column‑line combo chart and applies a custom fill color to the column series. | Show how to assign a border color to the line series of an Aspose.Cells combo chart and place the line on a secondary axis. | Provide an example of configuring series types and distinct colors for a mixed chart in Aspose.Cells for .NET.
// Common Searches: how to set a custom fill color for a column series in an Aspose.Cells combo chart (C#) | changing the line series border color in a mixed chart with Aspose.Cells | placing a line series on a secondary axis in an Aspose.Cells combo chart using C# | example of column and line series with distinct colors in Aspose.Cells | Aspose.Cells C# tutorial for styling series in a combo chart
// Tags: Aspose.Cells column series fill | Aspose.Cells line series border | Aspose.Cells combo chart secondary axis | Aspose.Cells mixed chart styling | Aspose.Cells C# series formatting

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsComboChartExample
{
    // The sample creates a workbook, adds month, sales, and profit data, builds a column‑line combo chart, sets a custom fill color for the column (sales) series, applies a border color to the line (profit) series, plots the line on a secondary axis, and saves the file as ComboChartWithDistinctColors.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Month");
                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["C1"].PutValue("Profit");

                string[] months = { "Jan", "Feb", "Mar", "Apr", "May" };
                double[] sales = { 12000, 15000, 13000, 17000, 16000 };
                double[] profit = { 3000, 4000, 3500, 4500, 4200 };

                for (int i = 0; i < months.Length; i++)
                {
                    sheet.Cells[i + 1, 0].PutValue(months[i]);   // Column A: Month
                    sheet.Cells[i + 1, 1].PutValue(sales[i]);   // Column B: Sales (Column series)
                    sheet.Cells[i + 1, 2].PutValue(profit[i]);  // Column C: Profit (Line series)
                }

                // Add a combo chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 25, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set chart title
                chart.Title.Text = "Monthly Sales and Profit";

                // Add the column series (Sales)
                int colSeriesIndex = chart.NSeries.Add("B2:B6", true);
                Series colSeries = chart.NSeries[colSeriesIndex];
                colSeries.Name = "Sales";
                colSeries.Type = ChartType.Column; // Ensure column display
                colSeries.Area.ForegroundColor = Color.CornflowerBlue; // Column fill color

                // Add the line series (Profit)
                int lineSeriesIndex = chart.NSeries.Add("C2:C6", true);
                Series lineSeries = chart.NSeries[lineSeriesIndex];
                lineSeries.Name = "Profit";
                lineSeries.Type = ChartType.Line; // Set series type to Line
                // Use border color to define line color
                lineSeries.Border.Color = Color.OrangeRed;

                // Optional: Plot the line series on a secondary axis if supported
                // This property exists in recent versions; if unavailable, the line will plot on the primary axis.
                lineSeries.PlotOnSecondAxis = true;

                // Save the workbook
                string outputPath = "ComboChartWithDistinctColors.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
