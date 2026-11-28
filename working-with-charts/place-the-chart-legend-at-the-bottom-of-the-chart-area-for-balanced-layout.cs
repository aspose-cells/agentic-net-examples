// Title: Place the legend at the bottom of a column chart using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells move chart legend to bottom of column chart | Set bottom legend position for Excel chart with Aspose.Cells .NET | Adjust legend size and horizontal alignment in an Aspose.Cells chart | Disable automatic legend sizing and define width ratio in Aspose.Cells | Center a bottom legend in a column chart using Aspose.Cells C# example
// Tags: Aspose.Cells set chart legend position | C# column chart legend bottom alignment | Aspose.Cells disable automatic legend size | Excel chart legend width ratio Aspose.Cells | center bottom legend Aspose.Cells C#

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace ChartLegendBottomDemo
{
    // The sample creates a workbook, adds quarterly sales data, generates a column chart, assigns series and categories, sets a chart title, then moves the legend to the bottom of the chart area, disables automatic sizing, specifies width and height ratios, centers the legend horizontally, and saves the file as ChartWithBottomLegend.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook.
                Workbook workbook = new Workbook();

                // Access the first worksheet.
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart.
                // Category column.
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Q1");
                sheet.Cells["A3"].PutValue("Q2");
                sheet.Cells["A4"].PutValue("Q3");
                sheet.Cells["A5"].PutValue("Q4");

                // Series 1 values.
                sheet.Cells["B1"].PutValue("Product A");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(150);
                sheet.Cells["B4"].PutValue(180);
                sheet.Cells["B5"].PutValue(200);

                // Series 2 values.
                sheet.Cells["C1"].PutValue("Product B");
                sheet.Cells["C2"].PutValue(100);
                sheet.Cells["C3"].PutValue(130);
                sheet.Cells["C4"].PutValue(160);
                sheet.Cells["C5"].PutValue(190);

                // Add a column chart to the worksheet.
                // Parameters: chart type, upper-left row, upper-left column, lower-right row, lower-right column.
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data source for the chart.
                chart.NSeries.Add("B2:B5", true);
                chart.NSeries[0].Name = "Product A";

                chart.NSeries.Add("C2:C5", true);
                chart.NSeries[1].Name = "Product B";

                // Set category data for all series.
                chart.NSeries.CategoryData = "A2:A5";

                // Set a title for the chart (optional, just for completeness).
                chart.Title.Text = "Quarterly Sales Comparison";

                // ---- Legend positioning ----
                Legend legend = chart.Legend;

                // Place the legend at the bottom of the chart area.
                legend.Position = LegendPositionType.Bottom;

                // Disable automatic sizing and define explicit dimensions using ratio properties.
                legend.IsAutomaticSize = false;
                legend.WidthRatioToChart = 0.4;   // Approx. 40% of chart width
                legend.HeightRatioToChart = 0.03; // Approx. 3% of chart height

                // Center the legend horizontally within the chart area.
                legend.XRatioToChart = 0.5 - (legend.WidthRatioToChart / 2.0);

                // Save the workbook to a file.
                string outFile = "ChartWithBottomLegend.xlsx";
                workbook.Save(outFile, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved successfully to '{outFile}'.");
                Console.WriteLine("Legend is positioned at the bottom with a balanced layout.");
            }
            catch (Exception ex)
            {
                // Output any errors that occur during processing.
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
