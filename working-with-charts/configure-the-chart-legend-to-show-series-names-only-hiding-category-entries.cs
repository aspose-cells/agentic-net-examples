// Title: Configure an Excel column chart legend to display only series names with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# hide category labels in chart legend | How to show only series names in Excel chart legend using Aspose.Cells | Set chart legend to series-only and position bottom in .NET workbook | Aspose.Cells chart legend customization to display series titles only | C# create column chart with legend showing only product series in Aspose.Cells
// Tags: Aspose.Cells chart legend series-only setting | C# set chart legend position bottom | Aspose.Cells NSeries legend entry customization | Excel column chart legend configuration with Aspose.Cells | Aspose.Cells workbook chart generation C#

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsLegendDemo
{
    // The example creates a new workbook, adds quarterly sales data for two products, builds a column chart, assigns category data, configures the legend to appear at the bottom showing only the series names, adds a chart title, and saves the file as ChartLegendSeriesOnly.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                // Categories (X‑axis)
                sheet.Cells["A2"].PutValue("Q1");
                sheet.Cells["A3"].PutValue("Q2");
                sheet.Cells["A4"].PutValue("Q3");
                sheet.Cells["A5"].PutValue("Q4");

                // Series 1
                sheet.Cells["B1"].PutValue("Product A");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(150);
                sheet.Cells["B4"].PutValue(130);
                sheet.Cells["B5"].PutValue(170);

                // Series 2
                sheet.Cells["C1"].PutValue("Product B");
                sheet.Cells["C2"].PutValue(80);
                sheet.Cells["C3"].PutValue(95);
                sheet.Cells["C4"].PutValue(100);
                sheet.Cells["C5"].PutValue(110);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Add the series data
                chart.NSeries.Add("B2:B5", true);
                chart.NSeries.Add("C2:C5", true);

                // Set the category (X‑axis) data
                chart.NSeries.CategoryData = "A2:A5";

                // Configure the legend to show only series names (default behavior)
                chart.Legend.Position = LegendPositionType.Bottom;

                // Give the chart a title
                chart.Title.Text = "Quarterly Sales Comparison";

                // Save the workbook
                string outputPath = "ChartLegendSeriesOnly.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                Console.WriteLine("The chart legend displays the series names (Product A, Product B).");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the chart:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
