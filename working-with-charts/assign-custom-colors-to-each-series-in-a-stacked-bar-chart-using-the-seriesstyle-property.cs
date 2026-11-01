// Title: How to assign custom colors to each series in a stacked bar chart using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a stacked bar chart with three data series and sets the series colors to red, green, and blue using Aspose.Cells Series.Area.ForegroundColor. | Show how to iterate over chart.NSeries in Aspose.Cells and apply a solid fill color to each series in a stacked bar chart.
// Common Searches: asp.net set individual series colors in a stacked bar chart with Aspose.Cells | c# Aspose.Cells change series fill color for stacked bar chart | how to use Series.Area.ForegroundColor to color chart series in Aspose.Cells | custom color palette for stacked bar chart series using Aspose.Cells .NET
// Tags: Aspose.Cells set series foreground color .NET | stacked bar chart custom series colors Aspose.Cells | Series.Area solid fill color Aspose.Cells C# | apply custom palette to chart series Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsStackedBarChart
{
    // The example creates a workbook, fills it with sample data, adds a stacked bar chart, and uses the Series.Area.ForegroundColor property in a loop to assign red, green, and blue colors to the three series before saving the file as StackedBarCustomColors.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook.
                Workbook workbook = new Workbook();

                // Access the first worksheet.
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the stacked bar chart.
                // Columns: Category, Series1, Series2, Series3
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Series 1");
                sheet.Cells["C1"].PutValue("Series 2");
                sheet.Cells["D1"].PutValue("Series 3");

                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");

                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                sheet.Cells["C2"].PutValue(15);
                sheet.Cells["C3"].PutValue(25);
                sheet.Cells["C4"].PutValue(35);

                sheet.Cells["D2"].PutValue(20);
                sheet.Cells["D3"].PutValue(30);
                sheet.Cells["D4"].PutValue(40);

                // Add a stacked bar chart.
                int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 6, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart (including categories and series).
                chart.NSeries.Add("B2:D4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Assign custom colors to each series.
                Color[] seriesColors = { Color.Red, Color.Green, Color.Blue };
                for (int i = 0; i < chart.NSeries.Count; i++)
                {
                    Series series = chart.NSeries[i];
                    Color col = seriesColors[i % seriesColors.Length];

                    // Apply the color to the entire series.
                    series.Area.ForegroundColor = col;
                    // Setting a solid fill is the default behavior; no explicit pattern property needed.
                }

                // Optional: Set a title for clarity.
                chart.Title.Text = "Custom Colored Stacked Bar Chart";

                // Save the workbook to a file.
                workbook.Save("StackedBarCustomColors.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
