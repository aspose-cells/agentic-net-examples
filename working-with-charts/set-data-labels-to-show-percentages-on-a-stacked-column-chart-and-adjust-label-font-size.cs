// Title: Add percentage data labels with custom font size to a stacked column chart using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that creates a stacked column chart, hides the raw values, shows each segment's percentage, and sets the data‑label font to 12 pt black. | Update an existing Aspose.Cells workbook so that a stacked column chart displays percentage labels instead of values and changes the label font size and color.
// Common Searches: Aspose.Cells C# how to display percentages on stacked column chart data labels | set data label font size in Aspose.Cells chart using C# | hide raw values and show percentages in Excel stacked column chart with Aspose.Cells | change data label color and size for stacked column chart in Aspose.Cells .NET | example code for percentage labels on stacked column chart Aspose.Cells
// Tags: Aspose.Cells stacked column chart percentage labels | C# set chart data label font size Aspose.Cells | Aspose.Cells hide raw values show percentages | customize Excel chart data label appearance Aspose.Cells | Aspose.Cells chart title and label styling

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The program creates a workbook, fills it with sample data, adds a stacked column chart, configures each series to hide raw values, show percentage labels with a 12‑point black font, sets a chart title, and saves the file as StackedColumnChart.xlsx.
class StackedColumnChartWithPercentageLabels
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and give it a name
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Populate sample data for a stacked column chart
            // Columns: Category, Series1, Series2, Series3
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["C1"].PutValue("Series2");
            sheet.Cells["D1"].PutValue("Series3");

            string[] categories = { "Q1", "Q2", "Q3", "Q4" };
            double[,] values = {
                { 30, 20, 10 },
                { 40, 25, 15 },
                { 35, 30, 20 },
                { 45, 35, 25 }
            };

            for (int i = 0; i < categories.Length; i++)
            {
                sheet.Cells[i + 2, 0].PutValue(categories[i]);               // Category names in column A
                sheet.Cells[i + 2, 1].PutValue(values[i, 0]);               // Series1 values in column B
                sheet.Cells[i + 2, 2].PutValue(values[i, 1]);               // Series2 values in column C
                sheet.Cells[i + 2, 3].PutValue(values[i, 2]);               // Series3 values in column D
            }

            // Add a stacked column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.ColumnStacked, 6, 0, 26, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart (including categories and all series)
            chart.NSeries.Add("B2:D5", true);
            chart.NSeries.CategoryData = "A2:A5";

            // Enable data labels and configure them to show percentages
            foreach (Series series in chart.NSeries)
            {
                series.DataLabels.ShowValue = false;               // Hide raw values
                series.DataLabels.ShowPercentage = true;           // Show percentage of each stacked segment
                series.DataLabels.Font.Size = 12;                  // Set desired font size
                series.DataLabels.Font.Color = Color.Black;
            }

            // Optional: add a chart title
            chart.Title.Text = "Stacked Column Chart – Percentages";

            // Save the workbook to a file
            workbook.Save("StackedColumnChart.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
