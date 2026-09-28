// Title: How to place data labels inside the end of columns in an Aspose.Cells column chart using C#
// AI Prompts: Write a C# program with Aspose.Cells that creates a 2‑D column chart and configures the series data labels to appear at the InsideEnd location. | Modify an existing Aspose.Cells workbook so that the column chart series displays its data labels inside the top of each bar to prevent overlap.
// Common Searches: Aspose.Cells C# move column chart data labels inside the bars | prevent overlapping labels in Aspose.Cells column chart | example of InsideEnd label position with Aspose.Cells chart series | C# Aspose.Cells set data label placement for column chart | how to adjust label position in Aspose.Cells column chart
// Tags: Aspose.Cells column chart label placement | C# InsideEnd data label position Aspose.Cells | chart series label positioning Aspose.Cells | avoid overlapping chart labels .NET | Aspose.Cells label position type configuration

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsLabelPositionExample
{
    // Demonstrates creating a workbook, adding a 2‑D column chart, enabling data labels for the first series, setting their position to InsideEnd so labels stay within the bars, and saving the result as ColumnChart_InsideEndLabels.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the column chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a 2‑D column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series and the category axis
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Enable data labels for the first series
            Series series = chart.NSeries[0];
            series.DataLabels.ShowValue = true;

            // Position data labels inside the end of each column to keep them within the bar area
            series.DataLabels.Position = LabelPositionType.InsideEnd;

            // Optional: adjust other label properties if needed
            // series.DataLabels.ShowCategoryName = true;
            // series.DataLabels.IsAutoText = true;

            // Save the workbook to a file
            workbook.Save("ColumnChart_InsideEndLabels.xlsx");
        }
    }
}
