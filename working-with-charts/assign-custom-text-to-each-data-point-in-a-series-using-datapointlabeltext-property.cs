// Title: Set custom text for each data point in an Aspose.Cells column chart with C#
// AI Prompts: Generate C# code that iterates over ChartPoint objects in an Aspose.Cells chart, disables IsAutoText, and assigns a formatted string to DataLabels.Text for each point. | Show how to enable data labels for a series and customize every point's label in a column chart using Aspose.Cells. | Provide a complete example that creates a workbook, adds a column chart, and applies custom labels to every data point in C#.
// Common Searches: Aspose.Cells C# set custom label for each point in a column chart | disable automatic data label text Aspose.Cells chart series | how to use ChartPoint.DataLabels.Text in Aspose.Cells C# example | custom data labels per data point Aspose.Cells column chart | C# Aspose.Cells chart point label formatting
// Tags: custom data labels per chart point Aspose.Cells | disable auto text for chart point labels .NET | set DataLabels.Text on ChartPoint C# | column chart custom labels Aspose.Cells | Aspose.Cells chart series point label customization

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsCustomDataLabelDemo
{
    // The example creates a workbook, adds sample data, inserts a column chart, enables data labels, then loops through each series point to turn off automatic label text and assign a custom string containing the point index and its Y value, finally saving the file as CustomDataLabels.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series and categories
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Enable data labels for each series
            foreach (Series series in chart.NSeries)
            {
                series.DataLabels.ShowValue = true; // show the default value

                // Assign custom text to each data point
                for (int i = 0; i < series.Points.Count; i++)
                {
                    ChartPoint point = series.Points[i];
                    // Disable auto-generated text and set custom label
                    point.DataLabels.IsAutoText = false;
                    point.DataLabels.Text = $"Point {i + 1}: {point.YValue}";
                }
            }

            // Save the workbook to an XLSX file
            workbook.Save("CustomDataLabels.xlsx");
        }
    }
}
