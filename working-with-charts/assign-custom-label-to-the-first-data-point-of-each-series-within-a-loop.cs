// Title: How to set a custom label for the first data point of each series in an Aspose.Cells column chart (C#)
// AI Prompts: Write C# code that iterates over all series in an Aspose.Cells column chart and assigns a custom text label to the first point of each series. | Show how to turn off the auto‑generated label for a specific chart point and replace it with a custom string using Aspose.Cells. | Provide a complete example that creates a workbook, adds a column chart, enables data labels, and customizes the first point label for each series in C#.
// Common Searches: Aspose.Cells C# set custom label for first point in column chart | loop through chart series to change data label text Aspose.Cells | disable auto text for a chart point Aspose.Cells C# example | assign series name to first data point label Aspose.Cells | how to customize data labels per point in Aspose.Cells chart C#
// Tags: custom data label first chart point Aspose.Cells | iterate chart series Aspose.Cells C# | disable auto text chart label Aspose.Cells | column chart series point label C# | assign series name to data label Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsCustomLabelDemo
{
    // Demonstrates creating a workbook, adding a column chart with two series, enabling data labels, looping through each series, and assigning a custom text label to the first data point while disabling the auto‑generated label, then saving the file as an .xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for two series
            // Column A – Category (X axis)
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");

            // Column B – Series 1 values
            sheet.Cells["B1"].PutValue("Series 1");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Column C – Series 2 values
            sheet.Cells["C1"].PutValue("Series 2");
            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(35);

            // Add a column chart
            int chartIdx = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 12);
            Chart chart = sheet.Charts[chartIdx];

            // Add the two series to the chart (vertical = true means each column is a series)
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.Add("C2:C4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Loop through each series in the chart
            foreach (Series series in chart.NSeries)
            {
                // Ensure data labels are visible for the series
                series.DataLabels.ShowValue = true;

                // Access the first data point (index 0) of the current series
                ChartPoint firstPoint = series.Points[0];

                // Assign a custom label text to the first point
                // Example: "First of {SeriesName}"
                string seriesName = series.Name; // May be empty if not set; you can set it earlier if needed
                firstPoint.DataLabels.Text = $"First of {seriesName}";
                // Optionally, you can also hide the auto‑generated text
                firstPoint.DataLabels.IsAutoText = false;
            }

            // Recalculate the chart to apply changes
            chart.Calculate();

            // Save the workbook
            workbook.Save("CustomFirstPointLabels.xlsx");
        }
    }
}
