// Title: C# – Assign Custom Data Labels to the First Point of Each Series in an Aspose.Cells Chart
// Description: This example creates a workbook, adds sample categories and two data series, builds a column chart, enables data labels, and then loops through each series to disable automatic text for the first point and set a custom label such as "Series1 - First". The chart is recalculated and saved as **CustomFirstPointLabels.xlsx** using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# chart data labels | custom label first point | Aspose.Cells series loop | .NET chart customization | ChartPoint DataLabels | set IsAutoText false | column chart example | Excel automation Aspose
// Common Searches: Aspose.Cells set custom label for first chart point | C# change data label text of a specific series point | How to disable IsAutoText in Aspose.Cells chart | Assign custom text to first data point in Aspose.Cells | Loop through series to modify chart labels .NET
// Developer Intent: Programmatically replace the default label of the first data point in every chart series with a custom text string.
// Use Cases: Highlight the opening value of each product line in a sales column chart. | Mark the initial month of a time‑series with a descriptive tag instead of a numeric value. | Create a dashboard where the first point of each series shows a custom identifier for quick reference.
// AI Prompts: Show C# code that iterates over chart series in Aspose.Cells and sets a custom label for the first point. | How to turn off automatic data label text and assign a specific string to a ChartPoint in Aspose.Cells? | Explain the steps to recalculate an Aspose.Cells chart after modifying point labels so the changes appear in the saved file.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsCustomLabelDemo
{
    // This example creates a workbook, adds sample categories and two data series, builds a column chart, enables data labels, and then loops through each series to disable automatic text for the first point and set a custom label such as "Series1 - First". The chart is recalculated and saved as **CustomFirstPointLabels.xlsx** using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (categories + two series)
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");

            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            sheet.Cells["C1"].PutValue("Series2");
            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(35);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Add the two series and set the category (X) data
            chart.NSeries.Add("B2:B4", true); // Series 1
            chart.NSeries.Add("C2:C4", true); // Series 2
            chart.NSeries.CategoryData = "A2:A4";

            // Enable data labels for all series (show the value by default)
            foreach (Series series in chart.NSeries)
            {
                series.DataLabels.ShowValue = true;
            }

            // Loop through each series and assign a custom label to its first point
            int seriesIdx = 0;
            foreach (Series series in chart.NSeries)
            {
                // Access the first point (index 0) of the current series
                ChartPoint firstPoint = series.Points[0];

                // Disable automatic text generation and set a custom label
                firstPoint.DataLabels.IsAutoText = false;
                firstPoint.DataLabels.Text = $"Series{seriesIdx + 1} - First";

                seriesIdx++;
            }

            // Optional: recalculate the chart to apply changes
            chart.Calculate();

            // Save the workbook
            workbook.Save("CustomFirstPointLabels.xlsx");
        }
    }
}
