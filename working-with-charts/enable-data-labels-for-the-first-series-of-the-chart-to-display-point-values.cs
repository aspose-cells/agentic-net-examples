// Title: Show Point Values with Data Labels on the First Series of an Aspose.Cells Column Chart (C#)
// Description: Creates a workbook, inserts sample data, adds a column chart, defines the first series, enables DataLabels.ShowValue to display each point's value, recalculates the chart, and saves the file as an Excel workbook.
// Keywords: Aspose.Cells C# chart data labels | enable data labels Aspose.Cells | show point values column chart | Series.DataLabels.ShowValue | Aspose.Cells chart calculation | Excel chart automation .NET
// Common Searches: Aspose.Cells show data labels on chart series C# | How to display point values in a column chart using Aspose.Cells | Enable data labels for first series Aspose.Cells .NET | Aspose.Cells chart label visibility example
// Developer Intent: Add data labels to the first series of a column chart so each column displays its numeric value.
// Use Cases: Financial statements where each column’s amount must be visible without hovering. | Executive dashboards that require immediate insight into individual data points. | Automated report generation that embeds value labels directly on charts for print‑ready Excel files.
// AI Prompts: Generate C# code with Aspose.Cells to add data labels to all series of a line chart and customize their font. | Write a method that toggles DataLabels.ShowValue for a given series index in an Aspose.Cells chart. | Explain how to position data labels (inside, outside, center) and change their number format after enabling ShowValue.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsDataLabelsDemo
{
    // Creates a workbook, inserts sample data, adds a column chart, defines the first series, enables DataLabels.ShowValue to display each point's value, recalculates the chart, and saves the file as an Excel workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");

            sheet.Cells["B1"].PutValue("Series 1");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Add the first series (vertical data range) and set category data
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Enable data labels for the first series to show point values
            Series firstSeries = chart.NSeries[0];
            firstSeries.DataLabels.ShowValue = true;

            // Optional: calculate the chart to ensure labels are rendered correctly
            chart.Calculate();

            // Save the workbook
            workbook.Save("ChartWithDataLabels.xlsx");
        }
    }
}
