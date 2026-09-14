// Title: How to apply a logarithmic scale to the primary Y‑axis of a column chart using Aspose.Cells for .NET
// AI Prompts: Generate a column chart in a new workbook and enable a logarithmic Y axis with base 10 via the Aspose.Cells .NET API. | Set custom minimum and maximum values for the logarithmic Y axis and assign a descriptive title using Aspose.Cells chart properties.
// Common Searches: Aspose.Cells .NET set primary Y axis to log scale for column chart | C# example of configuring logarithmic axis base and limits in Aspose.Cells | how to add a title to a logarithmic value axis in Aspose.Cells chart | create column chart with wide value range using logarithmic Y axis Aspose.Cells
// Tags: logarithmic Y axis Aspose.Cells .NET | column chart logarithmic scaling Aspose.Cells | set chart axis min max Aspose.Cells | configure value axis base Aspose.Cells | add axis title Aspose.Cells chart

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Demonstrates creating a workbook, inserting data, adding a column chart, and configuring the primary Y axis to use a base‑10 logarithmic scale with custom min/max limits and a title, then saving the file as LogarithmicYAxisDemo.xlsx.
    public class LogarithmicYAxisDemo
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data with a wide range of values
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Low");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("Medium");
            sheet.Cells["B3"].PutValue(1000);
            sheet.Cells["A4"].PutValue("High");
            sheet.Cells["B4"].PutValue(1000000);

            // Insert a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the chart data source
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Apply logarithmic scaling to the primary Y axis
            chart.ValueAxis.IsLogarithmic = true;   // Enable logarithmic scale
            chart.ValueAxis.LogBase = 10;           // Set logarithmic base (default is 10)
            chart.ValueAxis.MinValue = 1;           // Minimum value for the log axis
            chart.ValueAxis.MaxValue = 10000000;    // Maximum value for the log axis

            // Optional: give the axis a descriptive title
            chart.ValueAxis.Title.Text = "Value (Log Scale)";

            // Save the workbook
            workbook.Save("LogarithmicYAxisDemo.xlsx");
        }
    }
}
