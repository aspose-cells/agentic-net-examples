// Title: Aspose.Cells C# – Disable Chart Data Label Text Wrapping (DataLabels.IsTextWrapped = false)
// Description: Creates a workbook, adds a column chart with quarterly sales, enables data labels, sets DataLabels.IsTextWrapped to false so each label stays on a single line, and saves the file as ChartDataLabels_NoWrap.xlsx.
// Keywords: Aspose.Cells chart data label wrap | DataLabels.IsTextWrapped false | C# Excel chart label formatting | disable text wrap Aspose.Cells | .NET chart label property | Excel column chart label settings
// Common Searches: Aspose.Cells disable data label wrap C# | DataLabels.IsTextWrapped property example | turn off text wrapping for chart labels Aspose | C# chart data label no wrap Aspose.Cells
// Developer Intent: Prevent chart data label text from wrapping onto multiple lines.
// Use Cases: Generate a sales column chart where each label shows the value on one line for a compact layout. | Produce Excel reports with charts that require non‑wrapped labels to maintain visual consistency. | Export charts with long numeric values without label overflow or multiline rendering.
// AI Prompts: Show how to set DataLabels.IsTextWrapped = false for a chart in Aspose.Cells using C#. | Provide a complete Aspose.Cells C# example that creates a chart and disables data label text wrapping. | Explain other ways to control chart data label appearance in Aspose.Cells, such as adjusting label size or disabling AutoFit.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartDataLabelWrapDemo
{
    // Creates a workbook, adds a column chart with quarterly sales, enables data labels, sets DataLabels.IsTextWrapped to false so each label stays on a single line, and saves the file as ChartDataLabels_NoWrap.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["B2"].PutValue(12000);
            sheet.Cells["B3"].PutValue(15000);
            sheet.Cells["B4"].PutValue(18000);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series and categories
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Enable data labels for the first series
            DataLabels dataLabels = chart.NSeries[0].DataLabels;
            dataLabels.ShowValue = true;

            // Disable text wrapping for the data labels
            dataLabels.IsTextWrapped = false;

            // Save the workbook to a file
            workbook.Save("ChartDataLabels_NoWrap.xlsx");
        }
    }
}
