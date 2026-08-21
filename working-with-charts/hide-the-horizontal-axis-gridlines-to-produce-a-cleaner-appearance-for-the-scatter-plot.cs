// Title: Hide Horizontal Gridlines in an Aspose.Cells Scatter Chart – C# Example
// Description: Demonstrates how to create a workbook, add X/Y data, insert a scatter chart, and turn off both major and minor horizontal (value‑axis) gridlines using Aspose.Cells for .NET, resulting in a cleaner chart saved as an XLSX file.
// Keywords: Aspose.Cells scatter chart C# | hide chart gridlines Aspose.Cells | disable horizontal gridlines .NET | value axis gridlines Aspose.Cells | chart formatting Aspose.Cells C# | Excel scatter plot gridlines
// Common Searches: how to hide horizontal gridlines in Aspose.Cells scatter chart C# | Aspose.Cells remove value axis gridlines .NET | disable major and minor gridlines Aspose.Cells chart | C# Aspose.Cells scatter plot without horizontal gridlines | chart formatting hide gridlines Aspose.Cells
// Developer Intent: Turn off the horizontal (value‑axis) gridlines of a scatter chart to produce a cleaner visual layout.
// Use Cases: Generate a minimalist scatter plot for data analysis reports. | Create presentation‑ready Excel charts that show only vertical gridlines. | Automate workbook creation where chart clutter is reduced for better readability.
// AI Prompts: Provide C# code using Aspose.Cells to create a scatter chart and hide the horizontal gridlines while keeping vertical gridlines visible. | Show how to toggle visibility of major and minor gridlines on the value axis of an Aspose.Cells chart in .NET. | Explain step‑by‑step how to customize gridline visibility for different axes in an Aspose.Cells scatter chart.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsScatterGridlinesDemo
{
    // Demonstrates how to create a workbook, add X/Y data, insert a scatter chart, and turn off both major and minor horizontal (value‑axis) gridlines using Aspose.Cells for .NET, resulting in a cleaner chart saved as an XLSX file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for a scatter chart (X values in column A, Y values in column B)
            sheet.Cells["A1"].PutValue("X");
            sheet.Cells["B1"].PutValue("Y");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue(2);
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue(4);
            sheet.Cells["A4"].PutValue(3);
            sheet.Cells["B4"].PutValue(6);
            sheet.Cells["A5"].PutValue(4);
            sheet.Cells["B5"].PutValue(8);

            // Add a scatter chart
            int chartIndex = sheet.Charts.Add(ChartType.Scatter, 7, 0, 25, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the scatter series (X values, Y values)
            chart.NSeries.Add("A2:A5", true);
            chart.NSeries[0].XValues = "A2:A5";
            chart.NSeries[0].Values = "B2:B5";

            // Hide horizontal gridlines (major and minor) for a cleaner appearance
            // Horizontal gridlines are drawn from the value axis (Y axis)
            chart.ValueAxis.MajorGridLines.IsVisible = false;
            chart.ValueAxis.MinorGridLines.IsVisible = false;

            // Optionally hide vertical gridlines as well
            // chart.CategoryAxis.MajorGridLines.IsVisible = false;
            // chart.CategoryAxis.MinorGridLines.IsVisible = false;

            // Save the workbook
            workbook.Save("ScatterChart_NoHorizontalGridlines.xlsx");
        }
    }
}
