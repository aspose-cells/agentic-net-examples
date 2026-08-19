// Title: C# Example: Add Standard Deviation Y‑Error Bars to a Line Chart with Aspose.Cells
// Description: Learn how to create a workbook, populate categories and values, insert a line chart, and configure Y‑error bars to show standard deviation (both plus and minus) with custom color and weight using Aspose.Cells for .NET. The sample saves the result as LineChartWithStdDevErrorBar.xlsx.
// Keywords: Aspose.Cells C# error bars | standard deviation error bar line chart | YErrorBar StDev Aspose.Cells | customize chart error bars .NET | Aspose.Cells line chart example | Excel chart error bars C# | GitHub Aspose.Cells samples
// Common Searches: Aspose.Cells add Y error bar standard deviation | C# line chart error bars Aspose | How to set error bar type StDev in Aspose.Cells | Customize error bar color and weight Aspose.Cells | Example of error bars in Aspose.Cells chart
// Developer Intent: The developer needs a ready‑to‑run C# snippet that adds a line chart to an Excel workbook and displays standard‑deviation Y‑error bars (both positive and negative) with optional styling using Aspose.Cells.
// Use Cases: Show quarterly sales trends with variability by adding StDev error bars to each point. | Visualize scientific measurement uncertainty on a line plot, highlighting confidence intervals. | Create a financial performance dashboard where error‑bar styling matches corporate branding.
// AI Prompts: Generate C# code that creates a line chart with standard deviation Y‑error bars and custom styling using Aspose.Cells. | Explain how to compute custom error‑bar values and apply them to an Aspose.Cells chart series in .NET. | Provide step‑by‑step instructions to modify an existing Aspose.Cells chart to display both plus and minus error bars and change their color to red.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsErrorBarExample
{
    // Learn how to create a workbook, populate categories and values, insert a line chart, and configure Y‑error bars to show standard deviation (both plus and minus) with custom color and weight using Aspose.Cells for .NET. The sample saves the result as LineChartWithStdDevErrorBar.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for the line chart
            // Column A: X values (categories)
            // Column B: Y values (data series)
            cells["A1"].PutValue("Category");
            cells["A2"].PutValue("Q1");
            cells["A3"].PutValue("Q2");
            cells["A4"].PutValue("Q3");
            cells["A5"].PutValue("Q4");

            cells["B1"].PutValue("Series 1");
            cells["B2"].PutValue(10);
            cells["B3"].PutValue(15);
            cells["B4"].PutValue(20);
            cells["B5"].PutValue(25);

            // Add a line chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Line, 7, 0, 25, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series and categories
            chart.NSeries.Add("B2:B5", true);          // Y values
            chart.NSeries.CategoryData = "A2:A5";      // X categories

            // Configure the Y error bar to display standard deviation
            Series series = chart.NSeries[0];
            series.YErrorBar.Type = ErrorBarType.StDev;               // Use standard deviation
            series.YErrorBar.DisplayType = ErrorBarDisplayType.Both; // Show both plus and minus bars

            // Optionally customize appearance of the error bars
            series.YErrorBar.Color = System.Drawing.Color.Blue;
            series.YErrorBar.Weight = WeightType.SingleLine;
            series.YErrorBar.IsVisible = true;

            // Save the workbook to an XLSX file
            workbook.Save("LineChartWithStdDevErrorBar.xlsx", SaveFormat.Xlsx);
        }
    }
}
