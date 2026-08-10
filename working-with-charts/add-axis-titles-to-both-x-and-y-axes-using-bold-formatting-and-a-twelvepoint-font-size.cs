// Title: Add bold 12‑point X and Y axis titles to an Aspose.Cells chart (C#)
// Description: Create a workbook, insert sample data, add a column chart, and set visible bold 12‑point titles for both the Category (X) and Value (Y) axes using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# chart axis title | bold axis title | 12 point font | CategoryAxis | ValueAxis | set axis title visibility | Aspose.Cells chart formatting | Excel chart axis styling | Aspose.Cells .NET
// Common Searches: Aspose.Cells set axis title bold | C# add X axis title to chart | How to make chart axis title 12 pt in Aspose.Cells | Aspose.Cells chart title formatting example | Add Y axis label in Aspose.Cells C#
// Developer Intent: Add visible, bold, 12‑point titles to both X (category) and Y (value) axes of a chart.
// Use Cases: Automated financial dashboards that require clear, bold axis labels. | Generating Excel reports with presentation‑ready charts that follow corporate typography. | Batch‑processing workbooks to enforce consistent axis styling across multiple charts.
// AI Prompts: Generate C# code with Aspose.Cells that creates a column chart and applies bold 12‑point titles to the CategoryAxis and ValueAxis. | Show how to enable axis titles, set their text, make them bold, and set font size to 12 points in Aspose.Cells for .NET. | Explain how to modify an existing Aspose.Cells chart to update axis title font style and size programmatically.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AxisTitleExample
{
    // Create a workbook, insert sample data, add a column chart, and set visible bold 12‑point titles for both the Category (X) and Value (Y) axes using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["A4"].PutValue("C");
            worksheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Configure X‑axis (CategoryAxis) title
            chart.CategoryAxis.Title.Text = "Categories";
            chart.CategoryAxis.Title.IsVisible = true;
            chart.CategoryAxis.Title.Font.IsBold = true;
            chart.CategoryAxis.Title.Font.Size = 12;

            // Configure Y‑axis (ValueAxis) title
            chart.ValueAxis.Title.Text = "Values";
            chart.ValueAxis.Title.IsVisible = true;
            chart.ValueAxis.Title.Font.IsBold = true;
            chart.ValueAxis.Title.Font.Size = 12;

            // Save the workbook to a file
            workbook.Save("AxisTitles_Output.xlsx");
        }
    }
}
