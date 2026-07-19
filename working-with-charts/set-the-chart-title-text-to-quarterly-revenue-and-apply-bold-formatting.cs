// Title: Bold "Quarterly Revenue" Chart Title with Aspose.Cells for .NET (C#)
// Description: Creates an Excel workbook, adds quarterly revenue data, inserts a column chart, sets the visible title to "Quarterly Revenue", applies bold formatting (with optional size and color), and saves the file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | set chart title | bold chart title | chart title formatting | Excel chart title | column chart Aspose.Cells | font styling Excel | Aspose.Cells example | Excel automation
// Common Searches: Aspose.Cells set chart title C# | Bold chart title Aspose.Cells | How to format chart title font Aspose.Cells | Change chart title color Aspose.Cells | Create column chart with title Aspose.Cells
// Developer Intent: Set the chart title text to "Quarterly Revenue" and apply bold styling (including optional font size and color) in a C# workbook generated with Aspose.Cells.
// Use Cases: Produce a quarterly revenue report where the chart title stands out for better readability. | Build an Excel dashboard that highlights key metrics with a bold chart title. | Automate generation of presentation‑ready Excel files with styled chart titles.
// AI Prompts: Generate C# code using Aspose.Cells that adds a column chart and makes the title "Quarterly Revenue" bold. | Explain how to change font size, color, and boldness of a chart title in Aspose.Cells for .NET. | Show the steps to create a visible, bold chart title after populating data in an Aspose.Cells workbook.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates an Excel workbook, adds quarterly revenue data, inserts a column chart, sets the visible title to "Quarterly Revenue", applies bold formatting (with optional size and color), and saves the file using Aspose.Cells for .NET.
class SetChartTitleBold
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add some sample data for the chart (required for a valid chart)
        sheet.Cells["A1"].PutValue("Quarter");
        sheet.Cells["A2"].PutValue("Q1");
        sheet.Cells["A3"].PutValue("Q2");
        sheet.Cells["A4"].PutValue("Q3");
        sheet.Cells["A5"].PutValue("Q4");

        sheet.Cells["B1"].PutValue("Revenue");
        sheet.Cells["B2"].PutValue(15000);
        sheet.Cells["B3"].PutValue(20000);
        sheet.Cells["B4"].PutValue(18000);
        sheet.Cells["B5"].PutValue(22000);

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.NSeries.Add("B2:B5", true);
        chart.NSeries.CategoryData = "A2:A5";

        // Set the chart title text
        chart.Title.Text = "Quarterly Revenue";
        chart.Title.IsVisible = true;

        // Apply bold formatting to the title
        chart.Title.Font.IsBold = true;
        // Optional: set a larger font size for better visibility
        chart.Title.Font.Size = 14;
        chart.Title.Font.Color = Color.Black;

        // Save the workbook
        workbook.Save("QuarterlyRevenueChart.xlsx");
    }
}
