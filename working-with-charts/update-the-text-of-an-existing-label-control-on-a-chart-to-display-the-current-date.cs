// Title: Add or update a chart label with today’s date using Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, insert a column chart, add a label shape inside the chart area, set its text to the current date (yyyy‑MM‑dd) via C# and Aspose.Cells, and save the file as an Excel workbook.
// Keywords: Aspose.Cells | .NET | C# | chart label | AddLabelInChart | current date | Excel chart annotation | update chart label | Aspose.Cells chart shape | DateTime.Now
// Common Searches: Aspose.Cells add label to chart C# | Set chart label text to today’s date Aspose.Cells | Update existing chart label Aspose.Cells .NET | Change chart annotation dynamically C# | How to use AddLabelInChart with Aspose.Cells
// Developer Intent: Insert or modify a chart label to show the current date programmatically.
// Use Cases: Stamp daily reports with a date label on sales charts | Automate dashboard Excel exports with a dynamic date annotation | Replace static chart titles with a runtime date for scheduled generation
// AI Prompts: Write C# code using Aspose.Cells to add a label inside a chart and set its text to DateTime.Now formatted as yyyy-MM-dd. | Show how to find an existing label shape in a chart and update its text to the current date in Aspose.Cells for .NET. | Provide a sample that adds multiple chart labels each with different date formats using Aspose.Cells C#.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Demonstrates how to create a workbook, insert a column chart, add a label shape inside the chart area, set its text to the current date (yyyy‑MM‑dd) via C# and Aspose.Cells, and save the file as an Excel workbook.
class UpdateChartLabel
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate some sample data for the chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["B1"].PutValue(10);
        worksheet.Cells["B2"].PutValue(20);
        worksheet.Cells["B3"].PutValue(30);

        // Add a column chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = worksheet.Charts[chartIndex];
        chart.NSeries.Add("B1:B3", true);          // Set Y values
        chart.NSeries.CategoryData = "A1:A3";      // Set X categories

        // Add a label shape inside the chart area
        // Parameters are top, left, height, width (in 1/4000 of chart area)
        Label chartLabel = chart.Shapes.AddLabelInChart(100, 100, 200, 200);

        // Update the label text to display the current date
        chartLabel.Text = DateTime.Now.ToString("yyyy-MM-dd");

        // Save the workbook with the updated chart label
        workbook.Save("ChartWithCurrentDateLabel.xlsx");
    }
}
