// Title: Set a chart title programmatically with Aspose.Cells for .NET (C#)
// Description: This example creates a new workbook, fills cells A1:B4 with category and value data, adds a column chart, makes the chart title visible, assigns the text "Sales Overview", and saves the file as ChartWithTitle.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | chart title | column chart | set title programmatically | Excel automation | Chart.Title.IsVisible | Chart.Title.Text | .NET chart example
// Common Searches: Aspose.Cells set chart title C# | how to make chart title visible in Aspose.Cells | programmatically add title to Excel chart using Aspose | C# Aspose.Cells column chart with custom title | update chart title after creating chart Aspose.Cells
// Developer Intent: Add a visible, custom‑text title to a chart generated with Aspose.Cells.
// Use Cases: Generate sales dashboards where each column chart is clearly labeled. | Create automated reports that assign titles based on user input or data context. | Produce workbooks with multiple charts, each receiving a distinct, program‑defined title.
// AI Prompts: Write C# code with Aspose.Cells to create a line chart and set its title from a variable. | Show how to change the title of an existing chart after updating its data source in Aspose.Cells. | Explain how to toggle the visibility of a chart title in Aspose.Cells based on a boolean flag.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartTitleDemo
{
    // This example creates a new workbook, fills cells A1:B4 with category and value data, adds a column chart, makes the chart title visible, assigns the text "Sales Overview", and saves the file as ChartWithTitle.xlsx using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Set the chart title programmatically
            chart.Title.IsVisible = true;          // Ensure the title is displayed
            chart.Title.Text = "Sales Overview";   // Assign the desired title text

            // Save the workbook to a file
            workbook.Save("ChartWithTitle.xlsx");
        }
    }
}
