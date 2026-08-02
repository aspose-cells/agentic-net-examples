// Title: C# – Insert a Pie Chart from A1:A5 on a New Worksheet with Aspose.Cells
// Description: Creates a workbook, adds a separate sheet, fills A1:A5 with categories and B1:B5 with values, inserts a pie chart linked to those ranges, sets an optional title, positions the chart, and saves the file as PieChartWorkbook.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# pie chart | add chart new worksheet .NET | set chart data range A1:A5 | pie chart category labels Aspose | C# Aspose.Cells example | programmatic chart positioning | save workbook with chart | GitHub Aspose.Cells samples | Excel pie chart automation | Aspose.Cells chart title
// Common Searches: how to add a pie chart to a new sheet using Aspose.Cells C# | Aspose.Cells set category data for pie chart | C# code to create and position a pie chart in Excel | Aspose.Cells example for chart series and labels | generate pie chart programmatically with Aspose.Cells
// Developer Intent: Generate a dedicated worksheet, populate category/value cells, and embed a pie chart that references those ranges.
// Use Cases: Automated sales‑by‑region visualization on its own sheet for monthly reports. | Financial budget allocation pie chart placed in a separate worksheet for dashboard distribution. | Product‑mix distribution chart added to a workbook that will be shared with external stakeholders.
// AI Prompts: Show how to display data labels with percentages on the Aspose.Cells pie chart. | Provide code to convert the pie chart to a 3‑D pie and adjust its dimensions. | Explain how to export only the chart worksheet as an image using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds a separate sheet, fills A1:A5 with categories and B1:B5 with values, inserts a pie chart linked to those ranges, sets an optional title, positions the chart, and saves the file as PieChartWorkbook.xlsx using Aspose.Cells for .NET.
class InsertPieChart
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add a new worksheet that will contain the pie chart
        int newSheetIndex = workbook.Worksheets.Add();
        Worksheet chartSheet = workbook.Worksheets[newSheetIndex];
        chartSheet.Name = "PieChartSheet";

        // Populate sample categories in A1:A5
        chartSheet.Cells["A1"].PutValue("Category1");
        chartSheet.Cells["A2"].PutValue("Category2");
        chartSheet.Cells["A3"].PutValue("Category3");
        chartSheet.Cells["A4"].PutValue("Category4");
        chartSheet.Cells["A5"].PutValue("Category5");

        // Populate corresponding values in B1:B5
        chartSheet.Cells["B1"].PutValue(10);
        chartSheet.Cells["B2"].PutValue(20);
        chartSheet.Cells["B3"].PutValue(30);
        chartSheet.Cells["B4"].PutValue(25);
        chartSheet.Cells["B5"].PutValue(15);

        // Add a pie chart to the worksheet (positioned from row 5, column 0 to row 20, column 8)
        int chartIndex = chartSheet.Charts.Add(ChartType.Pie, 5, 0, 20, 8);
        Chart chart = chartSheet.Charts[chartIndex];

        // Set the data range for the chart (values) and category labels
        chart.NSeries.Add("B1:B5", true);
        chart.NSeries.CategoryData = "A1:A5";

        // Optional: set a title for the chart
        chart.Title.Text = "Sample Pie Chart";

        // Save the workbook to a file
        workbook.Save("PieChartWorkbook.xlsx");
    }
}
