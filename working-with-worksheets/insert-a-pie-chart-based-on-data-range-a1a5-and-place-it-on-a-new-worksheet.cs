// Title: C# – Insert a Pie Chart from A1:A5 onto a New Worksheet using Aspose.Cells
// Description: Demonstrates how to create a workbook, add a new worksheet, populate cells A1‑A5 with numeric data, place a Pie chart (rows 6‑20, columns 0‑8) that references the range A1:A5, set a chart title, and save the file as an Excel workbook with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | pie chart | new worksheet | chart data range A1:A5 | add chart programmatically | chart positioning | save Excel workbook | Excel automation
// Common Searches: Aspose.Cells add pie chart C# | create chart on new worksheet Aspose.Cells | set pie chart data source A1:A5 | position chart rows columns Aspose.Cells | save workbook with chart Aspose.Cells
// Developer Intent: Programmatically add a Pie chart based on cells A1:A5 to a freshly created worksheet and persist the workbook.
// Use Cases: Generate a sales‑by‑product pie chart for monthly reporting. | Build a financial dashboard that visualizes expense categories as a pie chart. | Automate batch creation of Excel files that include pre‑formatted charts for data analysis.
// AI Prompts: Show C# code to add a donut chart from range B2:B6 on an existing worksheet with Aspose.Cells. | Explain how to customize pie chart colors, legend placement, and data labels in Aspose.Cells. | Provide a method to export a pie chart created with Aspose.Cells as a PNG image.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates how to create a workbook, add a new worksheet, populate cells A1‑A5 with numeric data, place a Pie chart (rows 6‑20, columns 0‑8) that references the range A1:A5, set a chart title, and save the file as an Excel workbook with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add a new worksheet that will contain the pie chart
        int newSheetIndex = workbook.Worksheets.Add();
        Worksheet chartSheet = workbook.Worksheets[newSheetIndex];
        chartSheet.Name = "PieChartSheet";

        // Populate sample numeric data in the range A1:A5
        chartSheet.Cells["A1"].PutValue(10);
        chartSheet.Cells["A2"].PutValue(20);
        chartSheet.Cells["A3"].PutValue(30);
        chartSheet.Cells["A4"].PutValue(25);
        chartSheet.Cells["A5"].PutValue(15);

        // Add a pie chart to the worksheet
        // Parameters: chart type, top row, left column, bottom row, right column
        int chartIndex = chartSheet.Charts.Add(ChartType.Pie, 6, 0, 20, 8);
        Chart pieChart = chartSheet.Charts[chartIndex];

        // Set the data source for the pie chart using the specified range A1:A5
        pieChart.NSeries.Add("A1:A5", true);

        // Optional: give the chart a title
        pieChart.Title.Text = "Sample Pie Chart";

        // Save the workbook to a file
        workbook.Save("PieChartWorkbook.xlsx");
    }
}
