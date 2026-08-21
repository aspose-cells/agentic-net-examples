// Title: Create a Column Chart with Aspose.Cells for .NET (C#) – Monthly Sales Example
// Description: C# code that uses Aspose.Cells to build an XLSX workbook, writes month and sales values to A1:B7, inserts a Column chart positioned rows 8‑20 and columns A‑G, sets the data range, adds a chart title and legend, and saves the file as SalesColumnChart.xlsx.
// Keywords: Aspose.Cells | C# | column chart | Excel chart | sales data | ChartType.Column | SetChartDataRange | SaveFormat.Xlsx | Aspose.Cells example | CreateColumnChart
// Common Searches: Aspose.Cells add column chart C# | How to set chart data range Aspose.Cells .NET | Create sales chart with Aspose.Cells | Aspose.Cells chart title and legend | Generate Excel column chart programmatically
// Developer Intent: Create and embed a column chart in an Excel worksheet using Aspose.Cells for .NET to visualize monthly sales figures.
// Use Cases: Automated generation of monthly sales reports with embedded charts | Building Excel dashboards that include column charts for product performance | Exporting data analysis results to Excel with pre‑formatted visualizations
// AI Prompts: Generate C# code to add a stacked column chart with two series using Aspose.Cells. | Show how to customize column colors, axis titles, and data labels in an Aspose.Cells chart. | Explain how to modify an existing chart's data range and title after the workbook is saved.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// C# code that uses Aspose.Cells to build an XLSX workbook, writes month and sales values to A1:B7, inserts a Column chart positioned rows 8‑20 and columns A‑G, sets the data range, adds a chart title and legend, and saves the file as SalesColumnChart.xlsx.
class CreateColumnChart
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample sales data
        // Header row
        sheet.Cells["A1"].PutValue("Month");
        sheet.Cells["B1"].PutValue("Sales");

        // Data rows
        string[] months = { "Jan", "Feb", "Mar", "Apr", "May", "Jun" };
        int[] sales = { 12000, 15000, 13000, 17000, 16000, 18000 };

        for (int i = 0; i < months.Length; i++)
        {
            sheet.Cells[i + 2, 0].PutValue(months[i]);   // Column A
            sheet.Cells[i + 2, 1].PutValue(sales[i]);   // Column B
        }

        // Add a column chart to the worksheet
        // Parameters: ChartType, topRow, leftColumn, bottomRow, rightColumn
        int chartIndex = sheet.Charts.Add(ChartType.Column, 8, 0, 20, 7);
        Chart chart = sheet.Charts[chartIndex];

        // Define the data range for the chart (including headers)
        // The second argument 'true' indicates that data is plotted by column
        chart.SetChartDataRange("A1:B7", true);

        // Optional: set chart title and enable legend
        chart.Title.Text = "Monthly Sales";
        chart.ShowLegend = true;

        // Save the workbook to an XLSX file
        workbook.Save("SalesColumnChart.xlsx", SaveFormat.Xlsx);
    }
}
