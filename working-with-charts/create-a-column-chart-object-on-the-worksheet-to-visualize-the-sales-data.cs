// Title: Create a Column Chart with Aspose.Cells for .NET (C#) – Monthly Sales Example
// Description: C# code that builds a new workbook, fills cells A1:B7 with month and sales figures, inserts a Column chart positioned from row 10, column 2 to row 30, column 10, defines the data range by column, sets the title "Monthly Sales", applies a built‑in style, and saves the workbook as SalesColumnChart.xlsx using Aspose.Cells.
// Keywords: Aspose.Cells | C# | column chart | Excel chart generation | set chart data range | chart title | chart style | save workbook | sales dashboard | Aspose.Cells example
// Common Searches: how to add a column chart in Aspose.Cells C# | Aspose.Cells set chart data range by column | change column chart title Aspose.Cells .NET | apply built‑in style to Aspose.Cells chart | save Excel file with chart using Aspose.Cells
// Developer Intent: Add a column chart to an Excel worksheet to visualize sales figures.
// Use Cases: Automatically generate a monthly sales report with a visual column chart. | Create printable Excel dashboards that include styled charts for quick insight. | Integrate chart creation into a server‑side reporting service that outputs Excel workbooks.
// AI Prompts: Show how to add multiple series to the column chart in Aspose.Cells. | Explain how to customize axis labels and number formats for the chart. | Provide code to export the generated chart as a PNG image using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// C# code that builds a new workbook, fills cells A1:B7 with month and sales figures, inserts a Column chart positioned from row 10, column 2 to row 30, column 10, defines the data range by column, sets the title "Monthly Sales", applies a built‑in style, and saves the workbook as SalesColumnChart.xlsx using Aspose.Cells.
class CreateColumnChart
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample sales data
        // Header
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

        // Add a column chart to the worksheet (top-left row 10, column 2, bottom-right row 30, column 10)
        int chartIndex = sheet.Charts.Add(ChartType.Column, 10, 2, 30, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Define the data range for the chart (including headers)
        chart.SetChartDataRange("A1:B7", true); // true = plot by column

        // Optional: set chart title and style
        chart.Title.Text = "Monthly Sales";
        chart.Style = 2; // Built‑in style

        // Save the workbook
        workbook.Save("SalesColumnChart.xlsx", SaveFormat.Xlsx);
    }
}
