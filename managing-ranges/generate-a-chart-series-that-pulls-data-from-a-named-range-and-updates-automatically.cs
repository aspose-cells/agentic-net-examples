// Title: Create a column chart in Aspose.Cells for .NET that references named ranges for categories and values and updates automatically
// AI Prompts: Write C# code using Aspose.Cells to define a named range for month labels and another for sales data, then add a column chart whose series source points to those named ranges. | Show how to set the chart's category axis and value series to named ranges so the chart refreshes when the underlying cells change. | Modify the example to use a dynamic named range that expands automatically as new rows are added, keeping the chart up‑to‑date.
// Common Searches: Aspose.Cells .NET create column chart from named range that auto updates | how to bind chart series to named range in C# using Aspose.Cells | dynamic named range for Excel chart with Aspose.Cells .NET | auto refresh Excel column chart when data changes Aspose.Cells C# | using named ranges for chart categories and values Aspose.Cells example
// Tags: Aspose.Cells column chart named range | C# chart series from named range | auto updating Excel chart Aspose.Cells | dynamic named range for chart data .NET | bind chart categories to named range Aspose

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;

// Demonstrates creating a workbook, defining named ranges for months and sales, adding a column chart, and linking the chart series to those named ranges so the chart updates automatically.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (months and sales)
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Sales");
            string[] months = { "Jan", "Feb", "Mar", "Apr", "May" };
            double[] sales = { 1200, 1500, 1800, 1300, 1700 };
            for (int i = 0; i < months.Length; i++)
            {
                sheet.Cells[i + 1, 0].PutValue(months[i]);   // Column A (Month)
                sheet.Cells[i + 1, 1].PutValue(sales[i]);   // Column B (Sales)
            }

            // Define a named range for the sales values (including header) - B1:B6
            int firstRow = 0;          // zero‑based index for row 1
            int firstColumn = 1;       // zero‑based index for column B
            int totalRows = months.Length + 1; // header + data rows
            int totalColumns = 1;
            sheet.Cells.CreateRange(firstRow, firstColumn, totalRows, totalColumns).Name = "SalesData";

            // Define a named range for the month categories (including header) - A1:A6
            sheet.Cells.CreateRange(firstRow, 0, totalRows, 1).Name = "MonthCategories";

            // Add a column chart to the worksheet (position: row 6, column 0 to row 20, column 10)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Use the named ranges for the series
            chart.NSeries.Add("=Sheet1!MonthCategories", true);
            chart.NSeries[0].Values = "=Sheet1!SalesData";

            // Set chart title
            chart.Title.Text = "Monthly Sales";

            // Save the workbook
            workbook.Save("ChartWithNamedRange.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
