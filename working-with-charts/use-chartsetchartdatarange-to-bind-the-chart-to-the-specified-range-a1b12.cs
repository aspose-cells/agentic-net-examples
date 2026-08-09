// Title: Bind a Column Chart to Range A1:B12 with Chart.SetChartDataRange in Aspose.Cells for .NET (C#)
// Description: This example creates a new workbook, fills cells A1:B12 with category labels and numeric values, adds a column chart, and uses Chart.SetChartDataRange to bind the chart to that range as a vertical series. The workbook is saved as ChartWithDataRange.xlsx.
// Keywords: Aspose.Cells Chart.SetChartDataRange | C# bind chart to cell range | Aspose.Cells column chart example | set chart data range vertical series | Aspose.Cells .NET chart binding | Excel chart programmatic creation
// Common Searches: C# Aspose.Cells set chart data range | How to bind a column chart to A1:B12 using Aspose.Cells | Aspose.Cells Chart.SetChartDataRange vertical series example | Create Excel column chart from range programmatically | Aspose.Cells chart data source binding code
// Developer Intent: Bind a column chart to cells A1:B12 using Chart.SetChartDataRange.
// Use Cases: Generate a monthly sales report where the column chart updates automatically as values in A1:B12 change. | Build a KPI dashboard that visualizes real‑time data by linking charts directly to a predefined cell range. | Export pre‑formatted Excel files with embedded charts for downstream analysis without manual chart setup.
// AI Prompts: Show how to switch the chart to a horizontal series using Chart.SetChartDataRange. | Provide code to add data labels and customize column colors after binding the chart to A1:B12. | Explain how to bind multiple series from separate ranges with Chart.SetChartDataRange in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartDataRangeDemo
{
    // This example creates a new workbook, fills cells A1:B12 with category labels and numeric values, adds a column chart, and uses Chart.SetChartDataRange to bind the chart to that range as a vertical series. The workbook is saved as ChartWithDataRange.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data in the range A1:B12
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            for (int i = 2; i <= 12; i++)
            {
                sheet.Cells[$"A{i}"].PutValue($"Item {i - 1}");
                sheet.Cells[$"B{i}"].PutValue(i * 10); // Example numeric values
            }

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Bind the chart to the data range A1:B12 (vertical series)
            chart.SetChartDataRange("A1:B12", true);

            // Optional: set a title for clarity
            chart.Title.Text = "Sample Column Chart";

            // Save the workbook to a file
            workbook.Save("ChartWithDataRange.xlsx");
        }
    }
}
