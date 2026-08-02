// Title: Create a Column Chart in Aspose.Cells for .NET and Bind Its Series to a Numeric Range (C#)
// Description: Demonstrates how to generate a new workbook, fill cells A1:B13 with month names and sales figures, add a column chart, bind the series to the numeric range B2:B13, set month names from A2:A13 as category labels, apply a chart title, and save the file as MonthlySalesChart.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells chart creation C# | bind chart series numeric range | column chart Aspose.Cells .NET | set category axis labels Excel | programmatic Excel chart Aspose | C# Excel chart example | Aspose.Cells add chart to worksheet
// Common Searches: Aspose.Cells bind series to numeric range | Create column chart with Aspose.Cells C# | Set X axis labels for Excel chart using Aspose | Add and save chart in workbook Aspose.Cells | How to programmatically create charts in .NET Excel
// Developer Intent: Generate a column chart, link its data series to a numeric range, and define category labels programmatically in a worksheet.
// Use Cases: Automated monthly sales reporting with a visual column chart. | Dynamic dashboard generation where data ranges change each period. | Batch creation of Excel workbooks that include pre‑formatted charts for business presentations.
// AI Prompts: Write C# code with Aspose.Cells to create a line chart bound to a numeric range and assign month names as categories. | Show how to add multiple data series from different ranges to a single chart in Aspose.Cells. | Explain how to refresh chart data after updating worksheet values using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartDemo
{
    // Demonstrates how to generate a new workbook, fill cells A1:B13 with month names and sales figures, add a column chart, bind the series to the numeric range B2:B13, set month names from A2:A13 as category labels, apply a chart title, and save the file as MonthlySalesChart.xlsx using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate numeric data for the chart
            // Header
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Sales");

            // Sample data (Month, Sales)
            for (int i = 2; i <= 13; i++)
            {
                sheet.Cells[$"A{i}"].PutValue($"M{i - 1}");
                sheet.Cells[$"B{i}"].PutValue((i - 1) * 1000); // Example sales values
            }

            // Add a column chart to the worksheet
            // Parameters: ChartType, topRow, leftColumn, bottomRow, rightColumn
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 2, 25, 11);
            Chart chart = sheet.Charts[chartIndex];

            // Bind the series to the numeric data range (B2:B13)
            // The second argument 'true' indicates that data is plotted column‑by‑column (vertical)
            chart.NSeries.Add("=Sheet1!$B$2:$B$13", true);

            // Set the category (X‑axis) labels to the month names (A2:A13)
            chart.NSeries.CategoryData = "=Sheet1!$A$2:$A$13";

            // Optional: set a title for the chart
            chart.Title.Text = "Monthly Sales";

            // Save the workbook to a file
            workbook.Save("MonthlySalesChart.xlsx");
        }
    }
}
