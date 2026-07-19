// Title: Create a Column Chart in Aspose.Cells for .NET and Bind Its Series to a Numeric Range
// Description: Shows how to create a new workbook, populate month labels and sales figures, add a column chart, bind the series to cells B2:B13 with NSeries.Add, set category labels from A2:A13, give the chart a title, and save the file as an .xlsx using C#.
// Keywords: Aspose.Cells C# | column chart .NET | bind chart series numeric range | NSeries.Add example | Excel chart automation | C# Aspose.Cells tutorial | chart category data | save workbook with chart
// Common Searches: Aspose.Cells create column chart C# | bind chart series to numeric range Aspose.Cells | set X axis categories Aspose.Cells chart | save Excel file with chart using Aspose.Cells | C# example for chart data binding Aspose.Cells
// Developer Intent: Generate an Excel workbook that programmatically adds a column chart whose data series is linked to a numeric range and whose X‑axis displays month labels.
// Use Cases: Automated monthly sales reporting with visual charts | Dynamic dashboards that update chart data from calculations | Financial workbooks where chart series reflect changing numeric ranges | Batch generation of Excel files with pre‑formatted charts for distribution
// AI Prompts: Show C# code to create a line chart in Aspose.Cells and bind its series to a numeric range with category labels. | How can I update an existing Aspose.Cells chart's data source at runtime in .NET? | Provide examples of formatting the chart title, axes, and legend after binding data in Aspose.Cells C#.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartDemo
{
    // Shows how to create a new workbook, populate month labels and sales figures, add a column chart, bind the series to cells B2:B13 with NSeries.Add, set category labels from A2:A13, give the chart a title, and save the file as an .xlsx using C#.
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

            // Sample data (numeric values)
            for (int i = 2; i <= 13; i++)
            {
                sheet.Cells[$"A{i}"].PutValue($"M{i - 1}");
                sheet.Cells[$"B{i}"].PutValue(i * 100); // Example sales figures
            }

            // Add a column chart to the worksheet
            // Parameters: ChartType, topRow, leftColumn, bottomRow, rightColumn
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 2, 25, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Bind the series to the numeric data range (values column)
            // Using NSeries.Add with the data area and vertical orientation
            chart.NSeries.Add("=Sheet1!$B$2:$B$13", true);

            // Set the category (X‑axis) data to the month labels
            chart.NSeries.CategoryData = "=Sheet1!$A$2:$A$13";

            // Optional: give the chart a title
            chart.Title.Text = "Monthly Sales";

            // Save the workbook to a file
            workbook.Save("MonthlySalesChart.xlsx");
        }
    }
}
