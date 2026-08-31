// Title: Create a column chart in Aspose.Cells, bind its series to a numeric range, and set the category axis using C#
// AI Prompts: Generate C# code with Aspose.Cells that populates month and sales data, adds a column chart, binds the series to the range B2:B7, sets the category axis to A2:A7, adds a chart title, and saves the workbook. | Show how to use Aspose.Cells API to create a worksheet, fill it with numeric values, attach a column chart to those cells, configure the data source and category labels, and export the file as an .xlsx document.
// Common Searches: C# Aspose.Cells bind chart series to a specific cell range | how to set category axis data for a column chart in Aspose.Cells | example of creating a column chart with month labels using Aspose.Cells .NET | Aspose.Cells chart title assignment in C# code | saving an Excel workbook with an embedded chart using Aspose.Cells
// Tags: Aspose.Cells add column chart C# | Aspose.Cells bind series to numeric range | Aspose.Cells set category axis from cells | Aspose.Cells chart title configuration | Aspose.Cells save workbook with chart

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartDemo
{
    // The sample creates a new workbook, writes month names to column A and sales figures to column B, adds a column chart positioned on the sheet, binds the chart series to the numeric range B2:B7, sets the category (X‑axis) to A2:A7, assigns a title "Monthly Sales", and saves the file as MonthlySalesChart.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate numeric data for the chart
            // Column A: Categories (e.g., months)
            // Column B: Values (numeric)
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Sales");
            string[] months = { "Jan", "Feb", "Mar", "Apr", "May", "Jun" };
            for (int i = 0; i < months.Length; i++)
            {
                sheet.Cells[i + 2, 0].PutValue(months[i]);          // A column
                sheet.Cells[i + 2, 1].PutValue((i + 1) * 1000);    // B column (numeric)
            }

            // Add a column chart to the worksheet
            // Parameters: ChartType, topRow, leftColumn, bottomRow, rightColumn
            int chartIndex = sheet.Charts.Add(ChartType.Column, 8, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Bind the series to the numeric data range (B2:B7)
            // The second argument 'true' indicates that data is plotted column‑by‑column (vertical)
            chart.NSeries.Add("=Sheet1!$B$2:$B$7", true);

            // Set the category (X‑axis) data range (A2:A7)
            chart.NSeries.CategoryData = "=Sheet1!$A$2:$A$7";

            // Optional: set a chart title
            chart.Title.Text = "Monthly Sales";

            // Save the workbook to a file
            workbook.Save("MonthlySalesChart.xlsx");
        }
    }
}
