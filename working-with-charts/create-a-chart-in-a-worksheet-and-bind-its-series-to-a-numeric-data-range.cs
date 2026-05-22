using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Populate numeric data for the chart
        // Column A: Categories (e.g., months)
        // Column B: Values (e.g., sales)
        sheet.Cells["A1"].PutValue("Month");
        sheet.Cells["B1"].PutValue("Sales");
        string[] months = { "Jan", "Feb", "Mar", "Apr", "May", "Jun" };
        for (int i = 0; i < months.Length; i++)
        {
            sheet.Cells[i + 2, 0].PutValue(months[i]);          // A column
            sheet.Cells[i + 2, 1].PutValue((i + 1) * 1000);    // B column
        }

        // Add a column chart to the worksheet
        // Parameters: ChartType, topRow, leftColumn, bottomRow, rightColumn
        int chartIndex = sheet.Charts.Add(ChartType.Column, 8, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Bind the series to the numeric data range (B2:B7)
        // The second argument 'true' indicates that data is plotted by column (vertical)
        chart.NSeries.Add("=Sheet1!$B$2:$B$7", true);

        // Set the category (X‑axis) data range (A2:A7)
        chart.NSeries.CategoryData = "=Sheet1!$A$2:$A$7";

        // Optional: set a chart title
        chart.Title.Text = "Monthly Sales";

        // Save the workbook to a file
        workbook.Save("ChartWithNumericSeries.xlsx");
    }
}