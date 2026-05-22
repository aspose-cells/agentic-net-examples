using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook wb = new Workbook();

        // Set the regional setting for the workbook (example: Japan)
        wb.Settings.Region = CountryCode.Japan;

        // Get the first worksheet
        Worksheet sheet = wb.Worksheets[0];

        // Populate some sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Q1");
        sheet.Cells["A3"].PutValue("Q2");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(100);
        sheet.Cells["B3"].PutValue(200);

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = sheet.Charts[chartIndex];

        // Set chart title (used as the chart name for logging)
        chart.Title.Text = "Sales Chart";

        // Bind data to the chart
        chart.NSeries.Add("B2:B3", true);
        chart.NSeries.CategoryData = "A2:A3";

        // Prepare diagnostics information
        string diagnosticsFile = "diagnostics.txt";
        string logEntry = $"Region: {wb.Settings.Region}, ChartTitle: {chart.Title.Text}";

        // Append the log entry to the diagnostics file
        File.AppendAllText(diagnosticsFile, logEntry + Environment.NewLine);

        // Save the workbook
        wb.Save("Output.xlsx");
    }
}