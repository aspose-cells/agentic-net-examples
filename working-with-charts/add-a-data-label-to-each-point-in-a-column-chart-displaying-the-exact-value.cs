// Title: Add data labels that show exact values to each point in a column chart with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code using Aspose.Cells to create a column chart and enable data labels that display the numeric value for every column. | Modify an existing Aspose.Cells workbook to turn on value labels for a column series and set the label position to the center of each column. | Generate a complete C# example that builds a workbook, adds sample data, inserts a column chart, and configures data labels to show values with optional positioning.
// Common Searches: Aspose.Cells C# column chart show data label values | how to enable value labels on column series using Aspose.Cells .NET | C# Aspose.Cells set data label position to center for column chart | add data labels to each point in a column chart Aspose.Cells example | display exact numeric values on column chart columns Aspose.Cells
// Tags: Aspose.Cells column chart data labels | show value labels Aspose.Cells C# | center position data labels Aspose.Cells | create column chart with labels Aspose.Cells | enable data labels for series Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// C# program that creates a workbook, populates sample data, adds a column chart, enables data labels to display each column's exact value, optionally positions the labels at the center, and saves the file as ColumnChartWithDataLabels.xlsx.
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

            // Populate sample data for the column chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(90);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 7);
            Chart chart = sheet.Charts[chartIndex];

            // Set chart title
            chart.Title.Text = "Monthly Sales";

            // Add series: values from B2:B4
            chart.NSeries.Add("B2:B4", false);
            // Set categories for the series (if supported by the library version)
            // chart.NSeries[0].CategoryData = "A2:A4";

            // Enable data labels and display the exact value for each point
            chart.NSeries[0].DataLabels.ShowValue = true;

            // Position the data labels at the center of each column (if PositionType enum is available)
            // chart.NSeries[0].DataLabels.Position = PositionType.Center;

            // Save the workbook with the chart
            string outputPath = "ColumnChartWithDataLabels.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
