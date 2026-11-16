// Title: Create a column chart and bind its data range with SetChartDataRange using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells how to bind data range to a column chart | SetChartDataRange usage example for column charts in .NET | Create Excel column chart programmatically with Aspose.Cells | Assign categories and values to Aspose.Cells chart using SetChartDataRange | Aspose.Cells column chart data source from worksheet cells C#
// Tags: Aspose.Cells SetChartDataRange column chart | C# create Excel column chart Aspose.Cells | bind chart data range Aspose.Cells .NET | programmatic Excel chart generation C# | Excel workbook column chart data source

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a new workbook, fills cells A1:B4 with category and value data, adds a column chart to the first worksheet, sets the chart's data range (including categories and values) using SetChartDataRange, and saves the file as ColumnChart.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            // Parameters: chart type, upper-left row, upper-left column, lower-right row, lower-right column
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 5);
            Chart chart = sheet.Charts[chartIndex];

            // Assign the data source for the chart (including categories and values)
            // The second argument indicates whether the data series are plotted vertically.
            chart.SetChartDataRange("Sheet1!$A$1:$B$4", true);

            // Define output file path
            string outputPath = "ColumnChart.xlsx";

            // Save the workbook with the chart
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
