// Title: Create a column chart from sample data and save the workbook as an XLSX file with Aspose.Cells for .NET
// AI Prompts: Generate C# code that uses Aspose.Cells to fill a worksheet with data, add a column chart linked to a specific range, and save the workbook as an .xlsx file. | Show how to bind a data series to a column chart in Aspose.Cells and export the resulting workbook to XLSX format. | Provide a complete Aspose.Cells example that creates a chart, configures its NSeries, and calls Workbook.Save with SaveFormat.Xlsx.
// Common Searches: aspnet aspose.cells how to add a column chart and save as xlsx | c# create chart from range and export workbook to xlsx using Aspose.Cells | example of binding NSeries to chart in Aspose.Cells .NET | save workbook containing chart as xlsx file Aspose.Cells | asp.net core Aspose.Cells chart creation and xlsx export
// Tags: Aspose.Cells create column chart C# | Aspose.Cells bind NSeries to chart | Aspose.Cells save workbook as XLSX | Aspose.Cells chart data range binding | Aspose.Cells export chart workbook .xlsx

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;

// The program creates a new workbook, populates cells A1:B4 with sample data, adds a column chart referencing the values, and saves the workbook as GeneratedChart.xlsx in XLSX format using Aspose.Cells.
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

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 5);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart series
            chart.NSeries.Add("B2:B4", true);
            // Category data line removed because Series.CategoryData is not available in this API version

            // Save the workbook containing the chart as an XLSX file
            workbook.Save("GeneratedChart.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
