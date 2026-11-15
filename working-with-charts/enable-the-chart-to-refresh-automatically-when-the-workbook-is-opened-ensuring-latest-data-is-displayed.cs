// Title: How to make an Excel chart refresh automatically on workbook open with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set chart to refresh on workbook open | auto refresh chart data when opening an Excel file using Aspose.Cells .NET | enable RefreshOnLoad for a chart in Aspose.Cells C# example | C# Aspose.Cells chart updates automatically on file load | Excel chart automatic refresh on open with Aspose.Cells library
// Tags: Aspose.Cells chart RefreshOnLoad property | C# auto refresh Excel chart on open | programmatic chart data refresh Aspose.Cells | Excel workbook open chart update .NET | dynamic chart refresh with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, populates sample data, adds a column chart, sets the chart's RefreshOnLoad property to true so the chart updates each time the workbook is opened, and saves the file as ChartRefresh.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 5);
            Chart chart = sheet.Charts[chartIndex];

            // Add series with both category and value ranges (A2:B4)
            chart.NSeries.Add("A2:B4", true);

            // Determine output file path
            string outputFile = "ChartRefresh.xlsx";
            string outputPath = Path.GetFullPath(outputFile);

            // Ensure the directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
