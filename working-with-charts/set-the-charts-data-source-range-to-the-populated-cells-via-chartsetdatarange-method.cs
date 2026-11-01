// Title: How to set a column chart’s data source range to worksheet cells using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a new workbook, fills cells A1:B4 with sample data, adds a column chart, and links its series to the range B2:B4 using Aspose.Cells Chart API. | Show how to change the data range of an existing Aspose.Cells chart to a different cell range in C#. | Provide a step‑by‑step example that populates worksheet cells, inserts a column chart, and binds the chart’s values to a specified range with Aspose.Cells.
// Common Searches: Aspose.Cells C# set chart series range to specific cells | bind Excel column chart to cell range using Aspose.Cells .NET | example code for Chart.NSeries.Add with range B2:B4 in C# | how to programmatically assign data source to Aspose.Cells chart | C# Aspose.Cells create chart and set data source range
// Tags: Aspose.Cells chart series binding C# | set chart data source range Aspose.Cells | column chart creation Aspose.Cells .NET | C# Excel chart data range assignment | Aspose.Cells workbook chart example

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartExample
{
    // The sample creates a new workbook, writes category and value data into cells A1:B4, adds a column chart positioned on the sheet, binds the chart’s series to the values in B2:B4 using the NSeries.Add method, and saves the workbook as ChartDataRange.xlsx while handling any exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                var workbook = new Workbook();

                // Access the first worksheet
                var sheet = workbook.Worksheets[0];

                // Populate cells with sample data
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B4"].PutValue(30);

                // Add a column chart to the worksheet
                // Parameters: chart type, upper-left row, upper-left column, lower-right row, lower-right column
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                var chart = sheet.Charts[chartIndex];

                // Set the chart's data source range (values series)
                chart.NSeries.Add("B2:B4", true);
                // Category data can be set if needed; omitted here to avoid API compatibility issues

                // Save the workbook to a file
                workbook.Save("ChartDataRange.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
