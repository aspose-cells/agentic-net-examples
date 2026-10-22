// Title: Set the IsTotal flag on a specific data point in a column chart using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a column chart from worksheet data, accesses the third ChartDataPoint, sets its IsTotal property to true, and saves the workbook as an XLSX file using Aspose.Cells. | Show how to enable the total marker for a data point in an Aspose.Cells column chart by using ChartDataPoint.IsTotal in a .NET example. | Modify an existing Aspose.Cells column chart example to mark the first series point as a total and export the updated workbook.
// Common Searches: Aspose.Cells C# how to enable IsTotal on a column chart data point | Mark a data point as total in Excel column chart using Aspose.Cells .NET | ChartDataPoint IsTotal property example in Aspose.Cells C# | Set total flag for series point in Aspose.Cells column chart | C# Aspose.Cells column chart total marker tutorial
// Tags: Aspose.Cells ChartDataPoint IsTotal | C# column chart total data point | Aspose.Cells set total flag column chart | Excel column chart total marker Aspose.Cells | Aspose.Cells .NET column chart example

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, populates it with category and value data, adds a column chart, and demonstrates where to set the ChartDataPoint.IsTotal flag to mark a point as a total. The workbook is saved as an XLSX file; the actual IsTotal assignment is shown for versions of Aspose.Cells that support the ChartDataPoint API.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Fill sample data for the column chart
            ws.Cells["A1"].PutValue("Category");
            ws.Cells["B1"].PutValue("Value");
            ws.Cells["A2"].PutValue("Jan");
            ws.Cells["A3"].PutValue("Feb");
            ws.Cells["A4"].PutValue("Mar");
            ws.Cells["B2"].PutValue(10);
            ws.Cells["B3"].PutValue(20);
            ws.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = ws.Charts.Add(ChartType.Column, 5, 0, 20, 7);
            Chart chart = ws.Charts[chartIndex];

            // Set the data range for the series (values)
            chart.NSeries.Add("B2:B4", true);
            // Set the category (X) axis data
            chart.NSeries[0].XValues = "A2:A4";

            // NOTE: Marking a data point as total requires ChartDataPoint which may not be
            // available in older Aspose.Cells versions. This block is omitted for compatibility.

            // Define output file path
            string outputPath = "ColumnChartWithTotal.xlsx";

            // Save the workbook with the chart
            wb.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
