// Title: Create a column chart from a named range and refresh it after updating the range values using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that defines a named range, sets it as the data source for a column chart, changes the cell values inside the range, and saves the workbook with Aspose.Cells. | Demonstrate how to programmatically refresh a chart after modifying the cells referenced by a named range in an Aspose.Cells workbook.
// Common Searches: how to bind a named range to a chart series in Aspose.Cells C# | Aspose.Cells refresh chart after changing named range values | C# Aspose.Cells column chart using named range as source | update chart automatically when named range data changes Aspose.Cells .NET
// Tags: named range as chart data source Aspose.Cells | column chart bound to named range C# | refresh chart after cell update Aspose.Cells | modify named range values Aspose.Cells | create workbook with chart from named range .NET

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// Shows how to create a workbook, define a named range, bind it to a column chart, update the range values, and save the file so the chart reflects the changes using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (including headers)
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(30);

            // Define a named range that covers the data (including headers)
            sheet.Cells.CreateRange(0, 0, 4, 2).Name = "ChartData";

            // Add a column chart to the worksheet
            int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIdx];

            // Set the chart's data source to the named range
            chart.NSeries.Add("ChartData!$B$2:$B$4", true);
            chart.NSeries.CategoryData = "ChartData!$A$2:$A$4";

            // Modify the data within the named range
            sheet.Cells["B2"].PutValue(15);
            sheet.Cells["B3"].PutValue(25);
            sheet.Cells["B4"].PutValue(35);

            // Save the workbook
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "NamedRangeChart.xlsx");
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
