// Title: Create a chart template with predefined data label settings and reuse it in another workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that builds a source workbook, adds a column chart, configures its data labels to show values and category names, and then applies the same label configuration to a chart in a separate workbook. | Demonstrate how to programmatically copy chart data‑label settings from a template chart to multiple new charts across different workbooks with Aspose.Cells for .NET.
// Common Searches: asp.net aspose.cells create chart template with data labels | c# copy chart data label configuration between workbooks using Aspose.Cells | how to reuse chart formatting as a template in Aspose.Cells .NET | apply predefined data label settings to a new column chart in Aspose.Cells C# | Aspose.Cells example for chart template and data label reuse
// Tags: Aspose.Cells chart template creation | C# set chart data labels Aspose.Cells | copy chart formatting between workbooks Aspose.Cells | column chart data label configuration .NET | reuse chart settings across workbooks Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The program creates a source workbook with sample data and a column chart, configures the chart's data labels to display both values and category names, then creates a second workbook, adds a new column chart with different data, applies the same data label settings from the template chart, and saves the result as WorkbookWithTemplateChart.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // ------------------------------------------------------------
            // 1. Create a source workbook, add data and a chart (template source)
            // ------------------------------------------------------------
            Workbook srcWb = new Workbook();
            Worksheet srcWs = srcWb.Worksheets[0];

            // Sample data
            srcWs.Cells["A1"].PutValue("Category");
            srcWs.Cells["B1"].PutValue("Value");
            srcWs.Cells["A2"].PutValue("A");
            srcWs.Cells["A3"].PutValue("B");
            srcWs.Cells["A4"].PutValue("C");
            srcWs.Cells["B2"].PutValue(10);
            srcWs.Cells["B3"].PutValue(20);
            srcWs.Cells["B4"].PutValue(30);

            // Add a column chart
            int chartIdx = srcWs.Charts.Add(ChartType.Column, 5, 0, 20, 5);
            Chart srcChart = srcWs.Charts[chartIdx];
            srcChart.NSeries.Add("B2:B4", true);
            srcChart.NSeries.CategoryData = "A2:A4";

            // Configure data labels for the first series
            srcChart.NSeries[0].DataLabels.ShowValue = true;
            srcChart.NSeries[0].DataLabels.ShowCategoryName = true;
            // Position property may not be available in older versions; omit if not supported.

            // ------------------------------------------------------------
            // 2. Create a new workbook and add a chart that mimics the template settings
            // ------------------------------------------------------------
            Workbook destWb = new Workbook();
            Worksheet destWs = destWb.Worksheets[0];

            // Different data for the new chart
            destWs.Cells["A1"].PutValue("Item");
            destWs.Cells["B1"].PutValue("Amount");
            destWs.Cells["A2"].PutValue("X");
            destWs.Cells["A3"].PutValue("Y");
            destWs.Cells["A4"].PutValue("Z");
            destWs.Cells["B2"].PutValue(15);
            destWs.Cells["B3"].PutValue(25);
            destWs.Cells["B4"].PutValue(35);

            // Add an empty chart placeholder
            int destChartIdx = destWs.Charts.Add(ChartType.Column, 5, 0, 20, 5);
            Chart destChart = destWs.Charts[destChartIdx];

            // Apply data label settings similar to the source chart
            destChart.NSeries[0].DataLabels.ShowValue = true;
            destChart.NSeries[0].DataLabels.ShowCategoryName = true;

            // Set the data source for the new chart
            destChart.NSeries.Add("B2:B4", true);
            destChart.NSeries.CategoryData = "A2:A4";

            // Save the final workbook
            string outputPath = "WorkbookWithTemplateChart.xlsx";
            destWb.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
