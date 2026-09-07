// Title: Hide the legend of a column chart and confirm no legend entries with Aspose.Cells for .NET
// AI Prompts: Create a new workbook, add a column chart, set chart.ShowLegend = false, then output chart.ShowLegend and chart.Legend.LegendEntries.Count before saving. | Update an existing Aspose.Cells chart to disable its legend and programmatically verify that the LegendEntries collection is empty, printing the results to the console. | Generate an Excel file using Aspose.Cells where the chart legend is turned off, and log both the visibility flag and the number of legend entries for verification.
// Common Searches: asp.net aspose.cells hide chart legend and check legend entries count | how to disable legend in Aspose.Cells generated column chart | verify that no legend entries are rendered in Excel chart using Aspose.Cells .NET | Aspose.Cells ShowLegend false example for column chart
// Tags: Aspose.Cells chart ShowLegend property | disable legend Aspose.Cells column chart | legend entries count Aspose.Cells | Excel chart generation without legend .NET | Aspose.Cells verify legend visibility

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The code creates a workbook, fills cells A1:B4 with sample data, adds a column chart, disables its legend by setting ShowLegend to false, prints the legend visibility flag and the count of legend entries, and saves the file as ChartWithoutLegend.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);          // Values
        chart.NSeries.CategoryData = "A2:A4";      // Categories

        // Hide the legend so it will not be rendered
        chart.ShowLegend = false;

        // Confirm that the legend is hidden
        Console.WriteLine("Legend visible? " + chart.ShowLegend);

        // Optionally, check that no legend entries are rendered
        // (When the legend is hidden, entries are not displayed)
        Console.WriteLine("Legend entries count: " + chart.Legend.LegendEntries.Count);

        // Save the workbook
        workbook.Save("ChartWithoutLegend.xlsx");
    }
}
