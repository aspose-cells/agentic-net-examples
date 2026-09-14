// Title: How to set the gap width of a column chart series in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code that sets the GapWidth property of the first series in a 2‑D column chart to 200 using Aspose.Cells. | Show an example of increasing bar spacing in an Excel column chart by adjusting the series gap width with Aspose.Cells .NET. | Provide a step‑by‑step C# snippet that creates a column chart, modifies its primary series gap width, and saves the workbook. | Explain how to change the gap width percentage of a column series programmatically with Aspose.Cells.
// Common Searches: Aspose.Cells C# set column chart series gap width to 200 percent | increase spacing between bars in a 2‑D column chart using Aspose.Cells .NET | programmatically adjust column chart gap width in Excel workbook with Aspose.Cells | how to change GapWidth property of a chart series in Aspose.Cells for .NET | example code for modifying bar spacing in Excel column chart using Aspose.Cells
// Tags: set column series gap width Aspose.Cells | increase bar spacing Excel column chart .NET | adjust chart series GapWidth percentage | modify primary series gap width Aspose.Cells | column chart spacing Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, adds sample data, inserts a 2‑D column chart, sets the first series' GapWidth to 200 (200% of the default) to widen bar spacing, and saves the file as AdjustedGapWidth.xlsx.
class AdjustGapWidth
{
    static void Main()
    {
        try
        {
            Run();
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }

    public static void Run()
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

        // Add a 2‑D column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Adjust the gap width of the first (primary axis) series.
        // Value is a percentage of the default width (0‑500). 200 gives wider spacing.
        chart.NSeries[0].GapWidth = 200;

        // Define output file path
        string outputPath = "AdjustedGapWidth.xlsx";

        // Save the workbook with the modified chart
        workbook.Save(outputPath);
        Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
    }
}
