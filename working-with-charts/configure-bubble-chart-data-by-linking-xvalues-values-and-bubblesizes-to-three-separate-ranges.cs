// Title: Configure a bubble chart in Aspose.Cells for .NET by linking XValues, YValues, and BubbleSizes to separate cell ranges
// AI Prompts: Generate C# code using Aspose.Cells that creates a bubble chart and assigns the XValues, Values, and BubbleSizes properties to three distinct cell ranges. | Show how to change the data orientation of a bubble chart series in Aspose.Cells while preserving separate ranges for X, Y, and size values. | Provide a full example that populates sample data, adds a bubble chart, links XValues, Values, and BubbleSizes to ranges, and saves the workbook as an Excel file.
// Common Searches: Aspose.Cells C# assign XValues range for a bubble chart series | binding separate columns to X, Y, and bubble size in Aspose.Cells bubble chart | set bubble chart data source from multiple ranges using Aspose.Cells .NET | example of creating a bubble chart with custom X and size ranges in Aspose.Cells
// Tags: Aspose.Cells bubble chart XValues binding | Aspose.Cells set bubble series values range | Aspose.Cells configure bubble sizes from cells | C# Aspose.Cells create bubble chart example | Excel bubble chart multiple data ranges Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, fills columns A, B, and C with X, Y, and size data, adds a bubble chart, and links the series to the ranges A2:A5 (XValues), B2:B5 (Values), and C2:C5 (BubbleSizes) before saving the file as BubbleChartConfigured.xlsx.
class BubbleChartExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // -------------------------------------------------
        // Populate sample data for the bubble chart
        // Column A : X values
        // Column B : Y values (Values)
        // Column C : Bubble sizes
        // -------------------------------------------------
        sheet.Cells["A1"].PutValue("X");
        sheet.Cells["B1"].PutValue("Y");
        sheet.Cells["C1"].PutValue("Size");

        for (int i = 2; i <= 5; i++)
        {
            sheet.Cells[$"A{i}"].PutValue(i);          // X = i
            sheet.Cells[$"B{i}"].PutValue(i * 2);      // Y = i * 2
            sheet.Cells[$"C{i}"].PutValue(i * 0.5);    // Size = i * 0.5
        }

        // -------------------------------------------------
        // Add a bubble chart to the worksheet
        // Parameters: chart type, top row, left column, bottom row, right column
        // -------------------------------------------------
        int chartIndex = sheet.Charts.Add(ChartType.Bubble, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Add a series using the Y values range (Values)
        // The second argument 'true' indicates that the data is arranged vertically (by column)
        chart.NSeries.Add("B2:B5", true);

        // Link X values and bubble sizes to their respective ranges
        chart.NSeries[0].XValues = "A2:A5";        // X values range
        chart.NSeries[0].BubbleSizes = "C2:C5";   // Bubble sizes range

        // Optional: force chart layout calculation before saving
        chart.Calculate();

        // Save the workbook with the configured bubble chart
        workbook.Save("BubbleChartConfigured.xlsx");
    }
}
