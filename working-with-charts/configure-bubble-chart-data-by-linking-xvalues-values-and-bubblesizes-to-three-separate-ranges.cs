// Title: Configure a Bubble Chart in Aspose.Cells .NET – Link XValues, Y Values, and BubbleSizes to Separate Ranges
// Description: Shows how to create a workbook, fill columns A‑C with X, Y, and size data, add a Bubble chart, set the Y‑range as Values, then bind XValues to A2:A5 and BubbleSizes to C2:C5, calculate the layout, and save the file.
// Keywords: Aspose.Cells bubble chart | C# bubble chart XValues | set bubble sizes Aspose.Cells | link series ranges | Aspose.Cells chart series | bubble chart .NET example | configure bubble chart data | Aspose.Cells NSeries XValues | Aspose.Cells BubbleSizes
// Common Searches: Aspose.Cells set XValues for bubble chart | How to bind bubble size data in Aspose.Cells C# | Bubble chart series separate ranges Aspose.Cells | Create bubble chart with custom X and size columns Aspose.Cells | Aspose.Cells bubble chart example C#
// Developer Intent: Assign distinct cell ranges to the X‑axis, Y‑axis, and bubble‑size dimensions of a bubble chart using Aspose.Cells for .NET.
// Use Cases: Visualize regional sales where column A holds region index (X), column B holds revenue (Y), and column C holds sales volume (bubble size). | Generate a scientific plot that maps concentration (X), response (Y), and particle diameter (size) from three separate data columns. | Automate a monthly performance report that refreshes data and updates a bubble chart with new X, Y, and size ranges without manual re‑configuration.
// AI Prompts: Write C# code with Aspose.Cells to add a bubble chart and bind XValues, Values, and BubbleSizes to ranges A2:A10, B2:B10, and C2:C10. | Explain how to programmatically modify the bubble sizes of an existing Aspose.Cells bubble chart at runtime. | Provide a step‑by‑step guide for configuring multiple series in a bubble chart, each with independent X, Y, and size ranges.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace BubbleChartExample
{
    // Shows how to create a workbook, fill columns A‑C with X, Y, and size data, add a Bubble chart, set the Y‑range as Values, then bind XValues to A2:A5 and BubbleSizes to C2:C5, calculate the layout, and save the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data:
            // Column A – X values, Column B – Y values, Column C – Bubble sizes
            sheet.Cells["A1"].PutValue("X");
            sheet.Cells["B1"].PutValue("Y");
            sheet.Cells["C1"].PutValue("Size");

            for (int i = 2; i <= 5; i++)
            {
                sheet.Cells[$"A{i}"].PutValue(i);           // X = 2,3,4,5
                sheet.Cells[$"B{i}"].PutValue(i * 2);       // Y = 4,6,8,10
                sheet.Cells[$"C{i}"].PutValue(i * 0.5);     // Size = 1,1.5,2,2.5
            }

            // Add a bubble chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Bubble, 6, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Add the series using the Y values range (Values)
            // The Add method also creates the series and sets its Values property
            chart.NSeries.Add("B2:B5", true);

            // Link X values and bubble sizes to separate ranges
            Series series = chart.NSeries[0];
            series.XValues = "A2:A5";        // X‑axis data
            series.BubbleSizes = "C2:C5";    // Bubble size data

            // Optional: calculate the chart layout before saving
            chart.Calculate();

            // Save the workbook with the bubble chart
            workbook.Save("BubbleChartConfigured.xlsx");
        }
    }
}
