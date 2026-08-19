// Title: Aspose.Cells .NET: Configure XValues, Values, and BubbleSizes for a Bubble Chart
// Description: Shows how to create a workbook, fill columns A‑C with X, Y, and bubble‑size data, add a Bubble chart, and bind the series to three separate ranges (A2:A5, B2:B5, C2:C5) before saving the file.
// Keywords: Aspose.Cells | bubble chart | XValues | BubbleSizes | .NET | chart data binding | series range | Excel automation | C#
// Common Searches: Aspose.Cells bind bubble chart XValues | set bubble sizes from cells Aspose.Cells | C# bubble chart series range | link X and Y values in Aspose.Cells chart | Aspose.Cells chart data source example
// Developer Intent: Link a bubble chart’s XValues, Y values, and bubble sizes to three distinct cell ranges using Aspose.Cells for .NET.
// Use Cases: Create a bubble chart where X axis, Y axis, and bubble diameter are driven by data in separate columns. | Build a reusable routine that adds a bubble chart to any worksheet and assigns custom data ranges for dynamic reporting. | Export workbooks with pre‑configured bubble charts for dashboards or presentation decks.
// AI Prompts: Generate C# code to add multiple bubble series with independent X, Y, and size ranges using Aspose.Cells. | Provide a method that updates the XValues, Values, and BubbleSizes of an existing chart after the worksheet data changes. | Explain how to format bubble colors and add data labels programmatically with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace BubbleChartExample
{
    // Shows how to create a workbook, fill columns A‑C with X, Y, and bubble‑size data, add a Bubble chart, and bind the series to three separate ranges (A2:A5, B2:B5, C2:C5) before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for X values, Y values, and bubble sizes
            sheet.Cells["A1"].PutValue("X Values");
            sheet.Cells["B1"].PutValue("Y Values");
            sheet.Cells["C1"].PutValue("Bubble Sizes");

            for (int i = 2; i <= 5; i++)
            {
                sheet.Cells[$"A{i}"].PutValue(i);               // X value
                sheet.Cells[$"B{i}"].PutValue(i * 2);           // Y value
                sheet.Cells[$"C{i}"].PutValue(i * 0.5);         // Bubble size
            }

            // Add a bubble chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Bubble, 6, 0, 22, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Add the Y values (Values) series – this also creates the series object
            int seriesIndex = chart.NSeries.Add("B2:B5", true);
            Series series = chart.NSeries[seriesIndex];

            // Link X values and bubble sizes to separate ranges
            series.XValues = "A2:A5";
            series.BubbleSizes = "C2:C5";

            // Optional: set a title for clarity
            chart.Title.Text = "Sample Bubble Chart";

            // Save the workbook
            workbook.Save("BubbleChartConfigured.xlsx");
        }
    }
}
