// Title: Measure Aspose.Cells Chart Calculation and Rendering Time After Disabling Data Label Text Wrap (C#)
// Description: Creates a workbook with sample data, adds a column chart, enables data labels, turns off label text wrapping, and uses Stopwatch to record the duration of chart.Calculate() (layout) and Workbook.Save() (rendering). The output shows the elapsed milliseconds for each operation, helping assess the performance impact of disabling text wrap.
// Keywords: Aspose.Cells chart performance | disable data label text wrap | chart.Calculate timing | Workbook.Save benchmark | C# Aspose.Cells rendering speed | chart layout measurement | performance testing Aspose.Cells
// Common Searches: Aspose.Cells measure chart rendering time C# | How to benchmark chart.Calculate performance | Effect of DataLabels.IsTextWrapped on chart speed | Chart rendering latency Aspose.Cells .NET | Timing workbook save with charts
// Developer Intent: The developer wants to benchmark how disabling data label text wrapping influences the time required to calculate a chart layout and to save a workbook containing that chart.
// Use Cases: Compare chart generation speed with and without data label text wrapping to choose the optimal setting. | Profile chart calculation and workbook save times in high‑volume reporting scenarios. | Validate performance gains before deploying large workbooks that contain many charts.
// AI Prompts: Generate C# code that measures chart.Calculate() and Workbook.Save() times for different DataLabels.IsTextWrapped values using Aspose.Cells. | Explain how to analyze the timing results and recommend further optimizations for Aspose.Cells chart rendering. | Create a unit test that asserts chart rendering completes within a defined threshold when text wrapping is disabled.

using System;
using System.Diagnostics;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook with sample data, adds a column chart, enables data labels, turns off label text wrapping, and uses Stopwatch to record the duration of chart.Calculate() (layout) and Workbook.Save() (rendering). The output shows the elapsed milliseconds for each operation, helping assess the performance impact of disabling text wrap.
class ChartRenderingTimeMeasurement
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["A3"].PutValue("Orange");
        sheet.Cells["A4"].PutValue("Banana");
        sheet.Cells["B1"].PutValue("Sales");
        sheet.Cells["B2"].PutValue(120);
        sheet.Cells["B3"].PutValue(85);
        sheet.Cells["B4"].PutValue(65);

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Enable data labels and disable text wrapping for performance test
        DataLabels labels = chart.NSeries[0].DataLabels;
        labels.ShowValue = true;
        labels.IsTextWrapped = false;   // Disable wrapping

        // Measure time taken to calculate the chart layout
        Stopwatch calcTimer = Stopwatch.StartNew();
        chart.Calculate();               // Forces layout calculation
        calcTimer.Stop();

        // Measure time taken to render (save) the workbook containing the chart
        Stopwatch renderTimer = Stopwatch.StartNew();
        workbook.Save("ChartRenderingTime.xlsx", SaveFormat.Xlsx);
        renderTimer.Stop();

        // Output the measured times
        Console.WriteLine($"Chart.Calculate() time: {calcTimer.ElapsedMilliseconds} ms");
        Console.WriteLine($"Workbook.Save() (render) time: {renderTimer.ElapsedMilliseconds} ms");
    }
}
