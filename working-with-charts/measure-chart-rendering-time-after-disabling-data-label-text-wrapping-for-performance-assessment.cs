// Title: Benchmark Aspose.Cells chart rendering time after disabling data label wrapping (C#)
// Description: C# example that creates a workbook, adds a column chart with data labels, turns off text wrapping, and measures the Chart.Calculate execution time using a Stopwatch. The elapsed milliseconds are printed and the workbook is saved.
// Keywords: Aspose.Cells chart performance | disable data label wrapping | Chart.Calculate timing | measure chart rendering speed | .NET spreadsheet rendering benchmark | C# Aspose.Cells performance test | chart calculation time measurement | data label wrap impact | Aspose.Cells chart optimization | stopwatch chart rendering
// Common Searches: how to time Aspose.Cells chart rendering in C# | does disabling data label wrap improve chart speed | benchmark Chart.Calculate method Aspose.Cells | measure spreadsheet chart performance .NET | Aspose.Cells chart rendering latency
// Developer Intent: The developer wants to evaluate the performance gain of turning off data label text wrapping by timing chart calculation.
// Use Cases: Compare rendering speed of charts with and without data label wrapping to choose the most efficient setting. | Integrate chart timing into CI pipelines for regression testing of spreadsheet generation performance. | Select the fastest chart type for large reports by measuring calculation time under different label configurations.
// AI Prompts: Generate C# code that measures Aspose.Cells chart rendering time for both wrapped and unwrapped data labels and outputs the results side‑by‑side. | Explain how Chart.Calculate influences rendering performance and suggest best practices to minimize execution time. | Create a .NET unit test that asserts the chart calculation completes within a given threshold when data label wrapping is disabled.

using System;
using System.Diagnostics;
using Aspose.Cells;
using Aspose.Cells.Charts;

// C# example that creates a workbook, adds a column chart with data labels, turns off text wrapping, and measures the Chart.Calculate execution time using a Stopwatch. The elapsed milliseconds are printed and the workbook is saved.
class ChartRenderTiming
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Fill sample data for the chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["A4"].PutValue("C");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["B4"].PutValue(30);

        // Add a column chart
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        Chart chart = worksheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Enable data labels and disable text wrapping (performance test)
        DataLabels dataLabels = chart.NSeries[0].DataLabels;
        dataLabels.ShowValue = true;
        dataLabels.IsTextWrapped = false; // turn off wrapping

        // Measure the time taken to calculate (render) the chart
        Stopwatch stopwatch = new Stopwatch();
        stopwatch.Start();
        chart.Calculate(); // forces layout and rendering calculations
        stopwatch.Stop();

        Console.WriteLine($"Chart calculation time: {stopwatch.ElapsedMilliseconds} ms");

        // Save the workbook
        workbook.Save("ChartRenderTiming.xlsx");
    }
}
