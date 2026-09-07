// Title: Benchmark Aspose.Cells chart layout and file save time after turning off data label text wrapping in C#
// AI Prompts: Write C# code that creates a column chart with Aspose.Cells, disables DataLabels.IsTextWrapped, and measures the execution time of Chart.Calculate and Workbook.Save using Stopwatch. | Show how to profile the performance impact of turning off data label text wrapping on an Aspose.Cells chart by timing the layout calculation and workbook export.
// Common Searches: how to measure Aspose.Cells chart.Calculate execution time in .NET | benchmark workbook.save performance after modifying chart data labels in C# | does disabling DataLabels.IsTextWrapped improve chart rendering speed in Aspose.Cells | C# example for timing chart layout calculation with Aspose.Cells
// Tags: Aspose.Cells chart layout timing | Aspose.Cells DataLabels.IsTextWrapped impact | C# benchmark workbook.Save latency | Aspose.Cells chart rendering speed measurement | disable data label wrapping Aspose chart

using System;
using System.Diagnostics;
using Aspose.Cells;
using Aspose.Cells.Charts;

// // Demonstrates creating a workbook with a column chart, turning off data label text wrapping, and using Stopwatch to time Chart.Calculate and Workbook.Save operations for performance analysis.
class ChartRenderingTimeDemo
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

        // Add a column chart to the worksheet
        int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIdx];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Enable data labels and disable text wrapping (performance tweak)
        DataLabels labels = chart.NSeries[0].DataLabels;
        labels.ShowValue = true;
        labels.IsTextWrapped = false; // using DataLabels.IsTextWrapped property

        // Measure time taken to calculate the chart layout
        Stopwatch sw = new Stopwatch();
        sw.Start();
        chart.Calculate(); // using Chart.Calculate()
        sw.Stop();
        Console.WriteLine($"Chart.Calculate() elapsed time: {sw.ElapsedMilliseconds} ms");

        // Measure time taken to save the workbook (rendering)
        sw.Restart();
        workbook.Save("ChartRenderingTime.xlsx");
        sw.Stop();
        Console.WriteLine($"Workbook.Save() elapsed time: {sw.ElapsedMilliseconds} ms");
    }
}
