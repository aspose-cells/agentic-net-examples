// Title: Measure Aspose.Cells chart rendering time after disabling data label text wrapping (C#)
// Description: C# example that creates a workbook, adds a column chart with data labels, sets DataLabels.IsTextWrapped = false, and uses Stopwatch to time chart.Calculate and workbook.Save, reporting the elapsed milliseconds for rendering and saving.
// Keywords: Aspose.Cells chart performance | disable data label text wrap | chart.Calculate timing | workbook.Save benchmark | C# chart rendering measurement | Aspose.Cells IsTextWrapped | performance profiling Aspose.Cells | chart rendering stopwatch
// Common Searches: how to benchmark Aspose.Cells chart rendering | measure chart save time with Aspose.Cells .NET | impact of IsTextWrapped on chart performance | Aspose.Cells chart calculation speed | C# timing chart rendering Aspose.Cells
// Developer Intent: Find out how long chart calculation and workbook saving take when data label text wrapping is turned off.
// Use Cases: Compare rendering speed of charts with and without text wrapping for large data sets. | Profile chart generation in automated reporting pipelines to meet SLA requirements. | Log per‑chart rendering duration in batch exports to identify performance bottlenecks.
// AI Prompts: Generate C# code that measures rendering time for multiple Aspose.Cells chart types while toggling DataLabels.IsTextWrapped. | Explain how to capture CPU and memory usage during chart.Calculate and workbook.Save in Aspose.Cells. | Suggest optimization techniques to accelerate chart rendering after disabling data label text wrapping.

using System;
using System.Diagnostics;
using Aspose.Cells;
using Aspose.Cells.Charts;

// C# example that creates a workbook, adds a column chart with data labels, sets DataLabels.IsTextWrapped = false, and uses Stopwatch to time chart.Calculate and workbook.Save, reporting the elapsed milliseconds for rendering and saving.
class ChartRenderingTimeMeasurement
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
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Enable data labels and disable text wrapping (using the IsTextWrapped property)
        DataLabels labels = chart.NSeries[0].DataLabels;
        labels.ShowValue = true;
        labels.IsTextWrapped = false;

        // Start timing the rendering process (calculation + saving)
        Stopwatch sw = Stopwatch.StartNew();

        // Force chart layout calculation
        chart.Calculate();

        // Save the workbook (rendering occurs during save)
        workbook.Save("ChartRenderingTime.xlsx", SaveFormat.Xlsx);

        sw.Stop();
        Console.WriteLine($"Chart rendering and saving took {sw.ElapsedMilliseconds} ms");
    }
}
