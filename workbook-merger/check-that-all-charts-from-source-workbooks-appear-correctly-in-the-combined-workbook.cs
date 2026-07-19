// Title: Verify All Charts Preserve Data When Merging Workbooks with Aspose.Cells for .NET
// Description: This C# example creates two source workbooks, each with a different chart type, merges them into a destination workbook using `Workbook.Combine`, refreshes worksheets, counts expected vs. actual charts, recalculates each chart, checks the `IsChartDataChanged` flag, and saves the combined file. It demonstrates how to ensure charts remain intact and correctly linked after a merge.
// Keywords: Aspose.Cells combine workbooks C# | merge Excel files preserve charts | verify chart count after combine | IsChartDataChanged Aspose.Cells | chart data integrity workbook merge | C# Aspose.Cells chart validation | Workbook.Combine charts
// Common Searches: how to keep charts when merging Excel workbooks with Aspose.Cells | check chart data source after workbook combine C# | count charts in merged workbook Aspose.Cells | validate chart integrity after combining workbooks | Aspose.Cells chart data changed flag
// Developer Intent: The developer needs to confirm that every chart from the source workbooks appears in the merged workbook and that its data references stay unchanged after using the `Combine` method.
// Use Cases: Calculate the expected number of charts from all source worksheets and compare it with the actual count after `Workbook.Combine`. | Iterate through each worksheet in the merged workbook, call `Chart.Calculate()`, and log `Chart.IsChartDataChanged` to detect any broken data links. | Save the merged workbook only after all charts have been verified to ensure the final file contains the original visualizations.
// AI Prompts: Generate C# code that merges multiple Excel workbooks with Aspose.Cells and validates that each chart's data source remains unchanged. | Provide a method to log each chart's name and its `IsChartDataChanged` status after a workbook combine operation. | Explain strategies for handling chart name conflicts when combining workbooks that contain charts with identical names.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// This C# example creates two source workbooks, each with a different chart type, merges them into a destination workbook using `Workbook.Combine`, refreshes worksheets, counts expected vs. actual charts, recalculates each chart, checks the `IsChartDataChanged` flag, and saves the combined file. It demonstrates how to ensure charts remain intact and correctly linked after a merge.
class CheckCombinedCharts
{
    static void Main()
    {
        // ---------- Create first source workbook with a chart ----------
        Workbook source1 = new Workbook();
        Worksheet ws1 = source1.Worksheets[0];
        ws1.Name = "Source1";
        ws1.Cells["A1"].PutValue("Category");
        ws1.Cells["A2"].PutValue("A");
        ws1.Cells["A3"].PutValue("B");
        ws1.Cells["B1"].PutValue("Value");
        ws1.Cells["B2"].PutValue(10);
        ws1.Cells["B3"].PutValue(20);
        int chartIdx1 = ws1.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart1 = ws1.Charts[chartIdx1];
        chart1.NSeries.Add("B2:B3", true);
        chart1.NSeries.CategoryData = "A2:A3";

        // ---------- Create second source workbook with a chart ----------
        Workbook source2 = new Workbook();
        Worksheet ws2 = source2.Worksheets[0];
        ws2.Name = "Source2";
        ws2.Cells["A1"].PutValue("Month");
        ws2.Cells["A2"].PutValue("Jan");
        ws2.Cells["A3"].PutValue("Feb");
        ws2.Cells["B1"].PutValue("Sales");
        ws2.Cells["B2"].PutValue(150);
        ws2.Cells["B3"].PutValue(200);
        int chartIdx2 = ws2.Charts.Add(ChartType.Pie, 5, 0, 15, 5);
        Chart chart2 = ws2.Charts[chartIdx2];
        chart2.NSeries.Add("B2:B3", true);
        chart2.NSeries.CategoryData = "A2:A3";

        // ---------- Destination workbook ----------
        Workbook dest = new Workbook();

        // Expected total number of charts from all sources
        int expectedChartCount = ws1.Charts.Count + ws2.Charts.Count;

        // ---------- Combine source workbooks ----------
        dest.Combine(source1);
        dest.Combine(source2);

        // Refresh any pivot tables/charts (good practice)
        dest.Worksheets.RefreshAll();

        // ---------- Verify chart count ----------
        int actualChartCount = 0;
        foreach (Worksheet ws in dest.Worksheets)
        {
            actualChartCount += ws.Charts.Count;
        }

        Console.WriteLine($"Expected chart count: {expectedChartCount}");
        Console.WriteLine($"Actual chart count after combine: {actualChartCount}");

        // ---------- Detailed verification of each chart ----------
        foreach (Worksheet ws in dest.Worksheets)
        {
            foreach (Chart ch in ws.Charts)
            {
                // Force calculation to update internal data state
                ch.Calculate();

                // Check if chart data source has changed
                bool dataChanged = ch.IsChartDataChanged();
                Console.WriteLine($"Chart '{ch.Name}' in sheet '{ws.Name}' data changed flag: {dataChanged}");
            }
        }

        // ---------- Save the combined workbook ----------
        dest.Save("CombinedWithCharts.xlsx", SaveFormat.Xlsx);
    }
}
