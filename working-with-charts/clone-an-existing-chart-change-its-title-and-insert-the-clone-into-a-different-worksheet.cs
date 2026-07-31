// Title: Clone a chart, rename it, and place on another worksheet with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, add a source worksheet with sample data and a column chart, duplicate the chart on a different worksheet, change the cloned chart's title, and save the file as an Excel workbook using Aspose.Cells for C#.
// Keywords: Aspose.Cells clone chart | copy chart to another sheet | change chart title Aspose.Cells | SetChartDataRange C# | duplicate Excel chart programmatically | Aspose.Cells chart manipulation
// Common Searches: how to duplicate a chart in Aspose.Cells | clone chart to another worksheet .NET | change title of copied chart Aspose.Cells | set chart data range after cloning | Aspose.Cells copy chart example
// Developer Intent: The developer needs to replicate an existing chart, assign a new title, and insert the copy into a different worksheet using Aspose.Cells for .NET.
// Use Cases: Create a dashboard sheet that aggregates charts from multiple source sheets with custom titles. | Generate department‑specific reports by copying a master chart to each department's worksheet and renaming it. | Programmatically reuse a chart layout for a new data set while preserving its type and series.
// AI Prompts: Write C# code with Aspose.Cells to clone a chart from one worksheet to another and set a new title. | Show how to copy a chart, adjust its data range, and save the workbook using Aspose.Cells. | Explain how to extract the cell range from a chart's full data range string for reuse in a cloned chart.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates how to create a workbook, add a source worksheet with sample data and a column chart, duplicate the chart on a different worksheet, change the cloned chart's title, and save the file as an Excel workbook using Aspose.Cells for C#.
class CloneChartDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // ---------- Source worksheet with original chart ----------
        Worksheet sourceSheet = workbook.Worksheets[0];
        sourceSheet.Name = "SourceSheet";

        // Populate sample data
        sourceSheet.Cells["A1"].PutValue("Category");
        sourceSheet.Cells["A2"].PutValue("A");
        sourceSheet.Cells["A3"].PutValue("B");
        sourceSheet.Cells["A4"].PutValue("C");
        sourceSheet.Cells["B1"].PutValue("Value");
        sourceSheet.Cells["B2"].PutValue(10);
        sourceSheet.Cells["B3"].PutValue(20);
        sourceSheet.Cells["B4"].PutValue(30);

        // Add a chart to the source sheet
        int srcChartIdx = sourceSheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart srcChart = sourceSheet.Charts[srcChartIdx];
        srcChart.NSeries.Add("B2:B4", true);
        srcChart.NSeries.CategoryData = "A2:A4";
        srcChart.Title.Text = "Original Chart";

        // ---------- Target worksheet where the cloned chart will be placed ----------
        Worksheet targetSheet = workbook.Worksheets.Add("TargetSheet");

        // Clone the chart: replicate type and data range
        ChartType chartType = srcChart.Type;
        string fullRange = srcChart.GetChartDataRange(); // e.g., "SourceSheet!A1:B4"

        // Extract only the cell range part (after '!')
        string rangeOnly = fullRange;
        int exclPos = fullRange.IndexOf('!');
        if (exclPos >= 0 && exclPos < fullRange.Length - 1)
            rangeOnly = fullRange.Substring(exclPos + 1);

        // Add a new chart to the target sheet
        int clonedChartIdx = targetSheet.Charts.Add(chartType, 5, 0, 15, 5);
        Chart clonedChart = targetSheet.Charts[clonedChartIdx];

        // Set the same data range for the cloned chart
        clonedChart.SetChartDataRange(rangeOnly, true);

        // Change the title of the cloned chart
        clonedChart.Title.Text = "Cloned Chart Title";

        // Save the workbook
        workbook.Save("ClonedChartDemo.xlsx");
    }
}
