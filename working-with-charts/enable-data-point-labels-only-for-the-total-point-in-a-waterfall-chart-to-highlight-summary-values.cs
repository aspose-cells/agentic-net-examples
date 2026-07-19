// Title: Show a Data Label Only on the Total Bar of a Waterfall Chart with Aspose.Cells (C#)
// Description: Creates a workbook, inserts a Waterfall chart, marks the final bar as a subtotal (total) using the Subtotals property, and configures DataLabels so that only the total point shows a value. The label is positioned outside the bar and uses Excel's auto‑generated text before saving the file.
// Keywords: Aspose.Cells | C# | waterfall chart | total label | data labels | subtotal point | chart customization | Excel automation
// Common Searches: Aspose.Cells display label on waterfall chart total | C# add data label to specific point in waterfall chart | set subtotal index Aspose.Cells | waterfall chart label only for total bar | Aspose.Cells chart data labels per point
// Developer Intent: Add a waterfall chart and show a data label only on the total (subtotal) bar.
// Use Cases: Financial reports where only the final total column needs a visible value for quick summary. | Presentation‑ready Excel workbooks that highlight the summary bar while keeping intermediate bars label‑free. | Automated chart generation that positions the total label outside the bar for enhanced readability.
// AI Prompts: Generate C# code with Aspose.Cells that creates a waterfall chart and shows a data label only on the total bar. | Explain how to use the Subtotals property and per‑point DataLabels to display a label for a specific waterfall chart point in Aspose.Cells. | Provide steps to customize the label position and enable auto text for the total point of a waterfall chart using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, inserts a Waterfall chart, marks the final bar as a subtotal (total) using the Subtotals property, and configures DataLabels so that only the total point shows a value. The label is positioned outside the bar and uses Excel's auto‑generated text before saving the file.
class WaterfallTotalLabelDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for a waterfall chart
        // Column A – Categories, Column B – Values
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Start");
        sheet.Cells["A3"].PutValue("Increase");
        sheet.Cells["A4"].PutValue("Decrease");
        sheet.Cells["A5"].PutValue("Total");

        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(100);   // Start
        sheet.Cells["B3"].PutValue(40);    // Increase
        sheet.Cells["B4"].PutValue(-20);   // Decrease
        sheet.Cells["B5"].PutValue(120);   // Total (summary)

        // Add a Waterfall chart
        int chartIndex = sheet.Charts.Add(ChartType.Waterfall, 7, 0, 25, 15);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the series
        chart.NSeries.Add("B2:B5", true);
        chart.NSeries.CategoryData = "A2:A5";

        // Identify the total point (index 3 in zero‑based collection, i.e., the 4th point)
        // Waterfall treats total points as subtotals.
        Series series = chart.NSeries[0];
        series.LayoutProperties.Subtotals = new int[] { 3 };

        // Enable data labels only for the total point
        for (int i = 0; i < series.Points.Count; i++)
        {
            ChartPoint point = series.Points[i];
            // Show value only for subtotal (total) points
            point.DataLabels.ShowValue = (i == 3);
            // Optional: customize appearance of the total label
            if (i == 3)
            {
                point.DataLabels.Position = LabelPositionType.OutsideEnd;
                point.DataLabels.IsAutoText = true; // let Excel generate the label text
            }
        }

        // Calculate the chart to ensure data points are up‑to‑date
        chart.Calculate();

        // Save the workbook
        workbook.Save("WaterfallTotalLabelDemo.xlsx");
    }
}
