// Title: Show Only the Total Value Label in an Aspose.Cells Waterfall Chart (C#)
// Description: This Aspose.Cells for .NET example creates a workbook, adds a Waterfall chart with start, increase, decrease and total categories, marks the total bar as a subtotal, forces chart calculation, and displays a data label exclusively on the total (summary) point with custom text and positioning before saving the file.
// Keywords: Aspose.Cells | Aspose.Cells Waterfall chart | C# Waterfall chart total label | show data label for subtotal point | chart subtotal Aspose.Cells | highlight total bar | Waterfall chart data labels | .NET chart customization | ChartPoint.DataLabels | Subtotals property
// Common Searches: Aspose.Cells display only total label in waterfall chart C# | how to set data label for subtotal point Aspose.Cells | waterfall chart show summary value label .NET | custom text for total bar Aspose.Cells | hide all data labels except total in waterfall chart
// Developer Intent: Add a data label that appears only on the total (summary) point of a Waterfall chart.
// Use Cases: Financial reports where only the final total bar needs to be labeled for a clean visual. | Executive dashboards that highlight the summary value in a waterfall chart with custom text such as "Summary". | Automated workbook generation that programmatically formats the total point label while suppressing labels on all other bars.
// AI Prompts: Generate C# code using Aspose.Cells to create a Waterfall chart and show a data label only on the total point, including custom label text and positioning. | Explain how the Subtotals property and ChartPoint.DataLabels are used to enable a label for a specific point in an Aspose.Cells Waterfall chart. | Provide step‑by‑step instructions to format the total point label (show value, set position above, add custom text) while keeping other point labels hidden.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// This Aspose.Cells for .NET example creates a workbook, adds a Waterfall chart with start, increase, decrease and total categories, marks the total bar as a subtotal, forces chart calculation, and displays a data label exclusively on the total (summary) point with custom text and positioning before saving the file.
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
        sheet.Cells["B3"].PutValue(30);    // Increase
        sheet.Cells["B4"].PutValue(-20);   // Decrease
        sheet.Cells["B5"].PutValue(0);     // Total (will be calculated by the chart)

        // Add a Waterfall chart
        int chartIndex = sheet.Charts.Add(ChartType.Waterfall, 7, 0, 25, 15);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the series
        chart.NSeries.Add("B2:B5", true);
        chart.NSeries.CategoryData = "A2:A5";

        // Define the index of the total (summary) point.
        // In this example the total is the last point (index 3, zero‑based).
        chart.NSeries[0].LayoutProperties.Subtotals = new int[] { 3 };

        // Force the chart to calculate so that the total point value is generated
        chart.Calculate();

        // Enable data label only for the total point
        ChartPoint totalPoint = chart.NSeries[0].Points[3];
        totalPoint.DataLabels.ShowValue = true;               // show the calculated total value
        totalPoint.DataLabels.Position = LabelPositionType.Above; // place label above the bar
        totalPoint.DataLabels.IsAutoText = false;             // allow custom formatting if needed
        totalPoint.DataLabels.Text = "Summary";               // optional custom text

        // Save the workbook
        workbook.Save("WaterfallTotalLabelDemo.xlsx");
    }
}
