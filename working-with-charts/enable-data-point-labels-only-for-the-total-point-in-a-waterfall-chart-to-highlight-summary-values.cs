// Title: Show a Data Label Only on the Total Bar of a Waterfall Chart with Aspose.Cells (C#)
// Description: Creates a workbook, inserts sample data, builds a Waterfall chart, marks the final point as a subtotal via LayoutProperties.Subtotals, calculates the chart, and enables a data label exclusively for that total point using ChartPoint.DataLabels (ShowValue, Position, IsAutoText). Saves the result as an Excel file.
// Keywords: Aspose.Cells | C# | Waterfall chart | total bar label | chart point data label | LayoutProperties.Subtotals | ChartPoint.DataLabels | Excel automation | financial waterfall | summary value label
// Common Searches: Aspose.Cells show label only on waterfall total bar | C# add data label to specific point in waterfall chart | how to label subtotal point in Aspose.Cells chart | display total value label in Excel waterfall using Aspose | set data label for one point Aspose.Cells C#
// Developer Intent: Add a single data label to the total column of a waterfall chart using Aspose.Cells in C#.
// Use Cases: Emphasize the final total in a financial waterfall chart while keeping intermediate bars unlabeled. | Produce a clean budget‑tracking report where only the overall cost column displays its value. | Create a project‑status waterfall diagram that highlights the cumulative result without cluttering the view.
// AI Prompts: Generate C# code with Aspose.Cells that adds a waterfall chart and shows a data label only on the total bar. | Explain how LayoutProperties.Subtotals and ChartPoint.DataLabels work together to label a single point in a waterfall chart. | Provide step‑by‑step instructions for enabling a data label exclusively for the subtotal point in an Aspose.Cells waterfall chart.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, inserts sample data, builds a Waterfall chart, marks the final point as a subtotal via LayoutProperties.Subtotals, calculates the chart, and enables a data label exclusively for that total point using ChartPoint.DataLabels (ShowValue, Position, IsAutoText). Saves the result as an Excel file.
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
        sheet.Cells["B5"].PutValue(0);     // Placeholder for total (will be calculated)

        // Add a Waterfall chart
        int chartIndex = sheet.Charts.Add(ChartType.Waterfall, 7, 0, 25, 15);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the series
        chart.NSeries.Add("B2:B5", true);
        chart.NSeries.CategoryData = "A2:A5";

        // Mark the last point as a total (subtotal) point
        // In a waterfall chart, subtotal points are defined via LayoutProperties.Subtotals
        chart.NSeries[0].LayoutProperties.Subtotals = new int[] { 4 }; // zero‑based index, point 4 is the total

        // Calculate the chart so that the total point value is generated
        chart.Calculate();

        // Enable data label only for the total point
        // The total point is at index 4 (the fifth point)
        ChartPoint totalPoint = chart.NSeries[0].Points[4];
        totalPoint.DataLabels.ShowValue = true;               // show the calculated total value
        totalPoint.DataLabels.Position = LabelPositionType.Above; // optional: position the label above the bar
        totalPoint.DataLabels.IsAutoText = true;              // let the label use the auto‑generated text

        // Save the workbook
        workbook.Save("WaterfallTotalLabelDemo.xlsx");
    }
}
