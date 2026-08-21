// Title: Right‑Align Data Labels in a Horizontal Bar Chart – Aspose.Cells for .NET (C#) Example
// Description: This C# sample creates a workbook, adds category and value data, inserts a horizontal bar chart, enables data labels for the first series, and sets the TextHorizontalAlignment property to Right so the labels are right‑justified. The workbook is saved as an XLSX file.
// Keywords: Aspose.Cells | .NET | C# | horizontal bar chart | chart data labels | right align | TextHorizontalAlignment | Excel chart formatting | sample code | GitHub example
// Common Searches: Aspose.Cells right align data labels | C# set TextHorizontalAlignment for chart series | horizontal bar chart label alignment Aspose.Cells | how to justify chart data labels to the right | Aspose.Cells chart formatting examples
// Developer Intent: Set the data‑label text of a horizontal bar chart to right‑justified using Aspose.Cells for .NET.
// Use Cases: Generate Excel reports where bar‑chart labels need to line up with right‑aligned columns. | Improve readability of horizontal bar charts by positioning values at the far end of each bar. | Create reusable chart‑formatting utilities that apply right‑aligned data labels across multiple series.
// AI Prompts: Show C# code to right‑align data labels in a horizontal bar chart with Aspose.Cells. | How do I apply TextHorizontalAlignment.Right to all series in an Aspose.Cells chart? | Explain the effect of TextHorizontalAlignment on data label placement in a bar chart using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// This C# sample creates a workbook, adds category and value data, inserts a horizontal bar chart, enables data labels for the first series, and sets the TextHorizontalAlignment property to Right so the labels are right‑justified. The workbook is saved as an XLSX file.
class SetDataLabelAlignment
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the horizontal bar chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");

        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        // Add a horizontal bar chart (ChartType.Bar) to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Bar, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.NSeries.Add("B2:B4", true);          // Values
        chart.NSeries.CategoryData = "A2:A4";      // Categories

        // Enable data labels for the first series
        chart.NSeries[0].DataLabels.ShowValue = true;

        // Right‑justify the data label text
        chart.NSeries[0].DataLabels.TextHorizontalAlignment = TextAlignmentType.Right;

        // Save the workbook
        workbook.Save("HorizontalBarChart_WithRightAlignedDataLabels.xlsx");
    }
}
