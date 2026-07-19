// Title: Exclude “Total” categories from an Aspose.Cells chart using C# AutoFilter
// Description: Shows how to build a workbook, add a column chart, apply an AutoFilter with the NotContains operator on the category column to hide rows whose labels contain “Total”, enable PlotVisibleCellsOnly so the chart reflects only visible data, and save the result.
// Keywords: Aspose.Cells C# chart filter | AutoFilter NotContains | exclude Total rows | PlotVisibleCellsOnly | filter chart categories | Aspose.Cells example | C# Excel chart filtering
// Common Searches: C# Aspose.Cells filter chart by text | How to hide rows with 'Total' in chart using Aspose.Cells | Aspose.Cells NotContains filter example | Plot only visible cells in Aspose chart | Exclude subtotal rows from Excel chart C#
// Developer Intent: Remove categories that contain the word “Total” so the chart displays only the remaining data.
// Use Cases: Create a sales dashboard that omits subtotal rows labeled “Total” before charting. | Generate financial reports where total lines are excluded from visualizations to avoid double‑counting. | Automate monthly Excel reports that automatically filter out any “Total” categories prior to rendering charts. | Apply the same text‑based filter to other columns and refresh linked charts for clean data presentation.
// AI Prompts: Write C# code with Aspose.Cells to filter out chart categories containing 'Total' and ensure the chart plots only visible cells. | Explain how to extend the example to filter multiple keywords such as 'Total' and 'Subtotal' in chart categories. | Show how to apply a NotContains AutoFilter on a different column and update the associated chart in Aspose.Cells. | Provide a step‑by‑step guide for using AutoFilter with PlotVisibleCellsOnly to create clean charts in .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Tables; // for FilterOperatorType

// Shows how to build a workbook, add a column chart, apply an AutoFilter with the NotContains operator on the category column to hide rows whose labels contain “Total”, enable PlotVisibleCellsOnly so the chart reflects only visible data, and save the result.
class FilterChartCategories
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data with some categories containing the word "Total"
        // Column A – Category, Column B – Value
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["A2"].PutValue("North");
        sheet.Cells["B2"].PutValue(120);
        sheet.Cells["A3"].PutValue("South Total");
        sheet.Cells["B3"].PutValue(150);
        sheet.Cells["A4"].PutValue("East");
        sheet.Cells["B4"].PutValue(200);
        sheet.Cells["A5"].PutValue("West Total");
        sheet.Cells["B5"].PutValue(180);
        sheet.Cells["A6"].PutValue("Central");
        sheet.Cells["B6"].PutValue(130);

        // Add a column chart using the data range
        int chartIdx = sheet.Charts.Add(ChartType.Column, 8, 0, 20, 8);
        Chart chart = sheet.Charts[chartIdx];
        chart.NSeries.Add("B2:B6", true);          // Values
        chart.NSeries.CategoryData = "A2:A6";      // Categories

        // Apply an AutoFilter on the Category column (field index 0)
        // Exclude rows where the category contains the word "Total"
        // Use the Custom filter with the NotContains operator
        sheet.AutoFilter.SetRange(0, 0, 5); // Header row 0, column 0, last row 5
        sheet.AutoFilter.Custom(0, FilterOperatorType.NotContains, "Total");
        sheet.AutoFilter.Refresh();

        // Instruct the chart to plot only visible cells (i.e., filtered rows)
        chart.PlotVisibleCellsOnly = true;

        // Save the workbook
        workbook.Save("FilteredChart.xlsx");
    }
}
