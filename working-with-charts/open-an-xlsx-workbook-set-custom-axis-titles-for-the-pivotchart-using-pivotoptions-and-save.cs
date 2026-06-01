using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Pivot;

class PivotChartAxisTitleExample
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Assume a pivot table named "PivotTable1" already exists in the worksheet.
        // If not, you would create it here before linking the chart.

        // Add a new PivotChart (or get an existing one)
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 10);
        Chart chart = worksheet.Charts[chartIndex];

        // Link the chart to the pivot table
        chart.PivotSource = "PivotTable1";

        // Configure PivotOptions (example: make pivot controls visible)
        PivotOptions pivotOptions = chart.PivotOptions;
        pivotOptions.DropZonesVisible = true;          // show pivot drop zones
        pivotOptions.ShowExpandCollapseFieldButtons = true; // optional visual setting

        // Set custom axis titles for the PivotChart
        // Category (X) axis title
        chart.CategoryAxis.Title.Text = "Custom Category Axis";
        chart.CategoryAxis.Title.IsVisible = true;

        // Value (Y) axis title
        chart.ValueAxis.Title.Text = "Custom Value Axis";
        chart.ValueAxis.Title.IsVisible = true;

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}