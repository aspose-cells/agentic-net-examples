// Title: Aspose.Cells C# – Export Column Chart to PNG with Data Labels Visible
// Description: Demonstrates how to create a column chart in a workbook, enable value and category data labels, position them outside the columns, calculate the layout, and export the chart to a PNG image while preserving label visibility. The workbook is also saved for future edits.
// Keywords: Aspose.Cells export chart PNG | C# chart ToImage data labels | preserve data labels Aspose.Cells | column chart PNG export .NET | Aspose.Cells ChartType.Column example | DataLabels.ShowValue C# | DataLabels.ShowCategoryName | Aspose.Cells chart image generation | export chart with labels
// Common Searches: export Aspose.Cells chart to PNG with labels | keep data labels visible when saving chart as image | Aspose.Cells C# chart ToImage preserving labels | how to show category names on exported chart image | Aspose.Cells column chart PNG export tutorial
// Developer Intent: Export a chart as a PNG image while ensuring all data labels remain visible.
// Use Cases: Create a sales chart for dashboards and deliver a ready‑to‑use PNG that includes values and categories. | Generate chart images for web reports where the source workbook stays editable. | Automate batch processing of multiple charts, exporting each to PNG with full label information.
// AI Prompts: Write C# code using Aspose.Cells to build a line chart, enable data labels, and export it to a high‑resolution PNG while keeping the labels visible. | Explain how to customize data label font, color, and position before calling Chart.ToImage in Aspose.Cells. | Provide a script that iterates through all charts in an Aspose.Cells workbook and saves each as a PNG with data labels preserved.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Demonstrates how to create a column chart in a workbook, enable value and category data labels, position them outside the columns, calculate the layout, and export the chart to a PNG image while preserving label visibility. The workbook is also saved for future edits.
class PreserveDataLabelsWhenExportingChart
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["A3"].PutValue("Orange");
        sheet.Cells["A4"].PutValue("Banana");

        sheet.Cells["B1"].PutValue("Sales");
        sheet.Cells["B2"].PutValue(1200);
        sheet.Cells["B3"].PutValue(800);
        sheet.Cells["B4"].PutValue(1500);

        // Add a column chart
        int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIdx];

        // Set the data range for the chart
        chart.SetChartDataRange("A1:B4", true);

        // Enable data labels to show the values on each column
        Series series = chart.NSeries[0];
        series.DataLabels.ShowValue = true;          // show numeric values
        series.DataLabels.ShowCategoryName = true;   // optional: show category names
        series.DataLabels.Position = LabelPositionType.OutsideEnd; // place labels outside

        // Ensure the chart layout is calculated before rendering
        chart.Calculate();

        // Export the chart to a PNG image while preserving the data labels
        chart.ToImage("ChartWithDataLabels.png", ImageType.Png);

        // Optionally save the workbook for reference
        workbook.Save("ChartWithDataLabels.xlsx");
    }
}
