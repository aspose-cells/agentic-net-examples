// Title: C# – Hide Horizontal Axis Gridlines in an Aspose.Cells Scatter Chart
// Description: Creates a workbook, adds X/Y data, inserts a scatter chart, and disables both major and minor gridlines on the horizontal (category) axis before saving the file as an Excel workbook.
// Keywords: Aspose.Cells hide horizontal gridlines | C# scatter chart gridlines | CategoryAxis visibility Aspose.Cells | remove chart gridlines .NET | Aspose.Cells chart formatting
// Common Searches: Aspose.Cells hide horizontal axis gridlines | C# scatter chart without gridlines | disable category axis gridlines Aspose.Cells | remove major and minor gridlines from chart | Aspose.Cells chart appearance settings
// Developer Intent: Turn off the horizontal (category) axis gridlines in a scatter chart generated with Aspose.Cells for .NET.
// Use Cases: Produce clean‑look scatter charts for dashboards where gridlines distract from data trends. | Automate Excel report generation with minimal visual clutter by hiding category axis lines. | Customize chart styling in batch processing scripts, keeping only vertical gridlines visible.
// AI Prompts: Generate C# code that hides major and minor horizontal gridlines for any Aspose.Cells chart. | Write a reusable method to toggle CategoryAxis gridline visibility based on a boolean flag. | Explain how to keep vertical axis gridlines while removing only the horizontal ones in an Aspose.Cells chart.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds X/Y data, inserts a scatter chart, and disables both major and minor gridlines on the horizontal (category) axis before saving the file as an Excel workbook.
class HideHorizontalGridlines
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for a scatter plot (X values in column A, Y values in column B)
        sheet.Cells["A1"].PutValue("X");
        sheet.Cells["B1"].PutValue("Y");
        sheet.Cells["A2"].PutValue(1);
        sheet.Cells["B2"].PutValue(2);
        sheet.Cells["A3"].PutValue(2);
        sheet.Cells["B3"].PutValue(4);
        sheet.Cells["A4"].PutValue(3);
        sheet.Cells["B4"].PutValue(6);

        // Add a scatter chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Scatter, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the series (Y values) and the category data (X values)
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Hide the horizontal axis (category axis) gridlines for a cleaner appearance
        chart.CategoryAxis.MajorGridLines.IsVisible = false;
        chart.CategoryAxis.MinorGridLines.IsVisible = false;

        // Save the workbook with the modified chart
        workbook.Save("Scatter_No_Horizontal_Gridlines.xlsx");
    }
}
