// Title: Auto‑Resize Radar Chart Data Labels for Long Category Names – Aspose.Cells for .NET
// Description: This C# example creates an Excel workbook, adds a radar chart with lengthy category names, enables data labels that show both the category and its value, sets the label shape to a rectangle, and activates the IsResizeShapeToFitText property so each label automatically expands to fit its text before saving the file.
// Keywords: Aspose.Cells | C# radar chart | data label auto resize | IsResizeShapeToFitText | long category names | Excel chart customization | .NET chart API | radar chart data labels
// Common Searches: Aspose.Cells auto resize radar chart data labels | IsResizeShapeToFitText radar chart C# | fit long category names in radar chart labels | Aspose.Cells set data label shape to rectangle | radar chart label resizing .NET
// Developer Intent: Automatically expand radar chart data label shapes to accommodate long category names.
// Use Cases: Generate radar charts where each label grows to display full category text without truncation. | Produce automated Excel reports with dynamic radar charts that maintain a clean layout despite variable label lengths. | Apply automatic label resizing across multiple series in a radar chart for real‑time data visualizations.
// AI Prompts: Show how to enable automatic resizing of radar chart data label shapes using Aspose.Cells in C#. | Provide C# code that sets IsResizeShapeToFitText for radar chart data labels with long category names. | Explain the impact of IsResizeShapeToFitText on radar chart label rendering and any required chart recalculation steps.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// This C# example creates an Excel workbook, adds a radar chart with lengthy category names, enables data labels that show both the category and its value, sets the label shape to a rectangle, and activates the IsResizeShapeToFitText property so each label automatically expands to fit its text before saving the file.
class RadarChartDataLabelResize
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate worksheet with long category names and some values
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("Very Long Category Name 1");
        worksheet.Cells["A3"].PutValue("Very Long Category Name 2");
        worksheet.Cells["A4"].PutValue("Very Long Category Name 3");

        worksheet.Cells["B1"].PutValue("Series1");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["B4"].PutValue(30);

        // Add a radar chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Radar, 5, 0, 20, 12);
        Chart chart = worksheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.NSeries.Add("B2:B4", true);               // Values
        chart.NSeries.CategoryData = "A2:A4";           // Categories (long names)

        // Enable category axis labels for the radar chart
        chart.NSeries[0].HasRadarAxisLabels = true;

        // Enable data labels and configure them
        Series series = chart.NSeries[0];
        series.DataLabels.ShowCategoryName = true;      // Show the long category names
        series.DataLabels.ShowValue = true;             // Show the values
        series.DataLabels.ShapeType = DataLabelShapeType.Rect; // Optional: set shape type

        // Enable automatic resizing of the data label shape to fit the text
        series.DataLabels.IsResizeShapeToFitText = true;

        // Recalculate the chart layout (optional but ensures proper rendering)
        chart.Calculate();

        // Save the workbook with the radar chart
        workbook.Save("RadarChartAutoResizeDataLabels.xlsx");
    }
}
