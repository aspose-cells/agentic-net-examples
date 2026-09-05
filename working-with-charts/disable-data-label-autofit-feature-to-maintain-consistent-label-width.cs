// Title: How to disable data label auto‑fit and set a fixed pixel width for chart labels in Aspose.Cells (C#)
// AI Prompts: Provide C# source that creates a column chart with Aspose.Cells, turns off the auto‑fit of each data label, and sets each label to 60 pixels wide. | Show an Aspose.Cells .NET snippet that adds data labels to a series, disables shape resizing, and applies a uniform label width. | Generate a complete workbook example where chart data labels keep a constant size by using IsResizeShapeToFitText = false and WidthPixel.
// Common Searches: how to stop chart data labels from resizing automatically in Aspose.Cells C# | setting a specific pixel width for Excel chart labels with Aspose.Cells | Aspose.Cells example for fixed-size data labels in a column chart | preventing variable label size in Aspose.Cells generated charts | C# Aspose.Cells chart label width configuration
// Tags: Aspose.Cells chart data label auto‑fit off | Aspose.Cells fixed label width C# | C# chart data label size control | column chart label width Aspose.Cells | IsResizeShapeToFitText false usage

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds sample data, inserts a column chart, enables data labels, disables automatic resizing for each label, assigns a constant pixel width, and saves the file as an .xlsx workbook.
class DisableDataLabelAutoFit
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["A4"].PutValue("C");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["B4"].PutValue(30);

        // Add a column chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = worksheet.Charts[chartIndex];

        // Set the data range for the series and categories
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Enable data labels for the first series
        Series series = chart.NSeries[0];
        series.DataLabels.ShowValue = true;

        // Disable auto‑fit (auto‑resize) for each data label to keep a consistent width
        foreach (ChartPoint point in series.Points)
        {
            // Prevent the label shape from resizing to fit its text
            point.DataLabels.IsResizeShapeToFitText = false;

            // Optionally set a fixed width (in pixels) for all labels
            point.DataLabels.WidthPixel = 60;
        }

        // Save the workbook to a file
        workbook.Save("DisableDataLabelAutoFit.xlsx");
    }
}
