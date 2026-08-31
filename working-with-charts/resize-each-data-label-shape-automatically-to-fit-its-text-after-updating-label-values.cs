// Title: How to auto‑fit data label shapes to custom text in an Aspose.Cells column chart using C#
// AI Prompts: Write C# code that iterates over each point in an Aspose.Cells chart series, assigns a custom label string, and turns on automatic shape resizing so the label fits its text. | Show how to create a column chart with Aspose.Cells, display data labels, and make the label boxes automatically adjust to the label content. | Provide a complete example that updates data label values with a prefix, enables auto‑fit for the label shapes, and saves the workbook.
// Common Searches: Aspose.Cells C# auto resize data label shape after setting custom label text | How to turn on automatic resizing of chart data label shapes in Aspose.Cells | C# example to change data label text and make label boxes fit content in an Excel chart using Aspose.Cells | Resize data labels to fit content in an Aspose.Cells column chart with C#
// Tags: auto‑fit chart data labels Aspose.Cells C# | custom data label text Aspose.Cells | IsResizeShapeToFitText property Aspose.Cells | column chart label shape resizing Aspose.Cells | chart point label customization Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a workbook, adds a column chart with sample data, enables data labels for the first series, assigns custom text to each point's label, activates automatic resizing of the label shapes so they fit the new text, and saves the workbook as an XLSX file.
class ResizeDataLabelsDemo
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
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Enable data labels for the first series
        Series series = chart.NSeries[0];
        series.DataLabels.ShowValue = true;

        // Update each data label's text and enable auto‑fit to the text
        foreach (ChartPoint point in series.Points)
        {
            // Set a custom label text (example: include a prefix)
            point.DataLabels.Text = $"Value: {point.YValue}";

            // Turn on automatic resizing so the shape fits the new text
            point.DataLabels.IsResizeShapeToFitText = true;
        }

        // Save the workbook
        workbook.Save("ResizeDataLabelsDemo.xlsx");
    }
}
