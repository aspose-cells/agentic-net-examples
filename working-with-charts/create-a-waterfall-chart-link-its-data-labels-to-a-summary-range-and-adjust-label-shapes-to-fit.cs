// Title: Aspose.Cells .NET: Create Waterfall Chart with Linked Data Labels and Auto‑Fit Shapes
// Description: Shows how to generate a waterfall chart in a new workbook, link its data‑label text to a summary range (C2:C5), display values, set a rectangular label shape, and enable automatic shape resizing to fit the linked text using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | waterfall chart | linked data labels | auto fit label shape | IsResizeShapeToFitText | DataLabelShapeType | Excel chart programmatically | chart series | export to Excel
// Common Searches: Aspose.Cells waterfall chart linked data labels | auto resize data label shape Aspose.Cells .NET | set rectangular data label shape Aspose.Cells | link chart data labels to cells C2:C5 | create waterfall chart with Aspose.Cells C#
// Developer Intent: Create a waterfall chart, bind its series to a value range, link the data‑label text to a separate summary range, show the linked text, and make label shapes automatically resize to fit the text.
// Use Cases: Financial reporting where start, increase, decrease, and end values need custom labels sourced from another column. | Dynamic dashboards that adjust label shapes automatically when summary text changes, keeping the layout tidy. | Generating Excel workbooks for presentations with waterfall charts that preserve linked label text and rectangular formatting.
// AI Prompts: Write C# code with Aspose.Cells to add a waterfall chart, link its data labels to cells C2:C5, show values, set rectangular shape, and enable IsResizeShapeToFitText. | Explain the steps to configure linked data labels and auto‑fit shape behavior for a waterfall chart in Aspose.Cells for .NET. | Provide a modification guide to change an existing Aspose.Cells chart so that its data labels use a rectangular shape and are linked to a summary range.

using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Shows how to generate a waterfall chart in a new workbook, link its data‑label text to a summary range (C2:C5), display values, set a rectangular label shape, and enable automatic shape resizing to fit the linked text using Aspose.Cells for .NET.
class WaterfallChartDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate data for the waterfall chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["A2"].PutValue("Start");
        worksheet.Cells["B2"].PutValue(100);
        worksheet.Cells["A3"].PutValue("Increase");
        worksheet.Cells["B3"].PutValue(30);
        worksheet.Cells["A4"].PutValue("Decrease");
        worksheet.Cells["B4"].PutValue(-20);
        worksheet.Cells["A5"].PutValue("End");
        worksheet.Cells["B5"].PutValue(110);

        // Summary range that will be linked to the data labels
        worksheet.Cells["C2"].PutValue("Start");
        worksheet.Cells["C3"].PutValue("Inc");
        worksheet.Cells["C4"].PutValue("Dec");
        worksheet.Cells["C5"].PutValue("End");

        // Add a waterfall chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Waterfall, 7, 0, 20, 10);
        Chart chart = worksheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.NSeries.Add("B2:B5", true);
        chart.NSeries.CategoryData = "A2:A5";

        // Access the first series and configure its data labels
        Series series = chart.NSeries[0];
        series.DataLabels.ShowValue = true;                     // Show the numeric values
        series.DataLabels.LinkedSource = "C2:C5";               // Link labels to the summary range
        series.DataLabels.ShowCellRange = true;                // Use the linked cells as label text
        series.DataLabels.IsResizeShapeToFitText = true;       // Auto‑fit shape to the text
        series.DataLabels.ShapeType = DataLabelShapeType.Rect; // Optional: set a rectangular shape

        // Save the workbook with the chart
        workbook.Save("WaterfallChartWithLinkedLabels.xlsx");
    }
}
