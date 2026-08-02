// Title: C# Aspose.Cells Example: Waterfall Chart with Linked Data Labels & Auto‑Fit Shapes
// Description: Creates a workbook, adds stage/value data, defines a summary range, inserts a Waterfall chart, links data labels to that range, enables auto‑resize of label shapes, sets a round‑rectangle shape, recalculates layout, and saves the file.
// Keywords: Aspose.Cells | C# | Waterfall chart | linked data labels | auto resize data label shape | DataLabelShapeType | Excel chart example | Aspose.Cells tutorial
// Common Searches: Aspose.Cells waterfall chart linked labels C# | auto‑fit data label shape Aspose.Cells | set data label shape type in Aspose.Cells chart | link data labels to another range Aspose.Cells | waterfall chart example Aspose.Cells .NET
// Developer Intent: Generate a Waterfall chart, link its data labels to a summary range, and make the label shapes automatically fit the linked text.
// Use Cases: Financial reporting where each waterfall step shows a formatted unit string from a separate column. | Automated Excel dashboards that need data label shapes to expand for longer summary texts. | Presentations that require rounded‑rectangle labels linked to a summary range for visual emphasis.
// AI Prompts: Give C# code to change the data label shape to an ellipse while keeping the linked source and auto‑fit behavior. | Show how to update the linked source range dynamically based on user‑selected rows in Aspose.Cells. | Explain how to apply a custom font style and color to linked data labels in a Waterfall chart.

using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Creates a workbook, adds stage/value data, defines a summary range, inserts a Waterfall chart, links data labels to that range, enables auto‑resize of label shapes, sets a round‑rectangle shape, recalculates layout, and saves the file.
class WaterfallChartExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the waterfall chart
        worksheet.Cells["A1"].PutValue("Stage");
        worksheet.Cells["A2"].PutValue("Start");
        worksheet.Cells["A3"].PutValue("Increase");
        worksheet.Cells["A4"].PutValue("Decrease");
        worksheet.Cells["A5"].PutValue("End");

        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(100);
        worksheet.Cells["B3"].PutValue(30);
        worksheet.Cells["B4"].PutValue(-20);
        worksheet.Cells["B5"].PutValue(110);

        // Summary range that will be linked to the data labels
        worksheet.Cells["C2"].PutValue("100 units");
        worksheet.Cells["C3"].PutValue("130 units");
        worksheet.Cells["C4"].PutValue("110 units");
        worksheet.Cells["C5"].PutValue("110 units");

        // Add a waterfall chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Waterfall, 7, 0, 25, 15);
        Chart chart = worksheet.Charts[chartIndex];

        // Define the series data and categories
        chart.NSeries.Add("B2:B5", true);
        chart.NSeries.CategoryData = "A2:A5";

        // Configure data labels for the series
        Series series = chart.NSeries[0];
        series.DataLabels.ShowValue = true;                     // display the value
        series.DataLabels.LinkedSource = "C2:C5";                // link to summary range
        series.DataLabels.IsResizeShapeToFitText = true;        // auto‑fit shape to text
        series.DataLabels.ShapeType = DataLabelShapeType.RoundRect; // optional shape type

        // Recalculate chart layout (important after modifying labels)
        chart.Calculate();

        // Save the workbook with the chart
        workbook.Save("WaterfallChartWithLinkedLabels.xlsx");
    }
}
