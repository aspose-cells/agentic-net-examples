// Title: Aspose.Cells .NET – Build a Waterfall Chart with Linked Data Labels and Auto‑Fit Shapes
// Description: Demonstrates how to create a workbook, add category/value data, define a summary column, insert a Waterfall chart, link the series data labels to the summary range, enable label text from cells, set rectangle label shapes, auto‑fit shapes to the text, position labels inside the bar ends, recalculate the chart, and save the file.
// Keywords: Aspose.Cells | C# waterfall chart | linked data labels | auto fit label shape | rectangle data label | label position inside end | Excel chart automation | .NET chart example
// Common Searches: Aspose.Cells link data labels to cells | waterfall chart auto‑fit label shape C# | set rectangle data label shape Aspose.Cells | position waterfall chart labels inside end | create waterfall chart with summary column Aspose
// Developer Intent: Generate a Waterfall chart whose data labels are sourced from a separate summary range and automatically resize to fit the label text.
// Use Cases: Financial reporting: show start, increase, decrease, and end values with descriptive text from a summary column. | Dynamic Excel reports: allow label text to be edited independently of chart data by linking to cells. | Improved readability: use rectangle labels that auto‑adjust size for long strings in exported waterfall charts.
// AI Prompts: Write C# code using Aspose.Cells to add a Waterfall chart, link its data labels to column C, and enable auto‑fit of label shapes. | Show how to set data label shape to rectangle and position labels inside the end of bars for a Waterfall chart in Aspose.Cells. | Provide an Aspose.Cells example that creates a Waterfall chart, links labels to a summary range, and resizes label shapes to fit the linked text.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsWaterfallDemo
{
    // Demonstrates how to create a workbook, add category/value data, define a summary column, insert a Waterfall chart, link the series data labels to the summary range, enable label text from cells, set rectangle label shapes, auto‑fit shapes to the text, position labels inside the bar ends, recalculate the chart, and save the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // ----- Sample data for waterfall chart -----
            // Categories
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Start");
            sheet.Cells["A3"].PutValue("Increase");
            sheet.Cells["A4"].PutValue("Decrease");
            sheet.Cells["A5"].PutValue("End");

            // Values (numeric)
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(100);
            sheet.Cells["B3"].PutValue(30);
            sheet.Cells["B4"].PutValue(-20);
            sheet.Cells["B5"].PutValue(110);

            // Summary range that will be linked to data labels
            sheet.Cells["C1"].PutValue("Summary");
            sheet.Cells["C2"].PutValue("Start Total");
            sheet.Cells["C3"].PutValue("Added");
            sheet.Cells["C4"].PutValue("Subtracted");
            sheet.Cells["C5"].PutValue("Final Total");

            // ----- Add a waterfall chart -----
            int chartIndex = sheet.Charts.Add(ChartType.Waterfall, 7, 0, 25, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart (values only)
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries.CategoryData = "A2:A5";

            // Enable data labels for the series
            Series series = chart.NSeries[0];
            series.DataLabels.ShowValue = true;               // Show numeric values
            series.DataLabels.ShowCellRange = true;           // Use linked source for label text
            series.DataLabels.LinkedSource = "C2:C5";         // Link to summary range

            // Adjust label shape to fit the text
            series.DataLabels.IsResizeShapeToFitText = true;  // Auto‑fit shape to text
            series.DataLabels.ShapeType = DataLabelShapeType.Rect; // Use rectangle shape

            // Optional: set a position that works well for waterfall charts
            series.DataLabels.Position = LabelPositionType.InsideEnd;

            // Calculate the chart to ensure positions are updated
            chart.Calculate();

            // Save the workbook
            workbook.Save("WaterfallChartWithLinkedLabels.xlsx");
        }
    }
}
