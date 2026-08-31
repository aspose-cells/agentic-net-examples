// Title: Generate a Waterfall Chart with Linked Data Labels and Auto‑Resizing Shapes using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that builds a waterfall chart, connects its data labels to a separate label column, and turns on automatic shape resizing. | Demonstrate setting the LinkedSource property and enabling IsResizeShapeToFitText for data labels in an Aspose.Cells waterfall chart.
// Common Searches: asp.net create waterfall chart with custom label column using Aspose.Cells | how to connect chart data labels to cells in Aspose.Cells C# | enable auto‑fit for data label shapes in Aspose.Cells waterfall chart | set label position to InsideEnd for waterfall chart Aspose.Cells | save waterfall chart to XLSX file with Aspose.Cells example
// Tags: waterfall chart linked labels Aspose.Cells | auto resize data label shape Aspose.Cells | LinkedSource property chart series Aspose.Cells | LabelPositionType InsideEnd Aspose.Cells | save waterfall chart to XLSX Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // The example creates a new workbook, fills columns A‑C with categories, numeric values, and custom label texts, adds a Waterfall chart, defines its data range, links the series' data labels to column C, enables automatic resizing of label shapes, positions the labels inside the end of each bar, recalculates the chart, and saves the result as WaterfallChartWithLinkedLabels.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate data for a waterfall chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Start");
                sheet.Cells["A3"].PutValue("Increase");
                sheet.Cells["A4"].PutValue("Decrease");
                sheet.Cells["A5"].PutValue("End");

                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(100);   // Start
                sheet.Cells["B3"].PutValue(30);    // Increase
                sheet.Cells["B4"].PutValue(-20);   // Decrease
                sheet.Cells["B5"].PutValue(110);   // End (calculated)

                sheet.Cells["C1"].PutValue("Label");
                sheet.Cells["C2"].PutValue("Start");
                sheet.Cells["C3"].PutValue("Gain");
                sheet.Cells["C4"].PutValue("Loss");
                sheet.Cells["C5"].PutValue("Total");

                // Add a waterfall chart
                int chartIdx = sheet.Charts.Add(ChartType.Waterfall, 7, 0, 25, 15);
                Chart chart = sheet.Charts[chartIdx];

                // Set the data range for the chart (values only)
                chart.SetChartDataRange("A1:B5", true);

                // Configure the series (values are taken from column B)
                chart.NSeries.Add("B2:B5", true);
                chart.NSeries.CategoryData = "A2:A5";

                // Enable data labels for the series
                Series series = chart.NSeries[0];
                series.DataLabels.ShowValue = true;

                // Link data labels to the summary range (column C)
                series.DataLabels.LinkedSource = "C2:C5";
                series.DataLabels.ShowCellRange = true;   // Use linked cells as label text

                // Adjust label appearance (shape type setting removed due to API compatibility)
                series.DataLabels.IsResizeShapeToFitText = true; // Auto‑fit shape to text
                series.DataLabels.Position = LabelPositionType.InsideEnd; // Reasonable position

                // Recalculate the chart to apply changes
                chart.Calculate();

                // Save the workbook
                string outputPath = "WaterfallChartWithLinkedLabels.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
