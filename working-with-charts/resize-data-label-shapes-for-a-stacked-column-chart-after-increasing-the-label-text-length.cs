// Title: How to auto‑fit and resize data label shapes after extending label text in a stacked column chart using Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a stacked column chart, appends custom text to each data label, and enables the label shape to auto‑fit the longer text with Aspose.Cells. | Show how to set the IsResizeShapeToFitText property and define a minimum width for chart point labels in Aspose.Cells for .NET. | Demonstrate recalculating the chart after modifying data label properties to ensure the layout updates correctly.
// Common Searches: Aspose.Cells C# resize data label shape after increasing label text in stacked column chart | set IsResizeShapeToFitText true for chart point labels Aspose.Cells .NET | auto fit data label width stacked column chart Aspose.Cells example | how to change data label text and shape size programmatically in Aspose.Cells
// Tags: label shape resizing Aspose.Cells .NET | stacked column chart data label size C# | chart point label width configuration | extend data label text Aspose.Cells | auto‑fit label shape property Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // The example creates a workbook, adds sample data, inserts a stacked column chart, enables data labels, appends extra text to each point's label, activates automatic shape resizing with a minimum width, recalculates the chart layout, and saves the workbook as an XLSX file.
    public class ResizeDataLabelShapesDemo
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for a stacked column chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");

            // First series
            sheet.Cells["B1"].PutValue("Product A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Second series
            sheet.Cells["C1"].PutValue("Product B");
            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(35);

            // Add a stacked column chart (use ColumnStacked enum)
            int chartIndex = sheet.Charts.Add(ChartType.ColumnStacked, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set data range for the chart (including both series)
            chart.NSeries.Add("B2:C4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Enable data labels for each series
            foreach (Series ser in chart.NSeries)
            {
                ser.DataLabels.ShowValue = true;               // Show the numeric value
                ser.DataLabels.Position = LabelPositionType.Center;
            }

            // Increase the label text length for each point and resize the label shape
            foreach (Series ser in chart.NSeries)
            {
                foreach (ChartPoint pt in ser.Points)
                {
                    // Append extra text to make the label longer
                    pt.DataLabels.Text = $"Value: {pt.YValue} (extra info)";

                    // Allow the shape to auto‑fit the longer text
                    pt.DataLabels.IsResizeShapeToFitText = true;

                    // Optionally, set a minimum width so the shape does not become too small before auto‑fit
                    pt.DataLabels.Width = 80;   // width in pixels
                }
            }

            // Recalculate the chart to apply layout changes
            chart.Calculate();

            // Save the workbook
            workbook.Save("ResizeDataLabelShapesDemo.xlsx");
        }
    }
}
