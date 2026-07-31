// Title: Resize Data Label Shapes for a Stacked Column Chart in C# using Aspose.Cells
// Description: Shows how to create a workbook, add a stacked column chart, prepend custom text to each data label, enable automatic shape resizing with IsResizeShapeToFitText, recalculate the chart, and save the file so the labels expand to fit the longer content.
// Keywords: Aspose.Cells | C# | .NET | stacked column chart | data labels | IsResizeShapeToFitText | auto resize label shape | chart label length | Excel chart programming | Aspose.Cells chart API
// Common Searches: Aspose.Cells resize data label shape C# | auto fit chart data labels after changing text | IsResizeShapeToFitText example Aspose.Cells | increase data label length stacked column chart .NET | recalculate chart after modifying data labels Aspose
// Developer Intent: Automatically adjust the size of data label shapes so they accommodate longer custom text in a stacked column chart.
// Use Cases: Add a prefix (e.g., "Sales:") to each label and let the shape grow to avoid truncation. | Generate Excel reports where label content varies in length and must remain fully visible. | Create dynamic charts that display units or descriptions without manual resizing of label boxes.
// AI Prompts: Provide C# code that sets IsResizeShapeToFitText for each ChartPoint in a stacked column chart using Aspose.Cells. | How can I prepend custom text to data labels and make the label shapes auto‑expand in an Aspose.Cells chart? | Explain the steps to recalculate an Aspose.Cells chart after updating data label text so the resized shapes are applied.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsResizeDataLabelShapes
{
    // Shows how to create a workbook, add a stacked column chart, prepend custom text to each data label, enable automatic shape resizing with IsResizeShapeToFitText, recalculate the chart, and save the file so the labels expand to fit the longer content.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for a stacked column chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Q1");
                sheet.Cells["A3"].PutValue("Q2");
                sheet.Cells["A4"].PutValue("Q3");

                // Series 1
                sheet.Cells["B1"].PutValue("Product A");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                // Series 2
                sheet.Cells["C1"].PutValue("Product B");
                sheet.Cells["C2"].PutValue(15);
                sheet.Cells["C3"].PutValue(25);
                sheet.Cells["C4"].PutValue(35);

                // Add a stacked column chart (use ColumnStacked which is supported)
                int chartIndex = sheet.Charts.Add(ChartType.ColumnStacked, 5, 0, 20, 12);
                Chart chart = sheet.Charts[chartIndex];

                // Set data range for both series
                chart.NSeries.Add("B2:C4", true);               // Values
                chart.NSeries.CategoryData = "A2:A4";           // Categories

                // Enable data labels for the first series
                Series series = chart.NSeries[0];
                series.DataLabels.ShowValue = true;             // Show the numeric value
                series.DataLabels.Position = LabelPositionType.Center;

                // Increase the label text length by prefixing a custom string
                foreach (ChartPoint point in series.Points)
                {
                    point.DataLabels.Text = $"Sales: {point.YValue} units";
                    point.DataLabels.IsResizeShapeToFitText = true;
                }

                // Recalculate the chart to apply changes
                chart.Calculate();

                // Save the workbook
                string outputPath = "StackedColumn_ResizedDataLabels.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
