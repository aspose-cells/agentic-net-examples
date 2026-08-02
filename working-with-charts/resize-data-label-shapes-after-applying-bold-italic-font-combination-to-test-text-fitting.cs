// Title: Auto‑Resize Chart Data Label Shapes After Applying Bold‑Italic Font with Aspose.Cells for .NET (C#)
// Description: C# example that creates a column chart, enables data labels, applies a bold‑italic dark‑blue font (size 12) to all labels, propagates the style with ApplyFont, and sets IsResizeShapeToFitText to true for each ChartPoint so the label shapes automatically expand to fit the styled text.
// Keywords: Aspose.Cells | C# chart data labels | auto resize data label shape | IsResizeShapeToFitText | bold italic font chart labels | ApplyFont method | column chart example | .NET Excel automation | ResizeDataLabelShapesDemo
// Common Searches: Aspose.Cells auto‑resize chart data label after font change | C# set bold and italic font for chart labels Aspose | IsResizeShapeToFitText property usage in Aspose.Cells | How to fit data label shape to text in Excel with Aspose | Resize individual data label shapes in a column chart .NET
// Developer Intent: Enable each chart data label shape to automatically grow and fit bold‑italic text after changing the font.
// Use Cases: Generate a column chart where data labels use a bold‑italic style without clipping. | Programmatically apply a uniform font to all data labels and let the shapes self‑adjust. | Set an initial small width/height for labels and rely on IsResizeShapeToFitText to handle longer values.
// AI Prompts: Show C# code that sets IsResizeShapeToFitText true for every ChartPoint after applying a bold‑italic font with Aspose.Cells. | Explain why ApplyFont is needed before resizing chart data label shapes in Aspose.Cells. | Provide a step‑by‑step guide to auto‑fit data label shapes to styled text in a column chart using Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // C# example that creates a column chart, enables data labels, applies a bold‑italic dark‑blue font (size 12) to all labels, propagates the style with ApplyFont, and sets IsResizeShapeToFitText to true for each ChartPoint so the label shapes automatically expand to fit the styled text.
    public class ResizeDataLabelShapesDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");

                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Enable data labels for the first series
                Series series = chart.NSeries[0];
                series.DataLabels.ShowValue = true;

                // Apply bold and italic font to the data labels
                series.DataLabels.Font.IsBold = true;
                series.DataLabels.Font.IsItalic = true;
                series.DataLabels.Font.Color = Color.DarkBlue;
                series.DataLabels.Font.Size = 12;

                // Propagate the font settings to all individual data label objects
                series.DataLabels.ApplyFont();

                // For each data point, enable auto‑resize of the label shape to fit the new text
                foreach (ChartPoint point in series.Points)
                {
                    // Allow the shape to auto‑fit the text after the font change
                    point.DataLabels.IsResizeShapeToFitText = true;

                    // Optionally set an initial size that is smaller than the expected text size
                    point.DataLabels.Width = 40;
                    point.DataLabels.Height = 20;
                }

                // Save the workbook
                string outputPath = "ResizeDataLabelShapesDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            ResizeDataLabelShapesDemo.Run();
        }
    }
}
