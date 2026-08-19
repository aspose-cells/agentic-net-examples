// Title: Resize Chart Data Labels After Adding Superscript Text with Aspose.Cells for .NET
// Description: Creates a workbook with a column chart, enables data labels, appends a superscript character to each label, applies superscript formatting, temporarily disables auto‑resize, sets a fixed WidthPixel and HeightPixel, then re‑enables auto‑fit before saving the file.
// Keywords: Aspose.Cells | C# chart data label resize | superscript data label Aspose.Cells | set WidthPixel HeightPixel chart label | auto fit data label after formatting | chart label shape size .NET | Aspose.Cells data label styling
// Common Searches: Aspose.Cells resize data label after superscript | set fixed size for chart data labels C# | how to apply superscript to chart label Aspose.Cells | auto‑fit data label after custom font changes | adjust width and height of Excel chart data labels programmatically
// Developer Intent: Adjust the size of chart data label shapes so they display superscript characters correctly without clipping.
// Use Cases: Generate a column chart where each label shows a value with an exponent (e.g., 10²) and ensure the label box expands to accommodate the superscript. | Programmatically define a specific pixel width and height for each data label, apply superscript styling to part of the text, then restore auto‑fit for consistent appearance. | Create Excel reports that include chart labels with unit symbols or footnote markers as superscript while preventing overlap with other chart elements.
// AI Prompts: Write C# code using Aspose.Cells to add a superscript character to the end of each chart data label, set WidthPixel and HeightPixel, and re‑enable IsResizeShapeToFitText. | Show how to loop through ChartPoint objects, apply Characters().Font.IsSuperscript, call DataLabels.ApplyFont, and adjust DataLabels.IsResizeShapeToFitText, WidthPixel, and HeightPixel to avoid truncation. | Explain the steps for combining superscript formatting with shape resizing so that auto‑fit works correctly for chart data labels in an Aspose.Cells workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Creates a workbook with a column chart, enables data labels, appends a superscript character to each label, applies superscript formatting, temporarily disables auto‑resize, sets a fixed WidthPixel and HeightPixel, then re‑enables auto‑fit before saving the file.
    public class ResizeDataLabelAfterSuperscript
    {
        // Entry point required for console execution
        public static void Main(string[] args)
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
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Enable data labels for the first series
            Series series = chart.NSeries[0];
            series.DataLabels.ShowValue = true;
            series.DataLabels.Position = LabelPositionType.Center;

            // Customize each data label to include a superscript character
            foreach (ChartPoint point in series.Points)
            {
                // Set label text (e.g., "10²")
                point.DataLabels.Text = $"{point.YValue}2";

                // Apply superscript style to the last character
                int superscriptStart = point.DataLabels.Text.Length - 1;
                point.DataLabels.Characters(superscriptStart, 1).Font.IsSuperscript = true;

                // Apply font changes to the label
                point.DataLabels.ApplyFont();

                // Temporarily disable auto‑resize, set a small size, then re‑enable auto‑fit
                point.DataLabels.IsResizeShapeToFitText = false;
                point.DataLabels.WidthPixel = 40;
                point.DataLabels.HeightPixel = 20;
                point.DataLabels.IsResizeShapeToFitText = true;
            }

            // Save the workbook
            workbook.Save("ResizeDataLabelAfterSuperscript.xlsx");
        }
    }
}
