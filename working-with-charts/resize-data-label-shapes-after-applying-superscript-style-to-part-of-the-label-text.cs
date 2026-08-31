// Title: How to resize a chart data label shape after applying superscript formatting to part of the label text using Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a column chart, disables automatic data‑label shape resizing, sets a fixed pixel width, applies superscript to a specific character in each label, then re‑enables auto‑fit so the shape expands to fit the styled text. | Show how to use Aspose.Cells to format a portion of a chart data label with superscript, keep the label width constant initially, and programmatically trigger the label shape to resize after the font change.
// Common Searches: Aspose.Cells C# resize chart data label after superscript formatting | set fixed width for chart data labels then auto‑fit in .NET | apply superscript to part of a data label in an Aspose.Cells chart | disable data label auto resize and enable it later Aspose.Cells | how to adjust data label shape size after styling text in Aspose.Cells
// Tags: superscript formatting for chart data labels Aspose.Cells | fixed pixel width for data label shape .NET | auto‑fit data label shape after font change Aspose.Cells | column chart label size manipulation Aspose.Cells | partial text styling in chart labels C#

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace AsposeCellsExamples
{
    // The example demonstrates creating a workbook with a column chart, disabling automatic resizing of data label shapes, assigning a fixed pixel width, applying superscript to the last character of each label, and then re‑enabling auto‑fit so the label shape expands to accommodate the styled text before saving the file.
    public class ResizeDataLabelAfterSuperscript
    {
        // Entry point for the example
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
            int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIdx];

            // Set the data range for the series
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Enable data labels for the first series
            Series series = chart.NSeries[0];
            series.DataLabels.ShowValue = true;
            series.DataLabels.Position = LabelPositionType.Center;

            // Disable automatic resizing of the data label shape
            series.DataLabels.IsResizeShapeToFitText = false;

            // Set a custom width that is smaller than the full text would need
            series.DataLabels.WidthPixel = 40; // width in pixels

            // Iterate through each point and customize its label
            foreach (ChartPoint point in series.Points)
            {
                // Set the label text (e.g., "10")
                point.DataLabels.Text = point.YValue.ToString();

                // Apply superscript to the last character (for demonstration)
                // Characters(startIndex, length) – startIndex is zero‑based
                int textLength = point.DataLabels.Text.Length;
                if (textLength > 0)
                {
                    // Make the last character superscript
                    var chars = point.DataLabels.Characters(textLength - 1, 1);
                    chars.Font.IsSuperscript = true;
                }

                // Apply the font changes to the whole label
                point.DataLabels.ApplyFont();
            }

            // After formatting, enable auto‑fit so the shape expands to contain the superscript text
            series.DataLabels.IsResizeShapeToFitText = true;

            // Save the workbook
            string outputPath = "ResizeDataLabelAfterSuperscript.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
