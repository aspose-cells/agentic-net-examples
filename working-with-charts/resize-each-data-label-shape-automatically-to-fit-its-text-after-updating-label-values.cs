// Title: Auto‑Resize Chart Data Labels to Fit Text with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, add a column chart, enable data labels, update each label's text, and automatically resize the label shape (IsResizeShapeToFitText) with optional font scaling using Aspose.Cells for .NET.
// Keywords: Aspose.Cells chart data label resize | IsResizeShapeToFitText .NET | auto fit chart labels Aspose | ChartPoint DataLabels update | auto scale font Aspose.Cells
// Common Searches: Aspose.Cells resize data label shape to fit text | set IsResizeShapeToFitText for chart points | auto‑fit chart data labels C# | update chart label text programmatically Aspose | auto scale font for Excel chart labels
// Developer Intent: Automatically adjust each chart data label's shape so it expands to contain the updated text.
// Use Cases: Add a prefix to every data label in a column chart while ensuring the label box grows to fit the longer string. | Generate Excel reports where label values vary widely and need dynamic resizing to avoid truncation. | Programmatically modify chart point labels and keep them readable by enabling automatic shape resizing and font scaling.
// AI Prompts: Write C# code with Aspose.Cells that sets IsResizeShapeToFitText and AutoScaleFont for all data labels in a chart. | Explain how to change ChartPoint.DataLabels.Text and enable automatic shape resizing in Aspose.Cells. | Show how to loop through series points, add a custom prefix to each label, and activate auto‑fit behavior.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add a column chart, enable data labels, update each label's text, and automatically resize the label shape (IsResizeShapeToFitText) with optional font scaling using Aspose.Cells for .NET.
    public class ResizeDataLabelsToFitText
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Alpha");
                sheet.Cells["A3"].PutValue("Beta");
                sheet.Cells["A4"].PutValue("Gamma");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(123);
                sheet.Cells["B3"].PutValue(4567);
                sheet.Cells["B4"].PutValue(89);

                // Add a column chart
                int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIdx];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Enable data labels for the first series
                Series series = chart.NSeries[0];
                series.DataLabels.ShowValue = true;

                // Update each data label's text and enable auto‑resize to fit the text
                foreach (ChartPoint point in series.Points)
                {
                    // Prepend a custom prefix to the label text
                    point.DataLabels.Text = $"Value: {point.YValue}";

                    // Ensure the shape of the data label resizes automatically to contain the text
                    point.DataLabels.IsResizeShapeToFitText = true;

                    // Optional: let the font scale automatically when the shape size changes
                    point.DataLabels.AutoScaleFont = true;
                }

                // Save the workbook
                workbook.Save("ResizeDataLabelsToFitText.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
