// Title: Adjusting data label shape size for wrapped multi‑line labels in an Aspose.Cells column chart (C#)
// Common Searches: Aspose.Cells C# set custom multi line data label text and adjust label size | Resize chart data label shape after enabling IsTextWrapped in Aspose.Cells | Prevent automatic data label shape auto‑fit and set height manually in Aspose.Cells chart | C# Aspose.Cells column chart data labels with wrapped text and fixed dimensions | Calculate data label height from line count for Aspose.Cells chart labels
// Tags: manual data label shape resizing Aspose.Cells | enable text wrap for chart data labels C# | disable automatic data label shape auto‑fit Aspose.Cells | set data label dimensions based on line count Aspose.Cells | column chart custom multi line labels C#

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace ResizeDataLabelShapesDemo
{
    // The example creates a workbook, adds a column chart, enables data labels with text wrapping, disables automatic shape resizing, assigns custom multi‑line text to each point, and manually sets each label's height (calculated from line count) and a fixed width before saving the file as an XLSX.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["B2"].PutValue(15000);
            sheet.Cells["B3"].PutValue(23000);
            sheet.Cells["B4"].PutValue(18000);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Enable data labels for the first series
            DataLabels seriesLabels = chart.NSeries[0].DataLabels;
            seriesLabels.ShowValue = true;               // Show the numeric value
            seriesLabels.Position = LabelPositionType.Center;
            seriesLabels.IsTextWrapped = true;           // Enable text wrapping

            // Disable automatic shape resizing – we will resize manually
            seriesLabels.IsResizeShapeToFitText = false;

            // Define a multi‑line custom text for each point
            string[] customTexts = new string[]
            {
                "First Quarter\n(Strong Growth)",
                "Second Quarter\n(Highest Sales)",
                "Third Quarter\n(Stable)"
            };

            // Apply custom text and manually resize each data label
            for (int i = 0; i < chart.NSeries[0].Points.Count; i++)
            {
                ChartPoint point = chart.NSeries[0].Points[i];
                DataLabels label = point.DataLabels;

                // Set the custom multi‑line text
                label.Text = customTexts[i];

                // Because we turned off auto‑fit, we need to give the label enough height.
                // Approximate height: 15 points per line + a small padding.
                int lineCount = customTexts[i].Split('\n').Length;
                label.Height = 15 * lineCount + 5;   // Height in points
                label.Width = 80;                    // Fixed width sufficient for wrapping

                // Ensure the label still respects the wrap setting
                label.IsTextWrapped = true;
                label.IsResizeShapeToFitText = false;
            }

            // Optional: Auto‑fit the column widths so the chart looks tidy
            sheet.AutoFitColumn(0);
            sheet.AutoFitColumn(1);

            // Save the workbook
            string outputPath = "ResizeDataLabelShapesDemo.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
