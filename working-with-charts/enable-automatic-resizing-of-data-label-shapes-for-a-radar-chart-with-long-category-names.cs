// Title: Enable automatic resizing of radar chart data label shapes for long category names with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# how to make radar chart data label shapes auto‑fit long category names | set IsResizeShapeToFitText for each point in a radar chart using Aspose.Cells | auto scale font of radar chart data labels in C# with Aspose.Cells | resize data label shape to fit text in Aspose.Cells radar chart example | C# code to adjust radar chart data labels for long category strings in Aspose.Cells
// Tags: IsResizeShapeToFitText chart point Aspose.Cells | AutoScaleFont radar chart data labels | Radar chart data label auto fit text | Long category names radar chart Aspose.Cells | C# Aspose.Cells radar chart data labels | Resize data label shape to fit text

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace RadarChartDataLabelResize
{
    // The sample creates a workbook, adds a radar chart with very long category names, enables data labels, and configures each chart point's DataLabels to automatically resize the label shape to fit the text and to auto‑scale the font, then saves the workbook as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate data with long category names
            // Category names (A column)
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Very Long Category Name Number One");
            sheet.Cells["A3"].PutValue("Extremely Long Category Name Number Two");
            sheet.Cells["A4"].PutValue("Super Long Category Name Number Three");
            sheet.Cells["A5"].PutValue("Another Very Long Category Name Four");

            // Values for the series (B column)
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["B5"].PutValue(25);

            // Add a radar chart (standard radar)
            int chartIndex = sheet.Charts.Add(ChartType.Radar, 7, 0, 25, 15);
            Chart radarChart = sheet.Charts[chartIndex];

            // Set the data source for the series
            radarChart.NSeries.Add("B2:B5", true);          // Values
            radarChart.NSeries.CategoryData = "A2:A5";     // Categories

            // Enable data labels for the series
            Series series = radarChart.NSeries[0];
            series.DataLabels.ShowCategoryName = true;   // Show the long category names
            series.DataLabels.ShowValue = true;          // Show the numeric values

            // Iterate through each point and enable automatic shape resizing
            foreach (ChartPoint point in series.Points)
            {
                // This property makes the data label shape auto‑fit the text it contains.
                point.DataLabels.IsResizeShapeToFitText = true;

                // Optional: keep the font size adaptable when the shape size changes.
                point.DataLabels.AutoScaleFont = true;
            }

            // Save the workbook to a file
            string outputPath = "RadarChart_AutoResizeDataLabels.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}
