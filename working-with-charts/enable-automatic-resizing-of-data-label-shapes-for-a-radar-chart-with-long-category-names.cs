// Title: How to auto‑resize data label shapes for long category names in a radar chart using Aspose.Cells for .NET
// AI Prompts: Write C# code that builds a radar chart with long category names, displays those names in data labels, and enables the label shapes to automatically resize to fit the text with Aspose.Cells. | Show the steps to set the IsResizeShapeToFitText property on a radar chart series, recalculate the chart, and save the workbook in Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# radar chart auto resize data labels for long category names | How to make radar chart data label shapes fit long text using Aspose.Cells | Enable IsResizeShapeToFitText on radar chart series in .NET | Resize radar chart data labels automatically with Aspose.Cells example
// Tags: radar chart data label auto‑fit Aspose.Cells | resize label shape to fit long category names C# | set IsResizeShapeToFitText property Aspose.Cells | enable radar axis labels with long text .NET | auto‑adjust chart label dimensions Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsRadarChartAutoResize
{
    // Creates a workbook, adds a radar chart with long category names, shows those names in data labels, sets IsResizeShapeToFitText to true so label shapes auto‑resize to fit the text, recalculates the chart layout, and saves the file as RadarChart_AutoResizeDataLabels.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data with long category names
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Very Long Category Name 1");
                sheet.Cells["A3"].PutValue("Extremely Long Category Name 2");
                sheet.Cells["A4"].PutValue("Super Long Category Name 3");

                sheet.Cells["B1"].PutValue("Series1");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                // Add a radar chart
                int chartIndex = sheet.Charts.Add(ChartType.Radar, 5, 0, 20, 12);
                Chart chart = sheet.Charts[chartIndex];

                // Set series data and category data
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Enable axis labels for radar chart (category names)
                chart.NSeries[0].HasRadarAxisLabels = true;

                // Enable data labels for the series
                Series series = chart.NSeries[0];
                series.DataLabels.ShowCategoryName = true;   // display the long category names
                series.DataLabels.Position = LabelPositionType.Center;

                // Allow the data label shape to auto‑fit the text
                series.DataLabels.IsResizeShapeToFitText = true;

                // Recalculate the chart to apply layout changes
                chart.Calculate();

                // Save the workbook
                workbook.Save("RadarChart_AutoResizeDataLabels.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
