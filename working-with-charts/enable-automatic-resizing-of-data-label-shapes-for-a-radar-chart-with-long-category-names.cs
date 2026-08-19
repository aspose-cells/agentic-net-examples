// Title: Auto‑Resize Radar Chart Data Labels for Long Category Names in Aspose.Cells (.NET)
// Description: Demonstrates how to create a radar chart with lengthy category names, enable axis labels, show category names and values in data labels, and set IsResizeShapeToFitText so each label shape expands automatically before saving the workbook as XLSX.
// Keywords: Aspose.Cells | .NET | C# | radar chart | auto resize data labels | IsResizeShapeToFitText | long category names | chart axis labels | Excel export | data label shape fit
// Common Searches: Aspose.Cells auto resize radar chart data labels | C# radar chart long category names Aspose.Cells | IsResizeShapeToFitText property usage | show category name and value in radar chart labels | fit data label shape to text Aspose.Cells
// Developer Intent: Automatically adjust radar chart data label shapes to accommodate long category text.
// Use Cases: Generate radar charts with verbose category labels that remain fully readable. | Produce Excel reports where data labels display both category names and values without truncation. | Integrate auto‑fitting label shapes into .NET applications that export workbooks to XLSX.
// AI Prompts: Write C# code using Aspose.Cells to create a radar chart, enable axis labels, display category names and values in data labels, and set IsResizeShapeToFitText to true. | Explain the impact of the IsResizeShapeToFitText property on radar chart label rendering and any known limitations. | Provide a step‑by‑step guide for automatically resizing radar chart data label shapes for long category names in a .NET project.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsRadarChartAutoResize
{
    // Demonstrates how to create a radar chart with lengthy category names, enable axis labels, show category names and values in data labels, and set IsResizeShapeToFitText so each label shape expands automatically before saving the workbook as XLSX.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data with long category names
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Very Long Category Name 1");
                sheet.Cells["A3"].PutValue("Very Long Category Name 2");
                sheet.Cells["A4"].PutValue("Very Long Category Name 3");

                sheet.Cells["B1"].PutValue("Series1");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                // Add a radar chart (positioned from row 5, column 0 to row 20, column 12)
                int chartIndex = sheet.Charts.Add(ChartType.Radar, 5, 0, 20, 12);
                Chart chart = sheet.Charts[chartIndex];

                // Set series data and category data
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Enable axis labels for radar chart (required for category names)
                Series series = chart.NSeries[0];
                series.HasRadarAxisLabels = true;

                // Enable data labels and show category names and values
                series.DataLabels.ShowCategoryName = true;
                series.DataLabels.ShowValue = true;

                // Auto‑fit the data label shape to the text
                series.DataLabels.IsResizeShapeToFitText = true;

                // Note: Setting a specific shape type is optional; omitted if enum unavailable.

                // Recalculate the chart to apply layout changes
                chart.Calculate();

                // Save the workbook
                string outputPath = "RadarChartAutoResizeDataLabels.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
