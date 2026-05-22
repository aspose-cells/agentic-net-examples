using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;   // For ShapeType enum (if needed)

namespace AsposeCellsExample
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for a waterfall chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Start");
                sheet.Cells["A3"].PutValue("Increase");
                sheet.Cells["A4"].PutValue("Decrease");
                sheet.Cells["A5"].PutValue("Total");

                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(100);
                sheet.Cells["B3"].PutValue(30);
                sheet.Cells["B4"].PutValue(-20);
                sheet.Cells["B5"].PutValue(110);

                // Summary range that will be linked to data labels (e.g., formatted text)
                sheet.Cells["C1"].PutValue("Summary");
                sheet.Cells["C2"].PutValue("Start");
                sheet.Cells["C3"].PutValue("Add 30");
                sheet.Cells["C4"].PutValue("Subtract 20");
                sheet.Cells["C5"].PutValue("Final");

                // Add a waterfall chart
                int chartIndex = sheet.Charts.Add(ChartType.Waterfall, 7, 0, 25, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data series and categories
                chart.NSeries.Add("B2:B5", true);
                chart.NSeries.CategoryData = "A2:A5";

                // Access the first (and only) series
                Series series = chart.NSeries[0];

                // Enable data labels
                series.DataLabels.ShowValue = true;

                // Link data labels to the summary range
                series.DataLabels.ShowCellRange = true;               // Show linked cell range as label
                series.DataLabels.LinkedSource = "C2:C5";             // Range with summary text

                // Adjust label shape to fit the linked text
                series.DataLabels.IsResizeShapeToFitText = true;      // Auto‑fit shape to text
                // series.DataLabels.ShapeType = ShapeType.Rectangle; // Removed: ShapeType not available in this context

                // Optionally set a position for better readability
                series.DataLabels.Position = LabelPositionType.InsideEnd;

                // Save the workbook
                string outputPath = "WaterfallChartWithLinkedLabels.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}