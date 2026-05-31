using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace RadarChartDataLabelAutoResize
{
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
                sheet.Cells["A3"].PutValue("Extremely Long Category Name 2");
                sheet.Cells["A4"].PutValue("Super Long Category Name 3");

                sheet.Cells["B1"].PutValue("Series1");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                // Add a radar chart
                int chartIndex = sheet.Charts.Add(ChartType.Radar, 5, 0, 20, 12);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B4", true);          // Values
                chart.NSeries.CategoryData = "A2:A4";      // Categories

                // Enable category (axis) labels for the radar chart
                Series series = chart.NSeries[0];
                series.HasRadarAxisLabels = true;

                // Enable data labels and configure auto‑resize to fit long text
                series.DataLabels.ShowCategoryName = true;          // Show category names
                series.DataLabels.IsResizeShapeToFitText = true;    // Auto‑fit shape to text
                series.DataLabels.AutoScaleFont = true;            // Adjust font size as shape changes

                // The ShapeType property is not available in all versions; omit if unsupported
                // series.DataLabels.ShapeType = Aspose.Cells.Drawing.ShapeType.Rectangle;

                // Force chart layout calculation before saving
                chart.Calculate();

                // Prepare output file path
                string outputPath = "RadarChartAutoResizeDataLabels.xlsx";

                // Ensure the directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (string.IsNullOrEmpty(outputDir))
                {
                    outputDir = Directory.GetCurrentDirectory();
                }
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}