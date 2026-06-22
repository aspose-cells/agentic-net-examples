using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsGlowValidation
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // ---------- Sample shapes ----------
                // Add a rectangle shape and define a glow effect
                Shape rectShape = worksheet.Shapes.AddShape(
                    MsoDrawingType.Rectangle, // shape type
                    1, 0,                     // upper‑left row, column
                    1, 0,                     // top, left (in pixels)
                    100, 150);                // height, width (in pixels)

                rectShape.Glow.Size = 5;
                rectShape.Glow.Color = workbook.CreateCellsColor();
                rectShape.Glow.Color.Color = Color.Yellow;

                // Add a second shape without a glow (to demonstrate validation)
                Shape ellipseShape = worksheet.Shapes.AddShape(
                    MsoDrawingType.Rectangle, // fallback to rectangle if ellipse unavailable
                    2, 0,
                    2, 0,
                    80, 80);
                // No glow defined for ellipseShape

                // ---------- Sample chart with series shape properties ----------
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 10);
                Chart chart = worksheet.Charts[chartIndex];

                // Populate data for the chart
                worksheet.Cells["A1"].PutValue("Category 1");
                worksheet.Cells["A2"].PutValue("Category 2");
                worksheet.Cells["B1"].PutValue(10);
                worksheet.Cells["B2"].PutValue(20);

                chart.NSeries.Add("B1:B2", true);
                chart.NSeries.CategoryData = "A1:A2";

                // Access the first series and set a glow effect
                Series series = chart.NSeries[0];
                GlowEffect seriesGlow = series.ShapeProperties.GlowEffect;
                seriesGlow.Size = 8;
                seriesGlow.Color = workbook.CreateCellsColor();
                seriesGlow.Color.Color = Color.Green;

                // ---------- Validation: ensure every shape has a glow effect ----------
                foreach (Worksheet ws in workbook.Worksheets)
                {
                    foreach (Shape shp in ws.Shapes)
                    {
                        if (shp.Glow == null || shp.Glow.Size == 0)
                        {
                            Console.WriteLine($"Shape '{shp.Name}' does not have a glow effect.");
                        }
                        else
                        {
                            Console.WriteLine($"Shape '{shp.Name}' has glow size {shp.Glow.Size}.");
                        }
                    }
                }

                // Save the workbook (optional)
                string outputPath = "GlowValidationResult.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}