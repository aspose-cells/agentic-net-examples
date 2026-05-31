using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using System.Drawing;

namespace AsposeCellsHeatMapDataLabelResize
{
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for a heat‑map‑like chart
                sheet.Cells["A1"].PutValue("Region");
                sheet.Cells["A2"].PutValue("North");
                sheet.Cells["A3"].PutValue("South");
                sheet.Cells["A4"].PutValue("East");
                sheet.Cells["A5"].PutValue("West");

                sheet.Cells["B1"].PutValue("Q1");
                sheet.Cells["C1"].PutValue("Q2");
                sheet.Cells["D1"].PutValue("Q3");

                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["C2"].PutValue(150);
                sheet.Cells["D2"].PutValue(130);

                sheet.Cells["B3"].PutValue(80);
                sheet.Cells["C3"].PutValue(95);
                sheet.Cells["D3"].PutValue(100);

                sheet.Cells["B4"].PutValue(200);
                sheet.Cells["C4"].PutValue(180);
                sheet.Cells["D4"].PutValue(210);

                sheet.Cells["B5"].PutValue(160);
                sheet.Cells["C5"].PutValue(170);
                sheet.Cells["D5"].PutValue(155);

                // Add a chart (using ColumnStacked as a substitute for HeatMap)
                int chartIndex = sheet.Charts.Add(ChartType.ColumnStacked, 7, 0, 25, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range (values only) and category labels (rows)
                chart.NSeries.Add("B2:D5", true);
                chart.NSeries.CategoryData = "A2:A5";

                // Apply a one‑color vertical gradient to the series
                chart.NSeries[0].Area.FillFormat.SetOneColorGradient(
                    Color.LightSeaGreen,          // base color
                    0.5,                          // brightness (0‑1)
                    GradientStyleType.Vertical,   // direction
                    1);                           // variant

                // Enable data labels and position them at the centre
                chart.NSeries[0].DataLabels.ShowValue = true;
                chart.NSeries[0].DataLabels.Position = LabelPositionType.Center;

                // Resize each data label shape
                foreach (ChartPoint point in chart.NSeries[0].Points)
                {
                    point.DataLabels.IsResizeShapeToFitText = false; // disable auto‑fit
                    point.DataLabels.ShapeType = DataLabelShapeType.Rect; // rectangle shape
                    point.DataLabels.Width = 60;   // explicit width (points)
                    point.DataLabels.Height = 30;  // explicit height (points)
                }

                // Recalculate the chart to apply changes
                chart.Calculate();

                // Define output file path
                string outputPath = "HeatMapDataLabelResize.xlsx";

                // Save the workbook (overwrite if it already exists)
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}