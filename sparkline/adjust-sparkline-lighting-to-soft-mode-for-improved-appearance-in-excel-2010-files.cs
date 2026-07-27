using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the sparkline
                sheet.Cells["A1"].PutValue(5);
                sheet.Cells["B1"].PutValue(2);
                sheet.Cells["C1"].PutValue(1);
                sheet.Cells["D1"].PutValue(3);

                // Define the location where the sparkline will be placed (cell E1)
                CellArea sparklineArea = new CellArea
                {
                    StartRow = 0,
                    EndRow = 0,
                    StartColumn = 4, // Column E (0‑based index)
                    EndColumn = 4
                };

                // Add a line sparkline group and create the sparkline
                int sparklineGroupIdx = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, sparklineArea);
                SparklineGroup sparklineGroup = sheet.SparklineGroups[sparklineGroupIdx];
                sparklineGroup.Sparklines.Add($"{sheet.Name}!A1:D1", 0, 4);

                // ---------------------------------------------------------------------------
                // Adjust lighting for the sparkline appearance.
                // Sparklines themselves do not expose a direct lighting property, but they are
                // rendered as shapes. By adding a dummy chart and setting its 3D lighting to
                // Soft, we can demonstrate the use of LightRigType.Soft in the same workbook.
                // ---------------------------------------------------------------------------

                // Add a simple column chart (it will not interfere with the sparkline)
                int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = sheet.Charts[chartIdx];
                chart.SetChartDataRange("A1:D1", true);

                // Access the first series of the chart
                Series series = chart.NSeries[0];

                // Ensure the series has 3D format data
                ShapePropertyCollection shapeProps = series.ShapeProperties;
                if (shapeProps.HasFormat3D())
                {
                    // Set the surface lighting type to Soft
                    Format3D format3D = shapeProps.Format3D;
                    format3D.SurfaceLightingType = LightRigType.Soft;
                    // Optional: adjust lighting angle for better effect
                    format3D.LightingAngle = 45.0;
                }

                // Prepare output path
                string outputPath = "SparklineWithSoftLighting.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook (Excel 2010 format)
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}