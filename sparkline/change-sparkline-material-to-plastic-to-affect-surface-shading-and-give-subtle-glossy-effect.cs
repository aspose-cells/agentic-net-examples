using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using System.Drawing;

namespace AsposeCellsExamples
{
    public class SparklineMaterialPlasticDemo
    {
        public static void Run()
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

                // Define the location where the sparkline will be placed (E1)
                CellArea location = new CellArea
                {
                    StartRow = 0,
                    EndRow = 0,
                    StartColumn = 4,
                    EndColumn = 4
                };

                // Add a line sparkline group using the data range A1:D1
                int sparklineGroupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, location);
                SparklineGroup sparklineGroup = sheet.SparklineGroups[sparklineGroupIndex];

                // OPTIONAL: customize sparkline appearance (color, markers, etc.)
                CellsColor seriesColor = workbook.CreateCellsColor();
                seriesColor.Color = Color.Orange;
                sparklineGroup.SeriesColor = seriesColor;
                sparklineGroup.ShowMarkers = true;

                // Demonstrate material change on a shape (TextBox) that supports ThreeDFormat
                Shape shape = sheet.Shapes.AddShape(MsoDrawingType.TextBox, 1, 1, 200, 100, 0, 0);
                shape.Text = "Sparkline Material: Plastic";

                // Set ThreeD material to Plastic
                shape.ThreeDFormat.Material = PresetMaterialType.Plastic;

                // Save the workbook
                string outputPath = "SparklineMaterialPlasticDemo.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            SparklineMaterialPlasticDemo.Run();
        }
    }
}