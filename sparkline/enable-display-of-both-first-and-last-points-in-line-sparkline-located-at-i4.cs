using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSparklineDemo
{
    public class ShowFirstAndLastPoints
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add sample data for the sparkline (row 4, columns A to D)
                sheet.Cells["A4"].PutValue(5);
                sheet.Cells["B4"].PutValue(2);
                sheet.Cells["C4"].PutValue(8);
                sheet.Cells["D4"].PutValue(3);

                // Define the location cell I4 (column index 8, row index 3)
                CellArea location = new CellArea
                {
                    StartColumn = 8,
                    EndColumn = 8,
                    StartRow = 3,
                    EndRow = 3
                };

                // Add a line sparkline group with the data range A4:D4 and place it at I4
                int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A4:D4", false, location);
                SparklineGroup group = sheet.SparklineGroups[groupIndex];

                // Enable highlighting of the first and last points
                group.ShowFirstPoint = true;
                group.ShowLastPoint = true;

                // Set colors for the first and last points
                CellsColor firstPointColor = workbook.CreateCellsColor();
                firstPointColor.Color = Color.Purple;
                group.FirstPointColor = firstPointColor;

                CellsColor lastPointColor = workbook.CreateCellsColor();
                lastPointColor.Color = Color.Yellow;
                group.LastPointColor = lastPointColor;

                // Define output file path
                string outputPath = "SparklineFirstLastPoints.xlsx";

                // Ensure the directory exists before saving
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
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
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ShowFirstAndLastPoints.Run();
        }
    }
}