using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    public class SparklineHighLowMarkersDemo
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
                sheet.Cells["C1"].PutValue(8);
                sheet.Cells["D1"].PutValue(3);

                // Define the location where the sparkline will be placed (E1)
                CellArea location = new CellArea
                {
                    StartColumn = 4, // Column E (0‑based index)
                    EndColumn = 4,
                    StartRow = 0,    // Row 1
                    EndRow = 0
                };

                // Add a line sparkline group with the data range A1:D1
                int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, location);
                SparklineGroup group = sheet.SparklineGroups[groupIndex];

                // Add the sparkline to the group (optional, already added by Add method)
                group.Sparklines.Add(sheet.Name + "!A1:D1", 0, 4);

                // Enable markers for both high and low points
                group.ShowHighPoint = true;
                group.ShowLowPoint = true;

                // Set colors for the high and low point markers
                CellsColor highColor = workbook.CreateCellsColor();
                highColor.Color = Color.Green;
                group.HighPointColor = highColor;

                CellsColor lowColor = workbook.CreateCellsColor();
                lowColor.Color = Color.Red;
                group.LowPointColor = lowColor;

                // Save the workbook
                string outputPath = "SparklineHighLowMarkers.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            SparklineHighLowMarkersDemo.Run();
        }
    }
}