using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    public class SparklineOutlierMarkersDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Sample data – some values exceed the outlier threshold (e.g., > 8)
                sheet.Cells["A1"].PutValue(5);
                sheet.Cells["B1"].PutValue(12); // outlier
                sheet.Cells["C1"].PutValue(3);
                sheet.Cells["D1"].PutValue(15); // outlier
                sheet.Cells["E1"].PutValue(4);

                // Define the location where the sparkline will be placed (column F, row 1)
                CellArea sparklineLocation = new CellArea
                {
                    StartColumn = 5, // column F (0‑based index)
                    EndColumn = 5,
                    StartRow = 0,
                    EndRow = 0
                };

                // Add a sparkline group for the data range A1:E1
                int groupIdx = sheet.SparklineGroups.Add(SparklineType.Line, "A1:E1", false, sparklineLocation);
                SparklineGroup group = sheet.SparklineGroups[groupIdx];

                // Add the sparkline to the group (the same range, placed at F1)
                group.Sparklines.Add(sheet.Name + "!A1:E1", 0, 5);

                // Enable markers – this will draw a marker for every point
                group.ShowMarkers = true;

                // Set a distinct color for the markers (e.g., Red) to make outliers stand out
                CellsColor markersColor = workbook.CreateCellsColor();
                markersColor.Color = Color.Red;
                group.MarkersColor = markersColor;

                // OPTIONAL: Highlight the highest and lowest points as additional visual cues
                group.ShowHighPoint = true;
                group.HighPointColor.Color = Color.Green;
                group.ShowLowPoint = true;
                group.LowPointColor.Color = Color.Blue;

                // Save the workbook
                string outputPath = "SparklineOutlierMarkersDemo.xlsx";
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
            SparklineOutlierMarkersDemo.Run();
        }
    }
}