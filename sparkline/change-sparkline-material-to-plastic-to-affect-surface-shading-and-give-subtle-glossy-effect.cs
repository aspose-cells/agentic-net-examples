using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    public class SparklineMaterialDemo
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
                sheet.Cells["E1"].PutValue(4);

                // Define the location where the sparkline will be placed (column F, row 1)
                CellArea location = new CellArea
                {
                    StartRow = 0,
                    EndRow = 0,
                    StartColumn = 5,
                    EndColumn = 5
                };

                // Add a line sparkline group with the data range A1:E1
                int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:E1", false, location);
                SparklineGroup group = sheet.SparklineGroups[groupIndex];

                // Add the sparkline to the group (the Add method already created one, but we keep it for clarity)
                group.Sparklines.Add(sheet.Name + "!A1:E1", 0, 5);

                // OPTIONAL: customize appearance (color, markers, etc.)
                CellsColor seriesColor = workbook.CreateCellsColor();
                seriesColor.Color = Color.Orange;
                group.SeriesColor = seriesColor;
                group.ShowMarkers = true;
                CellsColor markersColor = workbook.CreateCellsColor();
                markersColor.Color = Color.Black;
                group.MarkersColor = markersColor;

                // ------------------------------------------------------------
                // NOTE: Aspose.Cells Sparkline objects do not expose ShapeProperties
                // or 3‑D formatting APIs. The original code attempted to set a
                // material type, which is not supported. This block is kept for
                // future compatibility and will be skipped safely.
                // ------------------------------------------------------------
                if (group.Sparklines.Count > 0)
                {
                    // Sparkline sparkline = group.Sparklines[0];
                    // No 3D material support – operation omitted.
                }

                // Save the workbook with the configured sparkline
                workbook.Save("SparklineMaterialPlastic.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            SparklineMaterialDemo.Run();
        }
    }
}