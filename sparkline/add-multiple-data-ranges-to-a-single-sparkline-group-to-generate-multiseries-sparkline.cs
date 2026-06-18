using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsMultiSeriesSparkline
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

                // -------------------------------------------------
                // Populate sample data for three series (rows 1‑3)
                // Each series will be plotted as a separate sparkline
                // -------------------------------------------------
                for (int col = 0; col < 6; col++)
                {
                    sheet.Cells[0, col].PutValue(col + 1);               // Series 1
                    sheet.Cells[1, col].PutValue((col + 1) * 2);         // Series 2
                    sheet.Cells[2, col].PutValue((col + 1) * 3);         // Series 3
                }

                // -------------------------------------------------
                // Define the location for the first sparkline (cell H1)
                // -------------------------------------------------
                CellArea firstLocation = new CellArea
                {
                    StartRow = 0,
                    EndRow = 0,
                    StartColumn = 7,
                    EndColumn = 7
                };

                // -------------------------------------------------
                // Add a sparkline group of type Line for the first series
                // -------------------------------------------------
                int groupIndex = sheet.SparklineGroups.Add(
                    SparklineType.Line,
                    "A1:F1",          // data range for first series
                    false,            // plot by row (horizontal)
                    firstLocation);

                SparklineGroup group = sheet.SparklineGroups[groupIndex];

                // -------------------------------------------------
                // Add second and third sparklines to the same group,
                // each pointing to a different data range.
                // -------------------------------------------------
                group.Sparklines.Add("A2:F2", 1, 7); // series 2 at H2
                group.Sparklines.Add("A3:F3", 2, 7); // series 3 at H3

                // Optional: customize appearance (same for all series)
                CellsColor seriesColor = workbook.CreateCellsColor();
                seriesColor.Color = Color.DarkBlue;
                group.SeriesColor = seriesColor;
                group.ShowHighPoint = true;
                group.ShowLowPoint = true;

                // -------------------------------------------------
                // Save the workbook
                // -------------------------------------------------
                workbook.Save("MultiSeriesSparkline.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}