using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace SparklineDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the sparkline (row 0, columns A-D)
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["B1"].PutValue(2);
            sheet.Cells["C1"].PutValue(1);
            sheet.Cells["D1"].PutValue(3);

            // Define the location where the sparkline will be placed (cell E1)
            CellArea location = new CellArea
            {
                StartRow = 0,
                EndRow = 0,
                StartColumn = 4, // Column E (0‑based index)
                EndColumn = 4
            };

            // Add a sparkline group of type Line, using the data range A1:D1,
            // not vertical (plot by columns), and place it in the defined location
            int groupIdx = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, location);
            SparklineGroup group = sheet.SparklineGroups[groupIdx];

            // Add a sparkline to the group (data range A1:D1, placed at row 0, column 4)
            int sparklineIdx = group.Sparklines.Add("A1:D1", 0, 4);
            Sparkline sparkline = group.Sparklines[sparklineIdx];

            // Configure visual properties of the sparkline group
            // Series (line) color
            CellsColor seriesColor = workbook.CreateCellsColor();
            seriesColor.Color = Color.Orange;
            group.SeriesColor = seriesColor;

            // Show and color the highest point
            group.ShowHighPoint = true;
            CellsColor highColor = workbook.CreateCellsColor();
            highColor.Color = Color.Green;
            group.HighPointColor = highColor;

            // Show and color the lowest point
            group.ShowLowPoint = true;
            CellsColor lowColor = workbook.CreateCellsColor();
            lowColor.Color = Color.Red;
            group.LowPointColor = lowColor;

            // Set line weight
            group.LineWeight = 1.0;

            // Apply a preset style (optional)
            group.PresetStyle = SparklinePresetStyleType.Style5;

            // Save the workbook to an XLSX file
            workbook.Save("SparklineConfigured.xlsx");
        }
    }
}