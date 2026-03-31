using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace SparklineColorConfiguration
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the sparkline (A1:D1)
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["B1"].PutValue(-2);
            sheet.Cells["C1"].PutValue(8);
            sheet.Cells["D1"].PutValue(3);

            // Define the location where the sparkline will be placed (E1)
            CellArea location = new CellArea
            {
                StartRow = 0,
                EndRow = 0,
                StartColumn = 4,
                EndColumn = 4
            };

            // Add a sparkline group of type Line using the data range A1:D1
            int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, location);
            SparklineGroup group = sheet.SparklineGroups[groupIndex];

            // Add the sparkline to the group (required for some properties to take effect)
            group.Sparklines.Add(sheet.Name + "!A1:D1", 0, 4);

            // ----- Configure sparkline colors -----

            // Series (overall line) color
            CellsColor seriesColor = workbook.CreateCellsColor();
            seriesColor.Color = Color.DarkBlue;
            group.SeriesColor = seriesColor;

            // High point (maximum) color
            group.ShowHighPoint = true;
            CellsColor highColor = workbook.CreateCellsColor();
            highColor.Color = Color.Green;
            group.HighPointColor = highColor;

            // Low point (minimum) color
            group.ShowLowPoint = true;
            CellsColor lowColor = workbook.CreateCellsColor();
            lowColor.Color = Color.Red;
            group.LowPointColor = lowColor;

            // First point color
            group.ShowFirstPoint = true;
            CellsColor firstColor = workbook.CreateCellsColor();
            firstColor.Color = Color.Purple;
            group.FirstPointColor = firstColor;

            // Last point color
            group.ShowLastPoint = true;
            CellsColor lastColor = workbook.CreateCellsColor();
            lastColor.Color = Color.Orange;
            group.LastPointColor = lastColor;

            // Markers (each data point) color
            group.ShowMarkers = true;
            CellsColor markersColor = workbook.CreateCellsColor();
            markersColor.Color = Color.Black;
            group.MarkersColor = markersColor;

            // Negative points color
            group.ShowNegativePoints = true;
            CellsColor negativeColor = workbook.CreateCellsColor();
            negativeColor.Color = Color.Magenta;
            group.NegativePointsColor = negativeColor;

            // Horizontal axis color (enable axis display)
            group.ShowHorizontalAxis = true;
            CellsColor axisColor = workbook.CreateCellsColor();
            axisColor.Color = Color.Gray;
            group.HorizontalAxisColor = axisColor;

            // Optional: set line weight for better visibility
            group.LineWeight = 1.5;

            // Save the workbook to an XLSX file
            workbook.Save("SparklineColorsConfigured.xlsx", SaveFormat.Xlsx);
        }
    }
}