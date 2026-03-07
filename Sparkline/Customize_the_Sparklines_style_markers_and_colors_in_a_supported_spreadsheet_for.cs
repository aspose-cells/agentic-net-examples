using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace SparklineCustomizationDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the sparkline (row 1)
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["B1"].PutValue(2);
            sheet.Cells["C1"].PutValue(8);
            sheet.Cells["D1"].PutValue(3);
            sheet.Cells["E1"].PutValue(7);

            // Define the location where the sparkline will be placed (cell F1)
            CellArea location = new CellArea
            {
                StartRow = 0,
                EndRow = 0,
                StartColumn = 5,
                EndColumn = 5
            };

            // Add a sparkline group of type Line with the data range A1:E1
            int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:E1", false, location);
            SparklineGroup group = sheet.SparklineGroups[groupIndex];

            // Add the sparkline to the group (the Add method also creates the sparkline)
            group.Sparklines.Add(sheet.Name + "!A1:E1", 0, 5);

            // ----- Customize appearance -----

            // 1. Series (line) color
            CellsColor seriesColor = workbook.CreateCellsColor();
            seriesColor.Color = Color.Orange;
            group.SeriesColor = seriesColor;

            // 2. High and low point colors
            group.ShowHighPoint = true;
            group.ShowLowPoint = true;
            CellsColor highColor = workbook.CreateCellsColor();
            highColor.Color = Color.Green;
            group.HighPointColor = highColor;
            CellsColor lowColor = workbook.CreateCellsColor();
            lowColor.Color = Color.Red;
            group.LowPointColor = lowColor;

            // 3. First and last point colors
            group.ShowFirstPoint = true;
            group.ShowLastPoint = true;
            CellsColor firstColor = workbook.CreateCellsColor();
            firstColor.Color = Color.Purple;
            group.FirstPointColor = firstColor;
            CellsColor lastColor = workbook.CreateCellsColor();
            lastColor.Color = Color.Yellow;
            group.LastPointColor = lastColor;

            // 4. Markers (show each point) and marker color
            group.ShowMarkers = true;
            CellsColor markerColor = workbook.CreateCellsColor();
            markerColor.Color = Color.Black;
            group.MarkersColor = markerColor;

            // 5. Negative points color and visibility
            group.ShowNegativePoints = true;
            CellsColor negativeColor = workbook.CreateCellsColor();
            negativeColor.Color = Color.Blue;
            group.NegativePointsColor = negativeColor;

            // 6. Preset style (optional, overrides many individual settings if not Custom)
            group.PresetStyle = SparklinePresetStyleType.Style5;

            // 7. Line weight (thickness)
            group.LineWeight = 1.2;

            // 8. Horizontal axis visibility and color
            group.ShowHorizontalAxis = true;
            CellsColor hAxisColor = workbook.CreateCellsColor();
            hAxisColor.Color = Color.Gray;
            group.HorizontalAxisColor = hAxisColor;

            // 9. Plot empty cells as zero and display hidden cells
            group.PlotEmptyCellsType = PlotEmptyCellsType.Zero;
            group.DisplayHidden = true;

            // Save the workbook with the customized sparkline
            workbook.Save("CustomizedSparkline.xlsx", SaveFormat.Xlsx);
        }
    }
}