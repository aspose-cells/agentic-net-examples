using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace SparklineAppearanceDemo
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
            CellArea sparklineLocation = new CellArea
            {
                StartRow = 0,
                EndRow = 0,
                StartColumn = 5,
                EndColumn = 5
            };

            // Add a sparkline group of type Line with the data range A1:E1
            int groupIndex = sheet.SparklineGroups.Add(
                SparklineType.Line,          // Sparkline type
                "A1:E1",                     // Data range
                false,                       // Plot by row (false = by column)
                sparklineLocation);          // Where the sparkline will appear

            SparklineGroup group = sheet.SparklineGroups[groupIndex];

            // Add the sparkline to the group (optional when using Add with dataRange)
            // group.Sparklines.Add(sheet.Name + "!A1:E1", 0, 5);

            // ----- Customize visual properties -----

            // 1. Series (line) color
            CellsColor seriesColor = workbook.CreateCellsColor();
            seriesColor.Color = Color.Orange;
            group.SeriesColor = seriesColor;

            // 2. High and low point colors
            group.ShowHighPoint = true;
            group.ShowLowPoint = true;

            CellsColor highPointColor = workbook.CreateCellsColor();
            highPointColor.Color = Color.Green;
            group.HighPointColor = highPointColor;

            CellsColor lowPointColor = workbook.CreateCellsColor();
            lowPointColor.Color = Color.Red;
            group.LowPointColor = lowPointColor;

            // 3. First and last point colors
            group.ShowFirstPoint = true;
            group.ShowLastPoint = true;

            CellsColor firstPointColor = workbook.CreateCellsColor();
            firstPointColor.Color = Color.Purple;
            group.FirstPointColor = firstPointColor;

            CellsColor lastPointColor = workbook.CreateCellsColor();
            lastPointColor.Color = Color.Yellow;
            group.LastPointColor = lastPointColor;

            // 4. Markers (show each point) and marker color
            group.ShowMarkers = true;
            CellsColor markersColor = workbook.CreateCellsColor();
            markersColor.Color = Color.Black;
            group.MarkersColor = markersColor;

            // 5. Negative points color and visibility
            group.ShowNegativePoints = true;
            CellsColor negativeColor = workbook.CreateCellsColor();
            negativeColor.Color = Color.Blue;
            group.NegativePointsColor = negativeColor;

            // 6. Line weight (thickness)
            group.LineWeight = 1.5; // points

            // 7. Preset style (optional)
            group.PresetStyle = SparklinePresetStyleType.Style5;

            // 8. Horizontal axis visibility and color
            group.ShowHorizontalAxis = true;
            CellsColor hAxisColor = workbook.CreateCellsColor();
            hAxisColor.Color = Color.Gray;
            group.HorizontalAxisColor = hAxisColor;

            // 9. Plot empty cells as zero and include hidden cells
            group.PlotEmptyCellsType = PlotEmptyCellsType.Zero;
            group.DisplayHidden = true;

            // 10. Vertical axis range (custom)
            group.VerticalAxisMinValueType = SparklineAxisMinMaxType.Custom;
            group.VerticalAxisMaxValueType = SparklineAxisMinMaxType.Custom;
            group.VerticalAxisMinValue = 0.0;
            group.VerticalAxisMaxValue = 10.0;

            // Save the workbook with the customized sparkline
            workbook.Save("CustomizedSparkline.xlsx");
        }
    }
}