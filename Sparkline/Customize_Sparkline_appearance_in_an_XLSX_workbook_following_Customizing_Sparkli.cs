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
            // Create a new workbook
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
                StartColumn = 4, // Column E (0‑based index)
                EndColumn = 4
            };

            // Add a sparkline group of type Line with the data range A1:D1
            int groupIdx = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, location);
            SparklineGroup group = sheet.SparklineGroups[groupIdx];

            // Add a sparkline to the group (the sparkline will appear in cell E1)
            group.Sparklines.Add(sheet.Name + "!A1:D1", 0, 4);

            // ----- Customize appearance -----

            // Set series (line) color
            CellsColor seriesColor = workbook.CreateCellsColor();
            seriesColor.Color = Color.Orange;
            group.SeriesColor = seriesColor;

            // Highlight high and low points with custom colors
            group.ShowHighPoint = true;
            group.ShowLowPoint = true;

            CellsColor highColor = workbook.CreateCellsColor();
            highColor.Color = Color.Green;
            group.HighPointColor = highColor;

            CellsColor lowColor = workbook.CreateCellsColor();
            lowColor.Color = Color.Red;
            group.LowPointColor = lowColor;

            // Set line weight (thickness)
            group.LineWeight = 1.0;

            // Show markers for each data point and set marker color
            group.ShowMarkers = true;
            CellsColor markerColor = workbook.CreateCellsColor();
            markerColor.Color = Color.Black;
            group.MarkersColor = markerColor;

            // Apply a preset style (optional)
            group.PresetStyle = SparklinePresetStyleType.Style5;

            // Display data from hidden rows/columns
            group.DisplayHidden = true;

            // Highlight negative values with a custom color
            group.ShowNegativePoints = true;
            CellsColor negativeColor = workbook.CreateCellsColor();
            negativeColor.Color = Color.Blue;
            group.NegativePointsColor = negativeColor;

            // Show first and last points with custom colors
            group.ShowFirstPoint = true;
            CellsColor firstColor = workbook.CreateCellsColor();
            firstColor.Color = Color.Purple;
            group.FirstPointColor = firstColor;

            group.ShowLastPoint = true;
            CellsColor lastColor = workbook.CreateCellsColor();
            lastColor.Color = Color.Yellow;
            group.LastPointColor = lastColor;

            // Show horizontal axis and set its color
            group.ShowHorizontalAxis = true;
            CellsColor axisColor = workbook.CreateCellsColor();
            axisColor.Color = Color.Gray;
            group.HorizontalAxisColor = axisColor;

            // Set custom vertical axis range (optional)
            group.VerticalAxisMinValueType = SparklineAxisMinMaxType.Group;
            group.VerticalAxisMinValue = 0.0;
            group.VerticalAxisMaxValueType = SparklineAxisMinMaxType.Group;
            group.VerticalAxisMaxValue = 6.0;

            // Save the workbook
            workbook.Save("CustomSparklineAppearance.xlsx", SaveFormat.Xlsx);
        }
    }
}