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
            int groupIdx = sheet.SparklineGroups.Add(SparklineType.Line, "A1:E1", false, location);
            SparklineGroup group = sheet.SparklineGroups[groupIdx];

            // Add the sparkline to the group (optional, Add already created one)
            group.Sparklines.Add(sheet.Name + "!A1:E1", 0, 5);

            // -------------------- Customize Appearance --------------------

            // Apply a preset style (optional, can be overridden by individual settings)
            group.PresetStyle = SparklinePresetStyleType.Style5;

            // Set the main series color
            CellsColor seriesColor = workbook.CreateCellsColor();
            seriesColor.Color = Color.Orange;
            group.SeriesColor = seriesColor;

            // Enable and color markers for each point
            group.ShowMarkers = true;
            CellsColor markersColor = workbook.CreateCellsColor();
            markersColor.Color = Color.Black;
            group.MarkersColor = markersColor;

            // Highlight first and last points with custom colors
            group.ShowFirstPoint = true;
            CellsColor firstPointColor = workbook.CreateCellsColor();
            firstPointColor.Color = Color.Purple;
            group.FirstPointColor = firstPointColor;

            group.ShowLastPoint = true;
            CellsColor lastPointColor = workbook.CreateCellsColor();
            lastPointColor.Color = Color.Yellow;
            group.LastPointColor = lastPointColor;

            // Highlight high and low points with custom colors
            group.ShowHighPoint = true;
            CellsColor highPointColor = workbook.CreateCellsColor();
            highPointColor.Color = Color.Green;
            group.HighPointColor = highPointColor;

            group.ShowLowPoint = true;
            CellsColor lowPointColor = workbook.CreateCellsColor();
            lowPointColor.Color = Color.Red;
            group.LowPointColor = lowPointColor;

            // Highlight negative values with a distinct color
            group.ShowNegativePoints = true;
            CellsColor negativePointsColor = workbook.CreateCellsColor();
            negativePointsColor.Color = Color.Blue;
            group.NegativePointsColor = negativePointsColor;

            // Show horizontal axis and set its color
            group.ShowHorizontalAxis = true;
            CellsColor horizontalAxisColor = workbook.CreateCellsColor();
            horizontalAxisColor.Color = Color.Gray;
            group.HorizontalAxisColor = horizontalAxisColor;

            // Set line weight for the sparkline
            group.LineWeight = 1.0;

            // Save the workbook to an XLSX file
            workbook.Save("CustomizedSparkline.xlsx");
        }
    }
}