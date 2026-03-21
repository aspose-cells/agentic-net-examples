using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace SparklineConfigurationDemo
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

            // Add a sparkline group of type Line with the data range A1:D1
            int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, location);
            SparklineGroup group = sheet.SparklineGroups[groupIndex];

            // Add the sparkline to the group (the Add method of Sparklines collection)
            group.Sparklines.Add(sheet.Name + "!A1:D1", 0, 4);

            // ----- Style configuration -----
            // Apply a preset style (Style5)
            group.PresetStyle = SparklinePresetStyleType.Style5;

            // Set the line weight (thicker line for better visibility)
            group.LineWeight = 2.0;

            // Plot empty cells as zero and display hidden rows/columns
            group.PlotEmptyCellsType = PlotEmptyCellsType.Zero;
            group.DisplayHidden = true;

            // ----- Series color -----
            CellsColor seriesColor = workbook.CreateCellsColor();
            seriesColor.Color = Color.Orange;
            group.SeriesColor = seriesColor;

            // ----- Markers -----
            group.ShowMarkers = true;
            CellsColor markersColor = workbook.CreateCellsColor();
            markersColor.Color = Color.Black;
            group.MarkersColor = markersColor;

            // ----- High / Low points -----
            group.ShowHighPoint = true;
            CellsColor highPointColor = workbook.CreateCellsColor();
            highPointColor.Color = Color.Green;
            group.HighPointColor = highPointColor;

            group.ShowLowPoint = true;
            CellsColor lowPointColor = workbook.CreateCellsColor();
            lowPointColor.Color = Color.Red;
            group.LowPointColor = lowPointColor;

            // ----- First / Last points -----
            group.ShowFirstPoint = true;
            CellsColor firstPointColor = workbook.CreateCellsColor();
            firstPointColor.Color = Color.Purple;
            group.FirstPointColor = firstPointColor;

            group.ShowLastPoint = true;
            CellsColor lastPointColor = workbook.CreateCellsColor();
            lastPointColor.Color = Color.Yellow;
            group.LastPointColor = lastPointColor;

            // ----- Negative points -----
            group.ShowNegativePoints = true;
            CellsColor negativePointsColor = workbook.CreateCellsColor();
            negativePointsColor.Color = Color.Blue;
            group.NegativePointsColor = negativePointsColor;

            // ----- Horizontal axis -----
            group.ShowHorizontalAxis = true;
            CellsColor horizontalAxisColor = workbook.CreateCellsColor();
            horizontalAxisColor.Color = Color.Gray;
            group.HorizontalAxisColor = horizontalAxisColor;
            group.HorizontalAxisDateRange = "A1:D1";

            // ----- Vertical axis scaling (group-wide) -----
            group.VerticalAxisMaxValueType = SparklineAxisMinMaxType.Group;
            group.VerticalAxisMaxValue = 10.0;
            group.VerticalAxisMinValueType = SparklineAxisMinMaxType.Group;
            group.VerticalAxisMinValue = 0.0;

            // Save the workbook to a file
            workbook.Save("ConfiguredSparklineGroup.xlsx");
        }
    }
}