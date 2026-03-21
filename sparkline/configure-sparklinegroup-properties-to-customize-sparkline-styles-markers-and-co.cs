using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace SparklineCustomizationDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the sparkline
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["B1"].PutValue(2);
            sheet.Cells["C1"].PutValue(8);
            sheet.Cells["D1"].PutValue(3);
            sheet.Cells["E1"].PutValue(7);

            // Define the location where the sparkline will be placed
            CellArea location = new CellArea
            {
                StartRow = 0,
                EndRow = 0,
                StartColumn = 5, // column F
                EndColumn = 5
            };

            // Add a sparkline group (type: Line) with the data range and location
            int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:E1", false, location);
            SparklineGroup group = sheet.SparklineGroups[groupIndex];

            // Add the sparkline to the group (optional when using Add overload, but shown for clarity)
            group.Sparklines.Add(sheet.Name + "!A1:E1", 0, 5);

            // ----- Customize sparkline appearance -----

            // Series (line) color
            CellsColor seriesColor = workbook.CreateCellsColor();
            seriesColor.Color = Color.Orange;
            group.SeriesColor = seriesColor;

            // Line weight
            group.LineWeight = 1.2;

            // Preset style (optional, overrides many individual settings if not Custom)
            group.PresetStyle = SparklinePresetStyleType.Custom;

            // Show and color the first point
            group.ShowFirstPoint = true;
            CellsColor firstPointColor = workbook.CreateCellsColor();
            firstPointColor.Color = Color.Purple;
            group.FirstPointColor = firstPointColor;

            // Show and color the last point
            group.ShowLastPoint = true;
            CellsColor lastPointColor = workbook.CreateCellsColor();
            lastPointColor.Color = Color.Yellow;
            group.LastPointColor = lastPointColor;

            // Show and color the highest point
            group.ShowHighPoint = true;
            CellsColor highPointColor = workbook.CreateCellsColor();
            highPointColor.Color = Color.Green;
            group.HighPointColor = highPointColor;

            // Show and color the lowest point
            group.ShowLowPoint = true;
            CellsColor lowPointColor = workbook.CreateCellsColor();
            lowPointColor.Color = Color.Red;
            group.LowPointColor = lowPointColor;

            // Show markers for each data point and set their color
            group.ShowMarkers = true;
            CellsColor markersColor = workbook.CreateCellsColor();
            markersColor.Color = Color.Black;
            group.MarkersColor = markersColor;

            // Highlight negative values with a distinct color
            group.ShowNegativePoints = true;
            CellsColor negativePointsColor = workbook.CreateCellsColor();
            negativePointsColor.Color = Color.Blue;
            group.NegativePointsColor = negativePointsColor;

            // Plot data from right to left
            group.PlotRightToLeft = false;

            // Display the horizontal axis and set its color
            group.ShowHorizontalAxis = true;
            CellsColor horizontalAxisColor = workbook.CreateCellsColor();
            horizontalAxisColor.Color = Color.Gray;
            group.HorizontalAxisColor = horizontalAxisColor;

            // Set vertical axis to use the same range for all sparklines in the group
            group.VerticalAxisMaxValueType = SparklineAxisMinMaxType.Group;
            group.VerticalAxisMinValueType = SparklineAxisMinMaxType.Group;
            group.VerticalAxisMaxValue = 10.0;
            group.VerticalAxisMinValue = 0.0;

            // Save the workbook (lifecycle: save)
            workbook.Save("CustomizedSparkline.xlsx", SaveFormat.Xlsx);
        }
    }
}