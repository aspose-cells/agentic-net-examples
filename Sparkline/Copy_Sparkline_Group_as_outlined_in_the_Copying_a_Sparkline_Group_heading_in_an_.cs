using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsSparklineCopyDemo
{
    class Program
    {
        static void Main()
        {
            // Load the source workbook (replace with your actual file path)
            Workbook workbook = new Workbook("source.xlsx");

            // Get the worksheet that contains the sparkline group to be copied
            Worksheet srcSheet = workbook.Worksheets[0];

            // Ensure there is at least one sparkline group
            if (srcSheet.SparklineGroups.Count == 0)
            {
                Console.WriteLine("No sparkline groups found in the source worksheet.");
                return;
            }

            // Get the first sparkline group (source group)
            SparklineGroup srcGroup = srcSheet.SparklineGroups[0];

            // Retrieve the data range string from the first sparkline in the group
            // (All sparklines in a group share the same data range pattern)
            Sparkline firstSparkline = srcGroup.Sparklines[0];
            string dataRange = firstSparkline.DataRange; // e.g., "Sheet1!A1:D1"

            // Determine orientation (vertical/horizontal) – assume false (by column) for this demo
            bool isVertical = false;

            // Compute a new location range for the copied group.
            // Here we shift the original location one column to the right.
            // The original location is taken from the first sparkline's row/column.
            int srcRow = firstSparkline.Row;
            int srcCol = firstSparkline.Column;

            // Define the new location CellArea (single cell)
            CellArea newLocation = new CellArea
            {
                StartRow = srcRow,
                EndRow = srcRow,
                StartColumn = srcCol + 1, // shift one column right
                EndColumn = srcCol + 1
            };

            // Add a new sparkline group to the same worksheet with the same type and data range
            int newGroupIndex = srcSheet.SparklineGroups.Add(srcGroup.Type, dataRange, isVertical, newLocation);
            SparklineGroup newGroup = srcSheet.SparklineGroups[newGroupIndex];

            // Copy visual properties from the source group to the new group
            newGroup.SeriesColor = srcGroup.SeriesColor;
            newGroup.HighPointColor = srcGroup.HighPointColor;
            newGroup.LowPointColor = srcGroup.LowPointColor;
            newGroup.FirstPointColor = srcGroup.FirstPointColor;
            newGroup.LastPointColor = srcGroup.LastPointColor;
            newGroup.MarkersColor = srcGroup.MarkersColor;
            newGroup.NegativePointsColor = srcGroup.NegativePointsColor;
            newGroup.HorizontalAxisColor = srcGroup.HorizontalAxisColor;

            newGroup.ShowHighPoint = srcGroup.ShowHighPoint;
            newGroup.ShowLowPoint = srcGroup.ShowLowPoint;
            newGroup.ShowFirstPoint = srcGroup.ShowFirstPoint;
            newGroup.ShowLastPoint = srcGroup.ShowLastPoint;
            newGroup.ShowMarkers = srcGroup.ShowMarkers;
            newGroup.ShowNegativePoints = srcGroup.ShowNegativePoints;
            newGroup.ShowHorizontalAxis = srcGroup.ShowHorizontalAxis;
            newGroup.PlotRightToLeft = srcGroup.PlotRightToLeft;
            newGroup.LineWeight = srcGroup.LineWeight;
            newGroup.PresetStyle = srcGroup.PresetStyle;
            newGroup.PlotEmptyCellsType = srcGroup.PlotEmptyCellsType;
            newGroup.DisplayHidden = srcGroup.DisplayHidden;
            newGroup.VerticalAxisMaxValue = srcGroup.VerticalAxisMaxValue;
            newGroup.VerticalAxisMinValue = srcGroup.VerticalAxisMinValue;
            newGroup.VerticalAxisMaxValueType = srcGroup.VerticalAxisMaxValueType;
            newGroup.VerticalAxisMinValueType = srcGroup.VerticalAxisMinValueType;
            newGroup.HorizontalAxisDateRange = srcGroup.HorizontalAxisDateRange;

            // Save the workbook with the copied sparkline group
            workbook.Save("output_copied_sparklines.xlsx");
        }
    }
}