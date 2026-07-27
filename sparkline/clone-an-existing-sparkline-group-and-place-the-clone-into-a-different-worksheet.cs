using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace SparklineCloneDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the source worksheet
            Workbook workbook = new Workbook();
            Worksheet sourceSheet = workbook.Worksheets[0];
            sourceSheet.Name = "Source";

            // Populate sample data for the sparkline
            sourceSheet.Cells["A1"].PutValue(5);
            sourceSheet.Cells["B1"].PutValue(2);
            sourceSheet.Cells["C1"].PutValue(1);
            sourceSheet.Cells["D1"].PutValue(3);

            // Define the location where the sparkline will be placed
            CellArea sourceLocation = new CellArea
            {
                StartRow = 0,
                EndRow = 0,
                StartColumn = 4,
                EndColumn = 4
            };

            // Add a sparkline group to the source sheet
            int sourceGroupIndex = sourceSheet.SparklineGroups.Add(
                SparklineType.Line,          // type
                "A1:D1",                     // data range
                false,                       // isVertical
                sourceLocation);             // location range

            SparklineGroup sourceGroup = sourceSheet.SparklineGroups[sourceGroupIndex];

            // Add a sparkline item to the group (required to have at least one)
            sourceGroup.Sparklines.Add(sourceSheet.Name + "!A1:D1", 0, 4);

            // Customize some visual properties (optional, will be cloned later)
            CellsColor seriesColor = workbook.CreateCellsColor();
            seriesColor.Color = Color.Orange;
            sourceGroup.SeriesColor = seriesColor;
            sourceGroup.ShowHighPoint = true;
            sourceGroup.ShowLowPoint = true;
            CellsColor highColor = workbook.CreateCellsColor();
            highColor.Color = Color.Green;
            sourceGroup.HighPointColor = highColor;
            CellsColor lowColor = workbook.CreateCellsColor();
            lowColor.Color = Color.Red;
            sourceGroup.LowPointColor = lowColor;

            // -------------------------------------------------
            // Clone the sparkline group into a new worksheet
            // -------------------------------------------------

            // Add a new worksheet that will receive the cloned group
            Worksheet targetSheet = workbook.Worksheets.Add("Clone");
            
            // Define the location for the cloned sparkline group in the target sheet
            // Here we use the same cell coordinates as the source for simplicity
            CellArea targetLocation = new CellArea
            {
                StartRow = sourceLocation.StartRow,
                EndRow = sourceLocation.EndRow,
                StartColumn = sourceLocation.StartColumn,
                EndColumn = sourceLocation.EndColumn
            };

            // Create a new sparkline group in the target sheet with the same type,
            // data range, orientation and location as the source group
            int targetGroupIndex = targetSheet.SparklineGroups.Add(
                sourceGroup.Type,            // same sparkline type
                "A1:D1",                     // same data range (relative to target sheet)
                false,                       // same orientation
                targetLocation);             // location in target sheet

            SparklineGroup targetGroup = targetSheet.SparklineGroups[targetGroupIndex];

            // Copy visual properties from the source group to the target group
            targetGroup.SeriesColor = sourceGroup.SeriesColor;
            targetGroup.ShowHighPoint = sourceGroup.ShowHighPoint;
            targetGroup.ShowLowPoint = sourceGroup.ShowLowPoint;
            targetGroup.HighPointColor = sourceGroup.HighPointColor;
            targetGroup.LowPointColor = sourceGroup.LowPointColor;
            targetGroup.LineWeight = sourceGroup.LineWeight;
            targetGroup.PresetStyle = sourceGroup.PresetStyle;
            targetGroup.PlotEmptyCellsType = sourceGroup.PlotEmptyCellsType;
            targetGroup.DisplayHidden = sourceGroup.DisplayHidden;
            targetGroup.ShowNegativePoints = sourceGroup.ShowNegativePoints;
            targetGroup.NegativePointsColor = sourceGroup.NegativePointsColor;
            targetGroup.ShowFirstPoint = sourceGroup.ShowFirstPoint;
            targetGroup.FirstPointColor = sourceGroup.FirstPointColor;
            targetGroup.ShowLastPoint = sourceGroup.ShowLastPoint;
            targetGroup.LastPointColor = sourceGroup.LastPointColor;
            targetGroup.ShowMarkers = sourceGroup.ShowMarkers;
            targetGroup.MarkersColor = sourceGroup.MarkersColor;
            targetGroup.ShowHorizontalAxis = sourceGroup.ShowHorizontalAxis;
            targetGroup.HorizontalAxisColor = sourceGroup.HorizontalAxisColor;
            targetGroup.HorizontalAxisDateRange = sourceGroup.HorizontalAxisDateRange;
            targetGroup.VerticalAxisMaxValue = sourceGroup.VerticalAxisMaxValue;
            targetGroup.VerticalAxisMaxValueType = sourceGroup.VerticalAxisMaxValueType;
            targetGroup.VerticalAxisMinValue = sourceGroup.VerticalAxisMinValue;
            targetGroup.VerticalAxisMinValueType = sourceGroup.VerticalAxisMinValueType;
            targetGroup.PlotRightToLeft = sourceGroup.PlotRightToLeft;

            // Clone each sparkline item from the source group to the target group
            foreach (Sparkline spark in sourceGroup.Sparklines)
            {
                // The DataRange string can be used directly; it refers to the source sheet name,
                // so replace it with the target sheet name to point to the same cells in the target sheet.
                string clonedDataRange = spark.DataRange.Replace(sourceSheet.Name, targetSheet.Name);
                targetGroup.Sparklines.Add(clonedDataRange, spark.Row, spark.Column);
            }

            // Save the workbook with both the original and cloned sparkline groups
            workbook.Save("SparklineGroupCloneDemo.xlsx");
        }
    }
}