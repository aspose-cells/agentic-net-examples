// Title: Copy a Sparkline Group from one worksheet to another while preserving data ranges and formatting with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# copy sparkline group to another worksheet preserving colors | How to duplicate a sparkline group across sheets using Aspose.Cells .NET | Copy sparkline group with absolute data range reference in C# Aspose.Cells | Replicate sparkline formatting when moving between worksheets in Aspose.Cells | Clone CellsColor objects while copying sparkline groups in Aspose.Cells C#
// Tags: sparkline group copy worksheet Aspose.Cells | sparkline formatting preservation C# | absolute sparkline data range Aspose.Cells | cellscolor cloning Aspose.Cells | programmatic sparkline creation Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds sample data, defines a line sparkline group on Sheet1 with custom colors and point markers, then copies the entire group to Sheet2 using an absolute data range reference and replicates all formatting properties by cloning each CellsColor. The resulting workbook is saved as SparklineGroupCopy.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // Prepare worksheets.
            // -------------------------------------------------
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Sheet1";

            int sheet2Idx = workbook.Worksheets.Add();
            Worksheet sheet2 = workbook.Worksheets[sheet2Idx];
            sheet2.Name = "Sheet2";

            // -------------------------------------------------
            // Fill sample data in Sheet1 (range A1:D1).
            // -------------------------------------------------
            sheet1.Cells["A1"].PutValue(5);
            sheet1.Cells["B1"].PutValue(2);
            sheet1.Cells["C1"].PutValue(1);
            sheet1.Cells["D1"].PutValue(3);

            // -------------------------------------------------
            // Create a sparkline group on Sheet1.
            // -------------------------------------------------
            // Location where the sparkline will be displayed (cell E1).
            CellArea location1 = new CellArea
            {
                StartRow = 0,
                EndRow = 0,
                StartColumn = 4,
                EndColumn = 4
            };

            // Add the sparkline group.
            int groupIdx1 = sheet1.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, location1);
            SparklineGroup group1 = sheet1.SparklineGroups[groupIdx1];

            // Add a sparkline to the group.
            group1.Sparklines.Add("A1:D1", 0, 4);

            // Apply some formatting to the group.
            group1.ShowHighPoint = true;
            group1.ShowLowPoint = true;
            group1.ShowFirstPoint = true;

            CellsColor highColor = workbook.CreateCellsColor();
            highColor.Color = Color.Green;
            group1.HighPointColor = highColor;

            CellsColor lowColor = workbook.CreateCellsColor();
            lowColor.Color = Color.Red;
            group1.LowPointColor = lowColor;

            CellsColor seriesColor = workbook.CreateCellsColor();
            seriesColor.Color = Color.Orange;
            group1.SeriesColor = seriesColor;

            group1.LineWeight = 1.0;

            // -------------------------------------------------
            // Copy the entire sparkline group to Sheet2.
            // -------------------------------------------------
            // Location on Sheet2 where the copied sparkline will appear (cell E1).
            CellArea location2 = new CellArea
            {
                StartRow = 0,
                EndRow = 0,
                StartColumn = 4,
                EndColumn = 4
            };

            // Data range must be absolute because the source data stays on Sheet1.
            string absoluteDataRange = $"{sheet1.Name}!A1:D1";

            // Add a new sparkline group on Sheet2 with the same type.
            int groupIdx2 = sheet2.SparklineGroups.Add(group1.Type, absoluteDataRange, false, location2);
            SparklineGroup group2 = sheet2.SparklineGroups[groupIdx2];

            // Copy each sparkline (normally one) from the source group.
            foreach (Sparkline srcSparkline in group1.Sparklines)
            {
                group2.Sparklines.Add(absoluteDataRange, srcSparkline.Row, srcSparkline.Column);
            }

            // Copy formatting properties.
            group2.ShowHighPoint = group1.ShowHighPoint;
            group2.ShowLowPoint = group1.ShowLowPoint;
            group2.ShowFirstPoint = group1.ShowFirstPoint;
            group2.LineWeight = group1.LineWeight;
            group2.PresetStyle = group1.PresetStyle;
            group2.PlotEmptyCellsType = group1.PlotEmptyCellsType;
            group2.PlotRightToLeft = group1.PlotRightToLeft;
            group2.DisplayHidden = group1.DisplayHidden;

            // Copy colors (clone to avoid sharing the same instance).
            group2.HighPointColor = CloneColor(workbook, group1.HighPointColor);
            group2.LowPointColor = CloneColor(workbook, group1.LowPointColor);
            group2.FirstPointColor = CloneColor(workbook, group1.FirstPointColor);
            group2.LastPointColor = CloneColor(workbook, group1.LastPointColor);
            group2.SeriesColor = CloneColor(workbook, group1.SeriesColor);
            group2.MarkersColor = CloneColor(workbook, group1.MarkersColor);
            group2.NegativePointsColor = CloneColor(workbook, group1.NegativePointsColor);
            group2.HorizontalAxisColor = CloneColor(workbook, group1.HorizontalAxisColor);

            // -------------------------------------------------
            // Save the workbook.
            // -------------------------------------------------
            string outputPath = "SparklineGroupCopy.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }

    // Helper to clone a CellsColor (handles null values).
    private static CellsColor CloneColor(Workbook wb, CellsColor src)
    {
        if (src == null) return null;
        CellsColor dest = wb.CreateCellsColor();
        dest.Color = src.Color;
        return dest;
    }
}
