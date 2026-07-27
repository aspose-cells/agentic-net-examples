using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

class CopySparklineGroupExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // -------------------------------------------------
        // Prepare Sheet1 with sample data and a sparkline group
        // -------------------------------------------------
        Worksheet sheet1 = workbook.Worksheets[0];
        sheet1.Name = "Sheet1";

        // Fill some sample data (A1:D1)
        sheet1.Cells["A1"].PutValue(5);
        sheet1.Cells["B1"].PutValue(2);
        sheet1.Cells["C1"].PutValue(1);
        sheet1.Cells["D1"].PutValue(3);

        // Define where the sparkline will be placed (E1)
        CellArea location = CellArea.CreateCellArea("E1", "E1");

        // Add a sparkline group to Sheet1
        int originalGroupIndex = sheet1.SparklineGroups.Add(
            SparklineType.Line,          // type
            "A1:D1",                     // data range
            false,                       // isVertical
            location);                   // location range

        SparklineGroup originalGroup = sheet1.SparklineGroups[originalGroupIndex];

        // (Optional) Add an extra sparkline to demonstrate multiple sparklines in a group
        // Here we add the same data range to another cell (F1)
        originalGroup.Sparklines.Add("A1:D1", 0, 5); // row 0, column 5 => F1

        // Customize some formatting on the original group
        originalGroup.ShowHighPoint = true;
        originalGroup.ShowLowPoint = true;
        originalGroup.HighPointColor.Color = Color.Green;
        originalGroup.LowPointColor.Color = Color.Red;
        originalGroup.SeriesColor.Color = Color.Blue;
        originalGroup.LineWeight = 1.0;

        // -------------------------------------------------
        // Create Sheet2 where the sparkline group will be copied
        // -------------------------------------------------
        Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");

        // Copy the same data to Sheet2 (so the sparkline references are valid)
        sheet2.Cells["A1"].PutValue(5);
        sheet2.Cells["B1"].PutValue(2);
        sheet2.Cells["C1"].PutValue(1);
        sheet2.Cells["D1"].PutValue(3);

        // -------------------------------------------------
        // Replicate the sparkline group on Sheet2
        // -------------------------------------------------
        // Add a new group with the same type, data range, orientation and location range
        int copiedGroupIndex = sheet2.SparklineGroups.Add(
            originalGroup.Type,          // same sparkline type
            "A1:D1",                     // same data range
            false,                       // same orientation (assumed false)
            location);                   // same location range (E1 on Sheet2)

        SparklineGroup copiedGroup = sheet2.SparklineGroups[copiedGroupIndex];

        // Copy each sparkline from the original group to the new group
        foreach (Sparkline sp in originalGroup.Sparklines)
        {
            // Add sparkline with identical data range and cell location
            copiedGroup.Sparklines.Add(sp.DataRange, sp.Row, sp.Column);
        }

        // Copy formatting properties
        copiedGroup.ShowHighPoint = originalGroup.ShowHighPoint;
        copiedGroup.ShowLowPoint = originalGroup.ShowLowPoint;
        copiedGroup.HighPointColor.Color = originalGroup.HighPointColor.Color;
        copiedGroup.LowPointColor.Color = originalGroup.LowPointColor.Color;
        copiedGroup.SeriesColor.Color = originalGroup.SeriesColor.Color;
        copiedGroup.LineWeight = originalGroup.LineWeight;

        // -------------------------------------------------
        // Save the workbook
        // -------------------------------------------------
        workbook.Save("SparklineGroupCopyDemo.xlsx");
    }
}