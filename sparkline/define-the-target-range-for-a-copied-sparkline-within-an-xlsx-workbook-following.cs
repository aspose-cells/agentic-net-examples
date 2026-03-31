using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace SparklineCopyDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate source data for the sparkline (row 0, columns A‑D)
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["B1"].PutValue(2);
            sheet.Cells["C1"].PutValue(1);
            sheet.Cells["D1"].PutValue(3);

            // Define the original location range for the sparkline (cell E1)
            CellArea originalLocation = new CellArea
            {
                StartRow = 0,
                EndRow = 0,
                StartColumn = 4,   // Column E (zero‑based index)
                EndColumn = 4
            };

            // Add a sparkline group with the source data and original location
            int groupIdx = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, originalLocation);
            SparklineGroup sparklineGroup = sheet.SparklineGroups[groupIdx];

            // -----------------------------------------------------------------
            // Define the target range where the sparkline will be copied.
            // According to the guidelines, the target location must have the
            // same number of columns as the source data range when isVertical
            // is false. Here we place the copied sparkline in cell G1 (column 6).
            // -----------------------------------------------------------------
            CellArea targetLocation = CellArea.CreateCellArea(0, 6, 0, 6); // Row 0, Column G

            // Reset the ranges of the sparkline group to point to the same data
            // but relocate the sparkline to the target location.
            sparklineGroup.ResetRanges("A1:D1", false, targetLocation);

            // Save the workbook with the copied sparkline
            workbook.Save("SparklineCopyTargetRange.xlsx");
        }
    }
}