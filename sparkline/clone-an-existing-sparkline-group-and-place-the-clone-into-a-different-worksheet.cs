using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace SparklineCloneDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and obtain the source worksheet
            Workbook workbook = new Workbook();
            Worksheet sourceSheet = workbook.Worksheets[0];
            sourceSheet.Name = "Source";

            // Populate sample data in the source sheet
            for (int i = 0; i < 5; i++)
                sourceSheet.Cells[i, 0].PutValue(i + 1); // A1:A5

            // Define the location where the sparkline will be placed (B1:B5)
            CellArea sourceLocation = new CellArea
            {
                StartRow = 0,
                EndRow = 4,
                StartColumn = 1,
                EndColumn = 1
            };

            // Add a sparkline group to the source sheet
            int srcGroupIdx = sourceSheet.SparklineGroups.Add(
                SparklineType.Line,
                $"{sourceSheet.Name}!A1:A5",
                false,
                sourceLocation);
            SparklineGroup srcGroup = sourceSheet.SparklineGroups[srcGroupIdx];

            // Add a sparkline item to the group (optional, as the group already creates sparklines for the range)
            srcGroup.Sparklines.Add($"{sourceSheet.Name}!A1:A5", 0, 1);

            // OPTIONAL: customize the source group (demonstration purposes)
            srcGroup.ShowHighPoint = true;
            srcGroup.ShowLowPoint = true;

            // ------------------------------------------------------------
            // Create a target worksheet where the cloned sparkline group will reside
            Worksheet targetSheet = workbook.Worksheets.Add("Target");

            // Copy the source data to the target sheet so the sparkline can reference it
            for (int i = 0; i < 5; i++)
                targetSheet.Cells[i, 0].PutValue(sourceSheet.Cells[i, 0].Value);

            // Define the location for the cloned sparkline (B1:B5 in the target sheet)
            CellArea targetLocation = new CellArea
            {
                StartRow = 0,
                EndRow = 4,
                StartColumn = 1,
                EndColumn = 1
            };

            // Clone the sparkline group:
            // 1. Add a new group with the same type, data range (adjusted to target sheet), orientation and location.
            int tgtGroupIdx = targetSheet.SparklineGroups.Add(
                srcGroup.Type,
                $"{targetSheet.Name}!A1:A5",   // data range now points to target sheet
                false,
                targetLocation);
            SparklineGroup tgtGroup = targetSheet.SparklineGroups[tgtGroupIdx];

            // 2. Copy each sparkline item from the source group to the target group.
            foreach (Sparkline sp in srcGroup.Sparklines)
            {
                // Adjust the data range to refer to the target sheet.
                string newDataRange = sp.DataRange.Replace(sourceSheet.Name, targetSheet.Name);
                tgtGroup.Sparklines.Add(newDataRange, sp.Row, sp.Column);
            }

            // 3. Replicate visual properties (example: show high/low points).
            tgtGroup.ShowHighPoint = srcGroup.ShowHighPoint;
            tgtGroup.ShowLowPoint = srcGroup.ShowLowPoint;
            tgtGroup.SeriesColor = srcGroup.SeriesColor;
            tgtGroup.HighPointColor = srcGroup.HighPointColor;
            tgtGroup.LowPointColor = srcGroup.LowPointColor;
            tgtGroup.LineWeight = srcGroup.LineWeight;

            // Save the workbook with both original and cloned sparkline groups
            workbook.Save("ClonedSparklineGroup.xlsx");
        }
    }
}