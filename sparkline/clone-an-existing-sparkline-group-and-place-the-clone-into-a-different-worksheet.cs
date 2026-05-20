using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace SparklineCloneDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Source worksheet (contains the original sparkline group)
                Worksheet sourceSheet = workbook.Worksheets[0];
                sourceSheet.Name = "Source";

                // Target worksheet (where the cloned sparkline group will be placed)
                Worksheet targetSheet = workbook.Worksheets.Add("Target");

                // -------------------------------------------------
                // Populate sample data in the source worksheet
                // -------------------------------------------------
                for (int i = 0; i < 5; i++)
                {
                    // Fill column A with values 1..5
                    sourceSheet.Cells[i, 0].PutValue(i + 1);
                }

                // -------------------------------------------------
                // Add a sparkline group to the source worksheet
                // -------------------------------------------------
                // Define where the sparkline will be displayed (cell B1)
                CellArea sourceLocation = new CellArea
                {
                    StartRow = 0,
                    EndRow = 0,
                    StartColumn = 1,
                    EndColumn = 1
                };

                // Use a fully‑qualified data range (includes sheet name) to avoid reference errors
                string sourceDataRange = $"{sourceSheet.Name}!A1:A5";

                // Add the sparkline group (horizontal line type)
                int sourceGroupIndex = sourceSheet.SparklineGroups.Add(
                    SparklineType.Line,   // type
                    sourceDataRange,      // data range (with sheet name)
                    false,                // isVertical
                    sourceLocation);      // location range

                SparklineGroup sourceGroup = sourceSheet.SparklineGroups[sourceGroupIndex];

                // Add a sparkline item to the group
                sourceGroup.Sparklines.Add(sourceDataRange, 0, 1);

                // Optional: customize the source sparkline group
                sourceGroup.ShowHighPoint = true;
                sourceGroup.ShowLowPoint = true;
                sourceGroup.LineWeight = 1.0;

                // -------------------------------------------------
                // Clone the sparkline group into the target worksheet
                // -------------------------------------------------
                // Retrieve the data range string from the first sparkline (includes sheet name)
                string dataRange = sourceGroup.Sparklines[0].DataRange;

                // Define the location for the cloned sparkline in the target sheet (cell B1)
                CellArea targetLocation = new CellArea
                {
                    StartRow = 0,
                    EndRow = 0,
                    StartColumn = 1,
                    EndColumn = 1
                };

                // Add a new sparkline group to the target sheet using the same type and data range
                int targetGroupIndex = targetSheet.SparklineGroups.Add(
                    sourceGroup.Type,   // same sparkline type as the source
                    dataRange,          // same data range (refers to the source sheet)
                    false,              // isVertical (same as source)
                    targetLocation);    // location in the target sheet

                SparklineGroup targetGroup = targetSheet.SparklineGroups[targetGroupIndex];

                // Add a sparkline item to the cloned group (same data range)
                targetGroup.Sparklines.Add(dataRange, 0, 1);

                // Copy visual properties from the source group to the cloned group
                targetGroup.ShowHighPoint = sourceGroup.ShowHighPoint;
                targetGroup.ShowLowPoint = sourceGroup.ShowLowPoint;
                targetGroup.LineWeight = sourceGroup.LineWeight;
                targetGroup.SeriesColor = sourceGroup.SeriesColor;
                targetGroup.HighPointColor = sourceGroup.HighPointColor;
                targetGroup.LowPointColor = sourceGroup.LowPointColor;

                // -------------------------------------------------
                // Save the workbook
                // -------------------------------------------------
                workbook.Save("ClonedSparklineGroup.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}