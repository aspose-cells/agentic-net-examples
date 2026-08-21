// Title: Clone a SparklineGroup to Another Worksheet with Aspose.Cells for .NET (C#)
// Description: This C# example shows how to duplicate an existing SparklineGroup, preserve its visual settings (high‑point, low‑point, colors, line weight), and place the clone on a different worksheet using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | SparklineGroup | clone sparkline | copy sparkline group | worksheet | preserve formatting | line sparkline | duplicate sparklines | Aspose.Cells Sparkline
// Common Searches: Aspose.Cells clone sparkline group C# | copy sparkline group to another worksheet Aspose.Cells | duplicate sparklines across sheets .NET | preserve sparkline formatting Aspose.Cells | how to replicate SparklineGroup in C#
// Developer Intent: Duplicate a SparklineGroup from one worksheet to another while keeping its data range and visual properties intact.
// Use Cases: Generate a summary sheet that mirrors sparklines from a source sheet with identical markers and colors. | Apply a consistent sparkline style across multiple report tabs without manual re‑creation. | Automate batch reporting where the same sparkline visualisation is needed on several worksheets.
// AI Prompts: Write C# code using Aspose.Cells to clone a SparklineGroup from a source worksheet to a destination worksheet, copying all visual settings and adjusting the location. | Create a method that accepts a source SparklineGroup, a target Worksheet, and a CellArea, then replicates the group with correct row/column offsets. | Show how to transfer high‑point, low‑point, series color, and line weight of a sparkline group to another sheet with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace SparklineCloneDemo
{
    // This C# example shows how to duplicate an existing SparklineGroup, preserve its visual settings (high‑point, low‑point, colors, line weight), and place the clone on a different worksheet using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // ---------- Create a workbook and populate source worksheet ----------
            Workbook workbook = new Workbook();
            Worksheet srcSheet = workbook.Worksheets[0];
            srcSheet.Name = "Source";

            // Sample data for sparklines
            srcSheet.Cells["A1"].PutValue(5);
            srcSheet.Cells["A2"].PutValue(3);
            srcSheet.Cells["A3"].PutValue(7);
            srcSheet.Cells["A4"].PutValue(2);
            srcSheet.Cells["A5"].PutValue(9);

            // Define where the sparkline will be placed (column F, rows 1‑5)
            CellArea srcLocation = new CellArea
            {
                StartRow = 0,
                EndRow = 4,
                StartColumn = 5,
                EndColumn = 5
            };

            // Add a sparkline group to the source sheet
            int srcGroupIdx = srcSheet.SparklineGroups.Add(
                SparklineType.Line,
                srcSheet.Name + "!A1:A5",   // data range
                false,                      // horizontal (by row)
                srcLocation);               // location range

            SparklineGroup srcGroup = srcSheet.SparklineGroups[srcGroupIdx];

            // Add the sparkline itself (the Add method of SparklineCollection creates the item)
            srcGroup.Sparklines.Add(srcSheet.Name + "!A1:A5", 0, 5);

            // Optional: customize the source group (demonstrates that settings are copied)
            srcGroup.ShowHighPoint = true;
            srcGroup.ShowLowPoint = true;

            // ---------- Create destination worksheet ----------
            Worksheet destSheet = workbook.Worksheets.Add("Destination");

            // Define destination location (column H, rows 1‑5)
            CellArea destLocation = new CellArea
            {
                StartRow = 0,
                EndRow = 4,
                StartColumn = 7,
                EndColumn = 7
            };

            // Clone the sparkline group:
            // 1. Add a new group with the same type, data range and orientation.
            int destGroupIdx = destSheet.SparklineGroups.Add(
                srcGroup.Type,
                srcGroup.Sparklines[0].DataRange, // same data range string
                false,                             // same orientation as source
                destLocation);                     // new location range

            SparklineGroup destGroup = destSheet.SparklineGroups[destGroupIdx];

            // 2. Copy visual settings from source group to destination group.
            destGroup.ShowHighPoint = srcGroup.ShowHighPoint;
            destGroup.ShowLowPoint = srcGroup.ShowLowPoint;
            destGroup.SeriesColor = srcGroup.SeriesColor;
            destGroup.HighPointColor = srcGroup.HighPointColor;
            destGroup.LowPointColor = srcGroup.LowPointColor;
            destGroup.LineWeight = srcGroup.LineWeight;

            // 3. Replicate each sparkline in the source group.
            //    Adjust the row/column indices so they point to the destination location.
            int rowOffset = destLocation.StartRow - srcLocation.StartRow;
            int colOffset = destLocation.StartColumn - srcLocation.StartColumn;

            foreach (Sparkline srcSparkline in srcGroup.Sparklines)
            {
                int newRow = srcSparkline.Row + rowOffset;
                int newCol = srcSparkline.Column + colOffset;
                destGroup.Sparklines.Add(srcSparkline.DataRange, newRow, newCol);
            }

            // ---------- Save the workbook ----------
            workbook.Save("ClonedSparklineGroup.xlsx");
        }
    }
}
