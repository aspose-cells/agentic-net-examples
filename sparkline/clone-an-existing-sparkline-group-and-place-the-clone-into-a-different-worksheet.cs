// Title: Clone a SparklineGroup to another worksheet while preserving formatting with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that copies an existing SparklineGroup from a source worksheet to a target worksheet, keeping all visual properties intact using Aspose.Cells. | Show how to duplicate a line sparkline group and place the clone in cell F1 of a different sheet with the Aspose.Cells API. | Provide a step‑by‑step example of cloning a sparkline group, including data range handling and property transfer, in a .NET workbook.
// Common Searches: Aspose.Cells copy sparkline group to another sheet C# example | how to duplicate sparkline group with formatting using Aspose.Cells .NET | clone sparkline group across worksheets preserving colors Aspose.Cells | C# code to replicate sparkline group in a new worksheet Aspose.Cells | move sparkline group between worksheets without losing settings Aspose.Cells
// Tags: Aspose.Cells clone sparkline group | sparkline group copy between worksheets C# | preserve sparkline visual properties Aspose.Cells | duplicate line sparkline in .xlsx using Aspose | transfer sparkline data range Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSparklineCloneDemo
{
    // The example creates a workbook, adds sample data, defines a line SparklineGroup with custom visual settings on a source worksheet, then creates a new worksheet and clones the group by adding a matching group with the same type, data range, and location. All visual properties (high/low points, colors, line weight) are copied to the target group, and the workbook is saved as SparklineGroupCloneDemo.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // -------------------------------------------------
                // 1. Prepare source worksheet with a sparkline group
                // -------------------------------------------------
                Worksheet sourceSheet = workbook.Worksheets[0];
                sourceSheet.Name = "Source";

                // Populate sample data (A1:D1)
                sourceSheet.Cells["A1"].PutValue(5);
                sourceSheet.Cells["B1"].PutValue(2);
                sourceSheet.Cells["C1"].PutValue(1);
                sourceSheet.Cells["D1"].PutValue(3);

                // Define where the sparkline will be placed (cell F1)
                CellArea sourceLocation = new CellArea
                {
                    StartRow = 0,
                    EndRow = 0,
                    StartColumn = 5, // column F (0‑based index)
                    EndColumn = 5
                };

                // Add the original sparkline group
                int sourceGroupIdx = sourceSheet.SparklineGroups.Add(
                    SparklineType.Line,          // type
                    "A1:D1",                     // data range (relative to source sheet)
                    false,                       // isVertical
                    sourceLocation);             // location range

                SparklineGroup sourceGroup = sourceSheet.SparklineGroups[sourceGroupIdx];

                // Add a sparkline to the group (required for the group to have an item)
                sourceGroup.Sparklines.Add(sourceSheet.Name + "!A1:D1", 0, 5);

                // Set a few visual properties (to demonstrate cloning)
                sourceGroup.ShowHighPoint = true;
                sourceGroup.ShowLowPoint = true;
                sourceGroup.SeriesColor = workbook.CreateCellsColor();
                sourceGroup.SeriesColor.Color = System.Drawing.Color.Orange;

                // -------------------------------------------------
                // 2. Create target worksheet where the clone will go
                // -------------------------------------------------
                Worksheet targetSheet = workbook.Worksheets.Add("CloneTarget");

                // Define the location for the cloned sparkline (cell F1 of the target sheet)
                CellArea targetLocation = new CellArea
                {
                    StartRow = 0,
                    EndRow = 0,
                    StartColumn = 5,
                    EndColumn = 5
                };

                // -------------------------------------------------
                // 3. Clone the sparkline group
                // -------------------------------------------------
                // Retrieve the data range from the first sparkline of the source group
                string dataRange = sourceGroup.Sparklines[0].DataRange; // includes sheet name

                // Add a new group to the target sheet with the same type, data range and orientation
                int targetGroupIdx = targetSheet.SparklineGroups.Add(
                    sourceGroup.Type,   // same sparkline type
                    dataRange,          // same data range (already contains sheet name)
                    false,              // same orientation as source
                    targetLocation);    // new location range

                SparklineGroup targetGroup = targetSheet.SparklineGroups[targetGroupIdx];

                // Add a sparkline to the cloned group (mirrors the source)
                targetGroup.Sparklines.Add(dataRange, 0, 5);

                // Copy visual properties from source to target
                targetGroup.ShowHighPoint = sourceGroup.ShowHighPoint;
                targetGroup.ShowLowPoint = sourceGroup.ShowLowPoint;
                targetGroup.SeriesColor = sourceGroup.SeriesColor;
                targetGroup.HighPointColor = sourceGroup.HighPointColor;
                targetGroup.LowPointColor = sourceGroup.LowPointColor;
                targetGroup.LineWeight = sourceGroup.LineWeight;

                // -------------------------------------------------
                // 4. Save the workbook
                // -------------------------------------------------
                workbook.Save("SparklineGroupCloneDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
