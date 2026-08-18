// Title: Copy Sparkline Group Between Worksheets in C# with Aspose.Cells – Preserve Data Range & Formatting
// Description: Demonstrates how to duplicate a line sparkline group from Sheet1 to Sheet2 using Aspose.Cells for .NET. The example creates a workbook, adds a sparkline group, customizes high‑point/low‑point colors and line weight, extracts the original data range and cell location, then recreates an identical group on another sheet while copying all visual settings.
// Keywords: Aspose.Cells sparkline copy C# | duplicate sparkline group .NET | preserve sparkline formatting | sparkline data range transfer | clone SparklineGroup Aspose | C# workbook sparkline example
// Common Searches: copy sparkline group Aspose.Cells C# | how to duplicate sparkline worksheet .NET | preserve sparkline style when moving sheets | Aspose.Cells copy sparkline data range | clone sparkline group programmatically
// Developer Intent: Programmatically copy an entire sparkline group from one worksheet to another while retaining its source data range and all formatting attributes.
// Use Cases: Create a master template with predefined sparkline visuals and replicate it across multiple report sheets. | Build dashboards where the same sparkline chart appears on each tab with identical styling. | Automate worksheet cloning in large workbooks without losing sparkline appearance.
// AI Prompts: Generate C# code that copies a SparklineGroup from Sheet1 to Sheet2 using Aspose.Cells, including high‑point, low‑point, colors, and line weight. | Explain how to retrieve a sparkline's DataRange and location, then apply them to a new SparklineGroup on a different worksheet. | Show a concise method to clone an Aspose.Cells SparklineGroup without manually setting each property.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace SparklineGroupCopyDemo
{
    // Demonstrates how to duplicate a line sparkline group from Sheet1 to Sheet2 using Aspose.Cells for .NET. The example creates a workbook, adds a sparkline group, customizes high‑point/low‑point colors and line weight, extracts the original data range and cell location, then recreates an identical group on another sheet while copying all visual settings.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and obtain the first worksheet (Sheet1)
            Workbook workbook = new Workbook();
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Sheet1";

            // Populate sample data in Sheet1
            sheet1.Cells["A1"].PutValue(5);
            sheet1.Cells["B1"].PutValue(2);
            sheet1.Cells["C1"].PutValue(1);
            sheet1.Cells["D1"].PutValue(3);

            // Define the location where the sparkline will be placed (cell E1)
            CellArea location = new CellArea
            {
                StartRow = 0,
                EndRow = 0,
                StartColumn = 4,
                EndColumn = 4
            };

            // Add a sparkline group to Sheet1 (Line type)
            int groupIndex = sheet1.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, location);
            SparklineGroup sourceGroup = sheet1.SparklineGroups[groupIndex];

            // Customize some formatting of the source sparkline group
            sourceGroup.ShowHighPoint = true;
            sourceGroup.ShowLowPoint = true;
            sourceGroup.HighPointColor.Color = System.Drawing.Color.Green;
            sourceGroup.LowPointColor.Color = System.Drawing.Color.Red;
            sourceGroup.LineWeight = 1.0;

            // Add a second worksheet (Sheet2) where the sparkline group will be copied
            Worksheet sheet2 = workbook.Worksheets[workbook.Worksheets.Add()];
            sheet2.Name = "Sheet2";

            // Assume the source group contains at least one sparkline.
            // Retrieve the first sparkline to obtain its data range and location.
            Sparkline firstSparkline = sourceGroup.Sparklines[0];
            string dataRange = firstSparkline.DataRange;               // e.g., "Sheet1!A1:D1"
            int row = firstSparkline.Row;                             // row index of the sparkline cell
            int column = firstSparkline.Column;                       // column index of the sparkline cell

            // Build a CellArea for the destination location using the same row/column.
            CellArea destLocation = new CellArea
            {
                StartRow = row,
                EndRow = row,
                StartColumn = column,
                EndColumn = column
            };

            // Add a new sparkline group to Sheet2 with the same parameters.
            int destGroupIndex = sheet2.SparklineGroups.Add(SparklineType.Line, dataRange, false, destLocation);
            SparklineGroup destGroup = sheet2.SparklineGroups[destGroupIndex];

            // Copy formatting properties from the source group to the destination group.
            destGroup.ShowHighPoint = sourceGroup.ShowHighPoint;
            destGroup.ShowLowPoint = sourceGroup.ShowLowPoint;
            destGroup.HighPointColor.Color = sourceGroup.HighPointColor.Color;
            destGroup.LowPointColor.Color = sourceGroup.LowPointColor.Color;
            destGroup.LineWeight = sourceGroup.LineWeight;

            // Save the workbook with both sheets containing identical sparkline groups.
            workbook.Save("SparklineGroupCopyResult.xlsx");
        }
    }
}
