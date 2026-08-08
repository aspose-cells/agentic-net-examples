// Title: C# – Add a line SparklineGroup with simultaneous high‑ and low‑point markers using Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, write values to A1‑D1, place a line sparkline in E1, enable both high‑point and low‑point markers, apply custom green and red colors, and save the file as SparklineHighLowMarkers.xlsx with Aspose.Cells for .NET.
// Keywords: Aspose.Cells SparklineGroup C# | line sparkline high low markers | ShowHighPoint Aspose.Cells | ShowLowPoint Aspose.Cells | custom sparkline marker colors | programmatic sparkline Excel .NET | Aspose.Cells Sparkline example
// Common Searches: Aspose.Cells enable high and low markers in sparkline | C# line sparkline group show high point low point | set custom colors for sparkline markers Aspose.Cells | how to add SparklineGroup with markers using .NET | display both high and low points in Excel sparkline programmatically
// Developer Intent: Enable and style both high‑point and low‑point markers on a line SparklineGroup in an Excel workbook via Aspose.Cells for .NET.
// Use Cases: Sales trend sheets where each row’s sparkline highlights the peak and trough days with distinct colors. | Financial dashboards that instantly flag maximum and minimum values using colored sparkline markers. | Automated report generation that emphasizes extreme data points for quick visual analysis.
// AI Prompts: Generate C# code with Aspose.Cells to create a line SparklineGroup, turn on ShowHighPoint and ShowLowPoint, and assign green and red colors to the markers. | Provide a step‑by‑step tutorial for configuring SparklineGroup markers (high and low) and custom colors in Aspose.Cells for .NET. | Explain how to modify an existing SparklineGroup to display both high and low point markers with different colors using Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace SparklineMarkersDemo
{
    // Demonstrates how to create a workbook, write values to A1‑D1, place a line sparkline in E1, enable both high‑point and low‑point markers, apply custom green and red colors, and save the file as SparklineHighLowMarkers.xlsx with Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the sparkline
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["B1"].PutValue(2);
            sheet.Cells["C1"].PutValue(8);
            sheet.Cells["D1"].PutValue(3);

            // Define the location where the sparkline will be placed (E1)
            CellArea location = new CellArea
            {
                StartColumn = 4, // Column E (0‑based index)
                EndColumn = 4,
                StartRow = 0,    // Row 1
                EndRow = 0
            };

            // Add a line sparkline group with the data range A1:D1
            int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, location);
            SparklineGroup group = sheet.SparklineGroups[groupIndex];

            // Add the sparkline to the group (optional when using Add with dataRange)
            group.Sparklines.Add(sheet.Name + "!A1:D1", 0, 4);

            // Enable markers for both high and low points
            group.ShowHighPoint = true;
            group.ShowLowPoint = true;

            // (Optional) Set colors for the high and low point markers
            CellsColor highColor = workbook.CreateCellsColor();
            highColor.Color = Color.Green;
            group.HighPointColor = highColor;

            CellsColor lowColor = workbook.CreateCellsColor();
            lowColor.Color = Color.Red;
            group.LowPointColor = lowColor;

            // Save the workbook
            workbook.Save("SparklineHighLowMarkers.xlsx");
        }
    }
}
