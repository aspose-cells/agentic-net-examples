// Title: Aspose.Cells .NET – Sparkline that Excludes Hidden Columns
// Description: This C# example builds a workbook, hides column B, adds a line sparkline for range A1:D1, sets the SparklineGroup.DisplayHidden flag to false so only visible cells affect the sparkline, applies a blue series color, and saves the file as SparklineIgnoreHidden.xlsx.
// Keywords: Aspose.Cells | .NET | C# sparkline | DisplayHidden | ignore hidden cells | hide column | Excel sparkline | SparklineGroup | financial reporting | data visualization
// Common Searches: Aspose.Cells hide column sparkline C# | DisplayHidden property usage | sparkline ignore hidden rows Aspose | create sparkline for visible data only | C# example Aspose.Cells sparkline
// Developer Intent: Configure a sparkline to calculate using only visible cells.
// Use Cases: Generate trend lines that reflect data after columns are hidden in a financial dashboard. | Produce Excel reports where confidential rows are concealed but should not skew sparkline metrics. | Apply custom colors to sparklines while ensuring hidden cells are excluded from calculations.
// AI Prompts: How do I set Aspose.Cells SparklineGroup.DisplayHidden to false in C#? | Show a C# code snippet that hides a column and creates a sparkline that ignores the hidden column using Aspose.Cells. | Explain the impact of the DisplayHidden flag on sparkline values and how to toggle it programmatically.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSparklineDemo
{
    // This C# example builds a workbook, hides column B, adds a line sparkline for range A1:D1, sets the SparklineGroup.DisplayHidden flag to false so only visible cells affect the sparkline, applies a blue series color, and saves the file as SparklineIgnoreHidden.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data in row 1
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["B1"].PutValue(2);
            sheet.Cells["C1"].PutValue(1);
            sheet.Cells["D1"].PutValue(3);

            // Hide column B (index 1) to simulate hidden data
            sheet.Cells.Columns[1].IsHidden = true;

            // Define the location where the sparkline will be placed (cell E1)
            CellArea sparklineArea = new CellArea
            {
                StartColumn = 4, // Column E
                EndColumn = 4,
                StartRow = 0,    // Row 1
                EndRow = 0
            };

            // Add a sparkline group that uses the data range A1:D1
            int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, sparklineArea);
            SparklineGroup group = sheet.SparklineGroups[groupIndex];

            // Add the sparkline to the group (the Add method also creates the sparkline)
            group.Sparklines.Add(sheet.Name + "!A1:D1", 0, 4);

            // Configure the sparkline to ignore hidden cells (do NOT display hidden data)
            group.DisplayHidden = false; // Only visible cells are considered in calculations

            // Optional: set a series color for visual clarity
            CellsColor seriesColor = workbook.CreateCellsColor();
            seriesColor.Color = Color.Blue;
            group.SeriesColor = seriesColor;

            // Save the workbook
            workbook.Save("SparklineIgnoreHidden.xlsx");
        }
    }
}
