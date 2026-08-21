// Title: Handle Null or Empty Cells in Aspose.Cells Sparkline (C# .NET)
// Description: Shows how to create a workbook, fill a range with numbers and nulls, add a line sparkline, and use PlotEmptyCellsType (Interpolated, NotPlotted, Zero) to control rendering of empty cells and prevent errors.
// Keywords: Aspose.Cells sparkline null handling | C# sparkline empty cells | PlotEmptyCellsType Interpolated | PlotEmptyCellsType NotPlotted | PlotEmptyCellsType Zero | Aspose.Cells line sparkline example | Excel sparkline missing data | Aspose.Cells .NET sparkline | handle empty cells Aspose | sparkline rendering error
// Common Searches: Aspose.Cells sparkline empty cells | C# set PlotEmptyCellsType for sparkline | avoid sparkline errors with null values Aspose.Cells | interpolated sparkline missing data Aspose | NotPlotted sparkline option .NET | Zero sparkline empty cell Aspose.Cells
// Developer Intent: Configure a sparkline to render correctly when its source range contains null or empty cells.
// Use Cases: Display a smooth trend line by interpolating missing points. | Show gaps in the sparkline where source cells are empty. | Render missing values as zero for consistent visual scaling.
// AI Prompts: Generate C# code with Aspose.Cells that adds a line sparkline and sets PlotEmptyCellsType to Interpolated. | Explain the visual differences between PlotEmptyCellsType.Interpolated, NotPlotted, and Zero in a sparkline. | Describe how empty cells affect sparkline rendering in Aspose.Cells and how to prevent errors.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace SparklineNullHandlingDemo
{
    // Shows how to create a workbook, fill a range with numbers and nulls, add a line sparkline, and use PlotEmptyCellsType (Interpolated, NotPlotted, Zero) to control rendering of empty cells and prevent errors.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate the source data range with some values and nulls
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["B1"].PutValue(null); // Empty cell
            sheet.Cells["C1"].PutValue(12);
            sheet.Cells["D1"].PutValue(null); // Empty cell
            sheet.Cells["E1"].PutValue(8);

            // Define where the sparkline will be placed (single cell)
            CellArea sparklineLocation = new CellArea
            {
                StartColumn = 6, // Column G (0‑based index)
                EndColumn = 6,
                StartRow = 0,    // Row 1
                EndRow = 0
            };

            // Add a sparkline group for the data range A1:E1
            int groupIndex = sheet.SparklineGroups.Add(
                SparklineType.Line,          // Sparkline type
                "A1:E1",                     // Data range (includes nulls)
                false,                       // Plot by row (horizontal)
                sparklineLocation);          // Location of the sparkline

            SparklineGroup group = sheet.SparklineGroups[groupIndex];

            // Configure the group to handle empty cells gracefully.
            // Options: NotPlotted (gap), Zero (display as 0), Interpolated (smoothly connect points)
            group.PlotEmptyCellsType = PlotEmptyCellsType.Interpolated;

            // Optionally, show hidden data and other visual settings
            group.DisplayHidden = true;
            group.ShowHighPoint = true;
            group.ShowLowPoint = true;

            // Save the workbook
            workbook.Save("SparklineWithNullHandling.xlsx");
        }
    }
}
