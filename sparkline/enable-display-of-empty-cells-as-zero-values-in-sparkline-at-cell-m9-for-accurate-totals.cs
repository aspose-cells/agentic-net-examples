// Title: Set empty cells to display as zero in a sparkline located at M9 using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel workbook, locates the sparkline at cell M9, and sets its PlotEmptyCellsType property to Zero. | Show how to traverse SparklineGroup collections in Aspose.Cells to change the empty‑cell plotting behavior for a specific sparkline. | Provide a .NET snippet that updates the empty‑cell handling of a sparkline and saves the modified workbook as a new file.
// Common Searches: Aspose.Cells C# set sparkline empty cells to zero in a specific cell | How to change PlotEmptyCellsType for a sparkline at M9 using .NET | Display zero for blank values in Excel sparkline with Aspose.Cells | Update sparkline group property to plot empty cells as zero in C#
// Tags: Aspose.Cells sparkline PlotEmptyCellsType Zero | C# modify sparkline empty cell handling | set sparkline empty cells to zero Aspose | Excel sparkline configuration .NET | target sparkline cell M9 Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads an existing workbook, finds the sparkline positioned at M9, sets its group's PlotEmptyCellsType to Zero so blank cells are plotted as 0, and saves the updated file.
class EnableZeroForEmptySparkline
{
    static void Main()
    {
        // Load the existing workbook that contains the sparkline at cell M9
        Workbook workbook = new Workbook("InputWorkbook.xlsx");

        // Access the first worksheet (adjust index if needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Row and column indexes are zero‑based (M9 => column 12, row 8)
        int targetRow = 8;
        int targetColumn = 12;

        // Iterate through all sparkline groups in the worksheet
        foreach (SparklineGroup group in sheet.SparklineGroups)
        {
            // Check each sparkline in the group
            foreach (Sparkline sparkline in group.Sparklines)
            {
                if (sparkline.Row == targetRow && sparkline.Column == targetColumn)
                {
                    // Found the sparkline at M9 – set empty cells to be plotted as zero
                    group.PlotEmptyCellsType = PlotEmptyCellsType.Zero;
                    // No need to continue searching
                    break;
                }
            }
        }

        // Save the modified workbook
        workbook.Save("OutputWorkbook.xlsx");
    }
}
