// Title: Aspose.Cells C# – Show blank cells as zero in a sparkline placed in M9
// Description: Creates a workbook, fills A9:L9 with numeric values leaving gaps, adds a line sparkline to cell M9, sets PlotEmptyCellsType to Zero so empty source cells are plotted as 0, and saves the file as SparklineWithZeroEmptyCells.xlsx.
// Keywords: Aspose.Cells | C# | sparkline | PlotEmptyCellsType | Zero | blank cells as zero | M9 sparkline | line sparkline | Excel sparkline example | global
// Common Searches: Aspose.Cells sparkline treat empty cells as zero | Set PlotEmptyCellsType to Zero C# | Place sparkline in cell M9 Aspose.Cells | How to display blanks as zero in Excel sparkline using .NET | Sparkline empty cells zero Aspose example
// Developer Intent: Configure a sparkline so that any empty source cells are rendered as zero and the sparkline is positioned in cell M9 using Aspose.Cells for .NET.
// Use Cases: Financial dashboards where daily totals may be missing and should appear as zero to keep trend lines consistent. | Automated reporting templates that place row‑wise sparklines in column M and need gaps in data to show flat zero lines. | Sensor‑data worksheets where intermittent missing readings must be visualized as zero to avoid misleading spikes.
// AI Prompts: Generate C# code with Aspose.Cells that adds a line sparkline to cell M9 and sets PlotEmptyCellsType to Zero. | Explain the impact of PlotEmptyCellsType.Zero on sparkline rendering and show how to apply it to a sparkline group. | Provide step‑by‑step instructions to create a sparkline for range A9:L9, locate it in M9, and configure empty cells to be plotted as zeros.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, fills A9:L9 with numeric values leaving gaps, adds a line sparkline to cell M9, sets PlotEmptyCellsType to Zero so empty source cells are plotted as 0, and saves the file as SparklineWithZeroEmptyCells.xlsx.
class SparklineEmptyCellsAsZero
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Example data for the sparkline (some cells are left empty)
        // Fill cells A9 to L9 with values, leaving B9 and E9 empty for demonstration
        sheet.Cells["A9"].PutValue(10);
        // B9 left empty
        sheet.Cells["C9"].PutValue(15);
        sheet.Cells["D9"].PutValue(20);
        // E9 left empty
        sheet.Cells["F9"].PutValue(25);
        sheet.Cells["G9"].PutValue(30);
        sheet.Cells["H9"].PutValue(35);
        sheet.Cells["I9"].PutValue(40);
        sheet.Cells["J9"].PutValue(45);
        sheet.Cells["K9"].PutValue(50);
        sheet.Cells["L9"].PutValue(55);

        // Define the location where the sparkline will be placed (cell M9)
        // Column index for M is 12 (0‑based), row index for 9 is 8
        CellArea location = new CellArea
        {
            StartColumn = 12,
            EndColumn = 12,
            StartRow = 8,
            EndRow = 8
        };

        // Add a sparkline group:
        // - Type: Line
        // - Data range: A9:L9
        // - Horizontal orientation (isVertical = false)
        // - Location defined above
        int groupIdx = sheet.SparklineGroups.Add(
            SparklineType.Line,
            "A9:L9",
            false,
            location);

        SparklineGroup group = sheet.SparklineGroups[groupIdx];

        // Configure the group to treat empty cells as zeros
        group.PlotEmptyCellsType = PlotEmptyCellsType.Zero;

        // (Optional) Add the sparkline explicitly if not already added by Add method
        // The Add method already creates a sparkline, but this shows the syntax:
        // group.Sparklines.Add(sheet.Name + "!A9:L9", 8, 12);

        // Save the workbook
        workbook.Save("SparklineWithZeroEmptyCells.xlsx");
    }
}
