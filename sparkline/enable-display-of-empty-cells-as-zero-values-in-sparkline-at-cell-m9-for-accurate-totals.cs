using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate row 9 (index 8) with sample data, leaving some cells empty
        sheet.Cells["A9"].PutValue(10);
        sheet.Cells["B9"].PutValue(null); // empty cell
        sheet.Cells["C9"].PutValue(5);
        sheet.Cells["D9"].PutValue(null); // empty cell
        sheet.Cells["E9"].PutValue(7);
        sheet.Cells["F9"].PutValue(3);
        sheet.Cells["G9"].PutValue(null); // empty cell
        sheet.Cells["H9"].PutValue(2);
        sheet.Cells["I9"].PutValue(4);
        sheet.Cells["J9"].PutValue(null); // empty cell
        sheet.Cells["K9"].PutValue(6);
        sheet.Cells["L9"].PutValue(8);

        // Define the location where the sparkline will be placed (cell M9)
        CellArea location = new CellArea
        {
            StartRow = 8,   // Row 9 (zero‑based index)
            EndRow = 8,
            StartColumn = 12, // Column M (zero‑based index)
            EndColumn = 12
        };

        // Add a line sparkline group that uses the data range A9:L9
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A9:L9", false, location);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Configure the group to treat empty cells as zeros for accurate totals
        group.PlotEmptyCellsType = PlotEmptyCellsType.Zero;

        // Save the workbook
        workbook.Save("SparklineEmptyAsZero.xlsx");
    }
}