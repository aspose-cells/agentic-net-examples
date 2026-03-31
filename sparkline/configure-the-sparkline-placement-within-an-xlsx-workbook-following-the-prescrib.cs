using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data (3 rows × 5 columns) that will be used for the sparklines
        for (int row = 0; row < 3; row++)
        {
            for (int col = 0; col < 5; col++)
            {
                sheet.Cells[row, col].PutValue((row + 1) * (col + 1));
            }
        }

        // Define the location range where the sparklines will be placed:
        // Column F (index 5), rows 1‑3 (indices 0‑2)
        CellArea locationRange = new CellArea
        {
            StartRow = 0,
            EndRow = 2,
            StartColumn = 5,
            EndColumn = 5
        };

        // Add a sparkline group:
        // - Type: Line
        // - Data range: A1:E3 (covers the data populated above)
        // - isVertical: false (plot by rows)
        // - locationRange: defined above
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:E3", false, locationRange);
        SparklineGroup sparklineGroup = sheet.SparklineGroups[groupIndex];

        // Add a sparkline for each row of data.
        // Each sparkline uses the row's data range and is placed in column F of the same row.
        for (int row = 0; row < 3; row++)
        {
            string dataRange = $"A{row + 1}:E{row + 1}";
            sparklineGroup.Sparklines.Add(dataRange, row, 5);
        }

        // OPTIONAL: Change the placement of the entire group to a new location (column G, rows 1‑3)
        // This demonstrates the ResetRanges method which re‑creates the sparklines at the new location.
        CellArea newLocation = CellArea.CreateCellArea(0, 6, 2, 6); // G1:G3
        sparklineGroup.ResetRanges("A1:E3", false, newLocation);

        // Save the workbook as an XLSX file
        workbook.Save("SparklinePlacementDemo.xlsx");
    }
}