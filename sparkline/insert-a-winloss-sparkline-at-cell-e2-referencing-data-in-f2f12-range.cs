using Aspose.Cells;
using Aspose.Cells.Charts;

class WinLossSparklineExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data in the range F2:F12 (column index 5, rows 1‑11)
        for (int i = 0; i < 11; i++)
        {
            // Alternate positive and negative values to illustrate win/loss
            sheet.Cells[1 + i, 5].PutValue(i % 2 == 0 ? 1 : -1);
        }

        // Define the location cell for the sparkline (E2 -> column index 4, row index 1)
        CellArea location = new CellArea
        {
            StartRow = 1,
            EndRow = 1,
            StartColumn = 4,
            EndColumn = 4
        };

        // Add a Win/Loss sparkline group that references F2:F12
        // isVertical = true because the data range is a single column
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Stacked, "F2:F12", true, location);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Save the workbook
        workbook.Save("WinLossSparkline.xlsx", SaveFormat.Xlsx);
    }
}