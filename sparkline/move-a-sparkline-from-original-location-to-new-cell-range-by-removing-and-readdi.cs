using Aspose.Cells;
using Aspose.Cells.Charts;

class MoveSparklineDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the sparkline
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["B1"].PutValue(2);
        sheet.Cells["C1"].PutValue(1);
        sheet.Cells["D1"].PutValue(3);

        // Define the original location range for the sparkline (cell E1)
        CellArea originalLocation = new CellArea
        {
            StartRow = 0,
            EndRow = 0,
            StartColumn = 4, // Column E (0‑based index)
            EndColumn = 4
        };

        // Add a sparkline group with the data range A1:D1 at the original location
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, originalLocation);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Retrieve the sparkline that was created by the Add method
        Sparkline sparkline = group.Sparklines[0];

        // Define the new location for the sparkline (cell G1)
        int newRow = 0;   // Row 1 (0‑based)
        int newColumn = 6; // Column G (0‑based)

        // Remove the existing sparkline from the group
        group.Sparklines.RemoveSparkline(sparkline);

        // Add a new sparkline at the new location using the same data range
        group.Sparklines.Add(sparkline.DataRange, newRow, newColumn);

        // Save the workbook with the moved sparkline
        workbook.Save("MovedSparkline.xlsx");
    }
}