using Aspose.Cells;
using Aspose.Cells.Charts;

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
        sheet.Cells["C1"].PutValue(1);
        sheet.Cells["D1"].PutValue(3);

        // Define the location of the original sparkline (cell E1)
        CellArea originalLocation = new CellArea
        {
            StartRow = 0,
            EndRow = 0,
            StartColumn = 4, // column E (0‑based index)
            EndColumn = 4
        };

        // Add a sparkline group with the original sparkline
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, originalLocation);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Retrieve the original sparkline (created by the Add overload above)
        Sparkline originalSparkline = group.Sparklines[0];

        // Duplicate the sparkline to a new destination cell, e.g., F2 (row 1, column 5)
        string dataRange = originalSparkline.DataRange; // same data range as original
        int destinationRow = 1;    // row index for F2 (0‑based)
        int destinationColumn = 5; // column index for F2 (0‑based)

        // Add the duplicated sparkline at the new location
        int duplicatedIndex = group.Sparklines.Add(dataRange, destinationRow, destinationColumn);
        Sparkline duplicatedSparkline = group.Sparklines[duplicatedIndex];

        // Optional: output positions to verify
        System.Console.WriteLine($"Original sparkline - Row: {originalSparkline.Row}, Column: {originalSparkline.Column}");
        System.Console.WriteLine($"Duplicated sparkline - Row: {duplicatedSparkline.Row}, Column: {duplicatedSparkline.Column}");

        // Save the workbook with the duplicated sparkline
        workbook.Save("DuplicatedSparkline.xlsx");
    }
}