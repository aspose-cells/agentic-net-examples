using Aspose.Cells;
using Aspose.Cells.Charts;

class ConfigureSparklineType
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

        // Define the cell area where the sparkline will be placed (E1)
        CellArea location = new CellArea
        {
            StartRow = 0,
            EndRow = 0,
            StartColumn = 4,
            EndColumn = 4
        };

        // Add a sparkline group with an initial type (Column)
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Column, "A1:D1", false, location);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Change the sparkline type to Line as required
        group.Type = SparklineType.Line;

        // Save the workbook to an XLSX file
        workbook.Save("ConfiguredSparklineType.xlsx", SaveFormat.Xlsx);
    }
}