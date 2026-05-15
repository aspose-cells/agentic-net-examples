using Aspose.Cells;
using Aspose.Cells.Charts;

class ColumnSparklineExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a sparkline group of type Column
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Column);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Add a sparkline to the group:
        // Data range: D5:D15
        // Location: cell C5 (row index 4, column index 2)
        group.Sparklines.Add("D5:D15", 4, 2);

        // Save the workbook
        workbook.Save("ColumnSparkline.xlsx");
    }
}