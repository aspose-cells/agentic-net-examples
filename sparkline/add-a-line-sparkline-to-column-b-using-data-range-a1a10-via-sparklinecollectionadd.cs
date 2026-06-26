using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data in the range A1:A10
        for (int i = 0; i < 10; i++)
        {
            sheet.Cells[i, 0].PutValue(i + 1); // Column A (index 0)
        }

        // Add a line sparkline group to the worksheet
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line);
        SparklineGroup sparklineGroup = sheet.SparklineGroups[groupIndex];

        // Add a sparkline that uses the data range A1:A10
        // Place the sparkline in column B (index 1) at row 0
        sparklineGroup.Sparklines.Add("A1:A10", 0, 1);

        // Save the workbook
        workbook.Save("SparklineLine.xlsx");
    }
}