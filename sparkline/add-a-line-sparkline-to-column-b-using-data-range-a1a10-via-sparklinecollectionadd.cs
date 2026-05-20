using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data in the range A1:A10
        for (int i = 0; i < 10; i++)
        {
            // Row index i, column index 0 corresponds to column A
            worksheet.Cells[i, 0].PutValue(i + 1);
        }

        // Add a sparkline group of type Line to the worksheet
        int groupIndex = worksheet.SparklineGroups.Add(SparklineType.Line);
        SparklineGroup sparklineGroup = worksheet.SparklineGroups[groupIndex];

        // Add a sparkline that uses the data range A1:A10
        // Place the sparkline in column B (index 1) at the first row (index 0)
        sparklineGroup.Sparklines.Add("A1:A10", 0, 1);

        // Save the workbook to a file
        workbook.Save("SparklineLine.xlsx");
    }
}