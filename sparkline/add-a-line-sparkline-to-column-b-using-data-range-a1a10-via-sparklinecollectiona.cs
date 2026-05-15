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
            worksheet.Cells[i, 0].PutValue(i + 1); // Column A (index 0)
        }

        // Add a sparkline group of type Line
        int groupIndex = worksheet.SparklineGroups.Add(SparklineType.Line);
        SparklineGroup sparklineGroup = worksheet.SparklineGroups[groupIndex];

        // Add a line sparkline using the data range A1:A10,
        // placing the sparkline in column B (index 1) at the first row (index 0)
        sparklineGroup.Sparklines.Add("A1:A10", 0, 1);

        // Save the workbook to a file
        workbook.Save("SparklineLine.xlsx");
    }
}