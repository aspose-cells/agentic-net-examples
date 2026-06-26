using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data in D5:D15 (optional, for demonstration)
        for (int i = 0; i < 11; i++)
        {
            sheet.Cells[4 + i, 3].PutValue(i + 1); // Row index 4+i, column index 3 (D)
        }

        // Add a sparkline group of type Column
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Column);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Add a sparkline to cell C5 (row 4, column 2) using data from D5:D15
        group.Sparklines.Add("D5:D15", 4, 2);

        // Save the workbook
        workbook.Save("ColumnSparkline.xlsx");
    }
}