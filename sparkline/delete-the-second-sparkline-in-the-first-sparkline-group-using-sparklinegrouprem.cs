using Aspose.Cells;
using Aspose.Cells.Charts;

class DeleteSecondSparkline
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data (two rows, four columns)
        worksheet.Cells["A1"].PutValue(5);
        worksheet.Cells["B1"].PutValue(2);
        worksheet.Cells["C1"].PutValue(1);
        worksheet.Cells["D1"].PutValue(3);
        worksheet.Cells["A2"].PutValue(6);
        worksheet.Cells["B2"].PutValue(4);
        worksheet.Cells["C2"].PutValue(2);
        worksheet.Cells["D2"].PutValue(5);

        // Define the location range where the sparklines will be placed (E1 and E2)
        CellArea location = new CellArea { StartRow = 0, EndRow = 1, StartColumn = 4, EndColumn = 4 };

        // Add a sparkline group that uses the data range A1:D2
        int groupIndex = worksheet.SparklineGroups.Add(SparklineType.Line, "A1:D2", false, location);
        SparklineGroup sparklineGroup = worksheet.SparklineGroups[groupIndex];

        // Delete the second sparkline (index 1) from the group
        sparklineGroup.Sparklines.RemoveAt(1);

        // Save the workbook
        workbook.Save("DeleteSecondSparkline.xlsx");
    }
}