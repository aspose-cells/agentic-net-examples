using Aspose.Cells;
using Aspose.Cells.Charts;

class Activate3DSparklines
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the sparkline
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["B1"].PutValue(3);
        sheet.Cells["C1"].PutValue(7);
        sheet.Cells["D1"].PutValue(2);
        sheet.Cells["E1"].PutValue(9);

        // Define the cell area where the sparkline will be placed (single cell)
        CellArea location = new CellArea
        {
            StartRow = 0,
            EndRow = 0,
            StartColumn = 5,
            EndColumn = 5
        };

        // Add a sparkline group of type Line with the data range A1:E1
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:E1", false, location);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Add the sparkline to the group
        group.Sparklines.Add(sheet.Name + "!A1:E1", 0, 5);

        // Save the workbook
        workbook.Save("3D_Sparklines.xlsx");
    }
}