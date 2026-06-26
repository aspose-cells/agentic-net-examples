using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some sample data
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["B1"].PutValue(2);
        sheet.Cells["C1"].PutValue(1);
        sheet.Cells["D1"].PutValue(3);

        // Define the location where the sparkline will be placed
        CellArea location = new CellArea
        {
            StartColumn = 4,
            EndColumn = 4,
            StartRow = 0,
            EndRow = 0
        };

        // Add a sparkline group to the worksheet
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, location);
        SparklineGroup sparklineGroup = sheet.SparklineGroups[groupIndex];

        // Add a sparkline to the group (just for demonstration)
        sparklineGroup.Sparklines.Add(sheet.Name + "!A1:D1", 0, 4);

        // Remove all sparklines by clearing each group's SparklineCollection
        foreach (SparklineGroup sg in sheet.SparklineGroups)
        {
            sg.Sparklines.Clear(); // Clears all sparklines within the group
        }

        // Save the workbook
        workbook.Save("RemovedSparklines.xlsx", SaveFormat.Xlsx);
    }
}