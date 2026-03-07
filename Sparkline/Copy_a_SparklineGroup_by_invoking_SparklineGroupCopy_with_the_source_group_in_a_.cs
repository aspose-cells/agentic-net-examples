using Aspose.Cells;
using Aspose.Cells.Charts;

class SparklineGroupCopyDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet (source sheet)
        Workbook workbook = new Workbook();
        Worksheet sourceSheet = workbook.Worksheets[0];
        sourceSheet.Name = "Source";

        // Populate sample data for the sparkline
        sourceSheet.Cells["A1"].PutValue(5);
        sourceSheet.Cells["B1"].PutValue(2);
        sourceSheet.Cells["C1"].PutValue(1);
        sourceSheet.Cells["D1"].PutValue(3);

        // Define the location where the sparkline will be placed
        CellArea location = new CellArea
        {
            StartRow = 0,
            EndRow = 0,
            StartColumn = 4,
            EndColumn = 4
        };

        // Add a sparkline group to the source sheet
        int sourceGroupIndex = sourceSheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, location);
        SparklineGroup sourceGroup = sourceSheet.SparklineGroups[sourceGroupIndex];
        sourceGroup.Sparklines.Add(sourceSheet.Name + "!A1:D1", 0, 4);
        sourceGroup.ShowHighPoint = true;
        sourceGroup.ShowLowPoint = true;

        // Add a new worksheet that will receive the copied sparkline group
        int targetSheetIndex = workbook.Worksheets.Add();
        Worksheet targetSheet = workbook.Worksheets[targetSheetIndex];
        targetSheet.Name = "Target";

        // Create a sparkline group in the target sheet with the same type
        int targetGroupIndex = targetSheet.SparklineGroups.Add(sourceGroup.Type, "A1:D1", false, location);
        SparklineGroup targetGroup = targetSheet.SparklineGroups[targetGroupIndex];

        // Copy relevant properties
        targetGroup.ShowHighPoint = sourceGroup.ShowHighPoint;
        targetGroup.ShowLowPoint = sourceGroup.ShowLowPoint;

        // Add a sparkline to the target group referencing the same source data
        targetGroup.Sparklines.Add(sourceSheet.Name + "!A1:D1", 0, 4);

        // Save the workbook in XLSX format
        workbook.Save("SparklineGroupCopyDemo.xlsx", SaveFormat.Xlsx);
    }
}