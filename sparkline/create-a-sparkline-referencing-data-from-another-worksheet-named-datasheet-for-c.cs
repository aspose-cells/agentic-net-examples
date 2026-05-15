using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add a worksheet that will hold the source data
        Worksheet dataSheet = workbook.Worksheets[workbook.Worksheets.Add()];
        dataSheet.Name = "DataSheet";

        // Populate some sample data in DataSheet (row 1, columns A to E)
        for (int col = 0; col < 5; col++)
        {
            dataSheet.Cells[0, col].PutValue(col + 1); // Values 1,2,3,4,5
        }

        // Get the worksheet where the sparkline will be placed
        Worksheet summarySheet = workbook.Worksheets[0];
        summarySheet.Name = "Summary";

        // Define the cell (F1) where the sparkline will appear
        CellArea sparklineLocation = CellArea.CreateCellArea("F1", "F1");

        // Add a sparkline group to the summary sheet.
        // The data range references the other worksheet (DataSheet!A1:E1).
        int groupIndex = summarySheet.SparklineGroups.Add(
            SparklineType.Line,          // Sparkline type
            "DataSheet!A1:E1",           // Cross‑sheet data range
            false,                       // Plot by row (horizontal)
            sparklineLocation);          // Where the sparkline will be placed

        SparklineGroup sparklineGroup = summarySheet.SparklineGroups[groupIndex];

        // Optional: show high/low points for better visual cues
        sparklineGroup.ShowHighPoint = true;
        sparklineGroup.ShowLowPoint = true;

        // Save the workbook
        workbook.Save("CrossSheetSparkline.xlsx");
    }
}