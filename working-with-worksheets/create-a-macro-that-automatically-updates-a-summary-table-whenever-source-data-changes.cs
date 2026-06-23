using Aspose.Cells;
using Aspose.Cells.Pivot;

class Program
{
    static void Main()
    {
        // Create a new workbook and add a worksheet for source data
        Workbook workbook = new Workbook();
        Worksheet dataSheet = workbook.Worksheets[0];
        dataSheet.Name = "Data";

        // Populate source data (Category | Amount)
        dataSheet.Cells["A1"].PutValue("Category");
        dataSheet.Cells["B1"].PutValue("Amount");
        dataSheet.Cells["A2"].PutValue("Food");
        dataSheet.Cells["B2"].PutValue(100);
        dataSheet.Cells["A3"].PutValue("Transport");
        dataSheet.Cells["B3"].PutValue(50);
        dataSheet.Cells["A4"].PutValue("Food");
        dataSheet.Cells["B4"].PutValue(150);

        // Add a worksheet that will hold the summary (pivot table)
        Worksheet summarySheet = workbook.Worksheets.Add("Summary");

        // Define the source range for the pivot table
        string sourceData = "Data!A1:B4";

        // Create the pivot table at cell A1 of the summary sheet
        int pivotIndex = summarySheet.PivotTables.Add(sourceData, "A1", "SummaryPivot");
        PivotTable pivotTable = summarySheet.PivotTables[pivotIndex];

        // Configure the pivot table: Category as row field, Amount as data field
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

        // Save the initial workbook (optional)
        workbook.Save("SummaryMacroDemo.xlsx");

        // ----- Simulate changes in the source data -----
        dataSheet.Cells["B2"].PutValue(120); // Updated Food amount
        dataSheet.Cells["B3"].PutValue(70);  // Updated Transport amount

        // Refresh all pivot tables so the summary reflects the new data
        workbook.Worksheets.RefreshPivotTables();

        // Save the workbook after the automatic update
        workbook.Save("SummaryMacroDemo_Updated.xlsx");
    }
}