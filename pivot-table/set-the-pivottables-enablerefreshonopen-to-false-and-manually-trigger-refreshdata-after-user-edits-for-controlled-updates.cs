// Title: Aspose.Cells .NET: Disable PivotTable auto‑refresh on open and refresh manually after edits
// Description: Learn how to prevent a PivotTable from refreshing automatically when a workbook is opened using Aspose.Cells for .NET. The example sets RefreshDataOnOpeningFile to false, enables ManualUpdate, modifies source data, and then calls RefreshData and CalculateData to update the pivot on demand before saving the file.
// Keywords: Aspose.Cells PivotTable manual refresh | RefreshDataOnOpeningFile false | ManualUpdate property C# | disable pivot auto refresh Aspose | RefreshData after source edit | Aspose.Cells .NET pivot example | controlled pivot updates
// Common Searches: Aspose.Cells disable pivot auto refresh on open | how to manually refresh a PivotTable with Aspose.Cells | RefreshDataOnOpeningFile vs ManualUpdate Aspose.Cells | C# code to control PivotTable refresh Aspose | Aspose.Cells pivot table refresh after data change
// Developer Intent: Turn off automatic PivotTable refresh on workbook open and invoke a manual refresh after the source data is edited.
// Use Cases: Large workbooks where automatic refresh slows loading; refresh only after user edits. | Template files that keep pivot results static until a specific trigger runs. | Interactive Excel reports where a button or script updates the pivot on demand.
// AI Prompts: Show C# code using Aspose.Cells to set RefreshDataOnOpeningFile = false and refresh a PivotTable after modifying source cells. | Explain how ManualUpdate and RefreshData work together to control PivotTable updates in Aspose.Cells .NET. | Provide a step‑by‑step example that disables auto‑refresh, edits data, and calls RefreshData and CalculateData.

using Aspose.Cells;
using Aspose.Cells.Pivot;

// Learn how to prevent a PivotTable from refreshing automatically when a workbook is opened using Aspose.Cells for .NET. The example sets RefreshDataOnOpeningFile to false, enables ManualUpdate, modifies source data, and then calls RefreshData and CalculateData to update the pivot on demand before saving the file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate source data for the pivot table
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["A4"].PutValue("A");
        worksheet.Cells["B4"].PutValue(30);

        // Add a pivot table to the worksheet
        int pivotIndex = worksheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
        PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

        // Configure the pivot table fields
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Value");

        // Disable automatic refresh when the file is opened
        pivotTable.RefreshDataOnOpeningFile = false;

        // Enable manual update so the pivot table refreshes only on request
        pivotTable.ManualUpdate = true;

        // Initial calculation to populate the pivot table
        pivotTable.CalculateData();

        // ----- Simulate user editing the source data -----
        worksheet.Cells["B2"].PutValue(50); // modify a value
        worksheet.Cells["B3"].PutValue(40); // modify another value

        // Manually refresh the pivot table after edits
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Save the workbook
        workbook.Save("ControlledPivotRefresh.xlsx");
    }
}
