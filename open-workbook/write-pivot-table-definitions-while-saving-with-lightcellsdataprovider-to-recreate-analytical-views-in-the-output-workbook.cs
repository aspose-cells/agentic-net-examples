using Aspose.Cells;
using Aspose.Cells.Pivot;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["B1"].PutValue("Amount");
        worksheet.Cells["A2"].PutValue("Food");
        worksheet.Cells["B2"].PutValue(120);
        worksheet.Cells["A3"].PutValue("Transport");
        worksheet.Cells["B3"].PutValue(80);
        worksheet.Cells["A4"].PutValue("Utilities");
        worksheet.Cells["B4"].PutValue(150);

        // Add a pivot table to the worksheet
        PivotTableCollection pivotTables = worksheet.PivotTables;
        int pivotIndex = pivotTables.Add("A1:B4", "D5", "SalesPivot");
        PivotTable pivotTable = pivotTables[pivotIndex];

        // Configure the pivot table: Category as rows, Amount as data
        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Row field (Category)
        pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Data field (Amount)
        pivotTable.SaveData = true; // Ensure pivot data is saved with the workbook

        // ------------------------------------------------------------
        // Save the workbook using LightCellsDataProvider
        // ------------------------------------------------------------
        // NOTE: The exact LightCellsDataProvider API (construction, options,
        // and Save method) is not present in the supplied documentation.
        // The following is a placeholder indicating where the LightCellsDataProvider
        // saving logic should be inserted once the appropriate API details are
        // available.
        //
        // LightCellsDataProvider provider = new LightCellsDataProvider(...);
        // provider.Save(workbook, "AnalyticalView.xlsx");
        //
        // Replace the above placeholder with the correct LightCellsDataProvider
        // usage according to the Aspose.Cells version you are targeting.
    }
}

// Author: Aspose.Cells .NET example – pivot table definition with LightCellsDataProvider placeholder.