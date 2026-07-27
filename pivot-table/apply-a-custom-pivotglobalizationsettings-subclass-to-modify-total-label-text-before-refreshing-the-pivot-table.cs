using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Settings;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet for data
        Workbook workbook = new Workbook();
        Worksheet dataSheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        dataSheet.Cells["A1"].PutValue("Category");
        dataSheet.Cells["B1"].PutValue("Amount");
        dataSheet.Cells["A2"].PutValue("Fruit");
        dataSheet.Cells["B2"].PutValue(120);
        dataSheet.Cells["A3"].PutValue("Vegetable");
        dataSheet.Cells["B3"].PutValue(80);
        dataSheet.Cells["A4"].PutValue("Fruit");
        dataSheet.Cells["B4"].PutValue(150);
        dataSheet.Cells["A5"].PutValue("Vegetable");
        dataSheet.Cells["B5"].PutValue(200);

        // Add a worksheet that will contain the pivot table
        Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");

        // Create the pivot table (source range A1:B5, destination D1)
        int pivotIndex = pivotSheet.PivotTables.Add("A1:B5", "D1", "MyPivot");
        PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

        // Add fields: Category as row, Amount as data
        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);
        pivotTable.AddFieldToArea(PivotFieldType.Data, 1);

        // Create a SettablePivotGlobalizationSettings instance and change the "Total" label
        SettablePivotGlobalizationSettings pivotSettings = new SettablePivotGlobalizationSettings();
        pivotSettings.SetTextOfTotal("Custom Total");

        // Attach the custom settings to the workbook
        GlobalizationSettings globalization = new GlobalizationSettings();
        globalization.PivotSettings = pivotSettings;
        workbook.Settings.GlobalizationSettings = globalization;

        // Refresh and calculate the pivot table after applying the settings
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Save the workbook
        workbook.Save("CustomTotalPivot.xlsx");
    }
}