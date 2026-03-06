using System;
using System.Globalization;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Settings;

class CustomizePivotCulture
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Populate sample data for the pivot table
        Worksheet dataSheet = workbook.Worksheets[0];
        dataSheet.Cells["A1"].PutValue("Product");
        dataSheet.Cells["B1"].PutValue("Sales");
        dataSheet.Cells["A2"].PutValue("Apple");
        dataSheet.Cells["B2"].PutValue(1234.56);
        dataSheet.Cells["A3"].PutValue("Orange");
        dataSheet.Cells["B3"].PutValue(7890.12);

        // Set workbook culture to French (France) – affects number/date formatting
        workbook.Settings.CultureInfo = new CultureInfo("fr-FR");

        // Add a worksheet for the pivot table
        Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

        // Create the pivot table (source range A1:B3, destination D4)
        int pivotIndex = pivotSheet.PivotTables.Add("A1:B3", "D4", "SalesPivot");
        PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

        // Configure pivot fields
        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Product as row field
        pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Sales as data field

        // Create custom pivot globalization settings
        SettablePivotGlobalizationSettings pivotSettings = new SettablePivotGlobalizationSettings();
        pivotSettings.SetTextOfTotal("Total (FR)");
        pivotSettings.SetTextOfGrandTotal("Grand Total (FR)");
        pivotSettings.SetTextOfRowLabels("Lignes");
        pivotSettings.SetTextOfColumnLabels("Colonnes");

        // Apply the custom settings to the workbook
        SettableGlobalizationSettings globalSettings = new SettableGlobalizationSettings();
        globalSettings.PivotSettings = pivotSettings;
        workbook.Settings.GlobalizationSettings = globalSettings;

        // Refresh and calculate the pivot table to apply the localization
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Save the workbook in XLSX format
        workbook.Save("CustomPivotCulture.xlsx");
    }
}