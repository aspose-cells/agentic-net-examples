using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Settings;

namespace AsposeCellsPivotGlobalizationDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet dataSheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            dataSheet.Cells["A1"].PutValue("Product");
            dataSheet.Cells["B1"].PutValue("Region");
            dataSheet.Cells["C1"].PutValue("Sales");

            dataSheet.Cells["A2"].PutValue("Apple");
            dataSheet.Cells["B2"].PutValue("North");
            dataSheet.Cells["C2"].PutValue(1200);

            dataSheet.Cells["A3"].PutValue("Apple");
            dataSheet.Cells["B3"].PutValue("South");
            dataSheet.Cells["C3"].PutValue(1500);

            dataSheet.Cells["A4"].PutValue("Orange");
            dataSheet.Cells["B4"].PutValue("North");
            dataSheet.Cells["C4"].PutValue(800);

            dataSheet.Cells["A5"].PutValue("Orange");
            dataSheet.Cells["B5"].PutValue("South");
            dataSheet.Cells["C5"].PutValue(950);

            // Add a new worksheet for the pivot table
            Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

            // Create the pivot table (source range A1:C5, destination cell E1)
            int pivotIndex = pivotSheet.PivotTables.Add("A1:C5", "E1", "SalesPivot");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // Configure pivot fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);    // Product as row field
            pivotTable.AddFieldToArea(PivotFieldType.Column, 1); // Region as column field
            pivotTable.AddFieldToArea(PivotFieldType.Data, 2);   // Sales as data field

            // Create SettableGlobalizationSettings and its PivotSettings
            SettableGlobalizationSettings globalSettings = new SettableGlobalizationSettings();
            SettablePivotGlobalizationSettings pivotGlobalSettings = new SettablePivotGlobalizationSettings();

            // Customize various pivot table texts
            pivotGlobalSettings.SetTextOfTotal("Total Sales");
            pivotGlobalSettings.SetTextOfGrandTotal("Grand Total Sales");
            pivotGlobalSettings.SetTextOfAll("All Products");
            pivotGlobalSettings.SetTextOfColumnLabels("Region Columns");
            pivotGlobalSettings.SetTextOfRowLabels("Product Rows");
            pivotGlobalSettings.SetTextOfEmptyData("No Data");
            pivotGlobalSettings.SetTextOfDataFieldHeader("Sales Amount");

            // Assign the customized pivot settings to the global settings
            globalSettings.PivotSettings = pivotGlobalSettings;

            // Apply the globalization settings to the workbook
            workbook.Settings.GlobalizationSettings = globalSettings;

            // Refresh and calculate the pivot table to apply the custom texts
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook in XLSX format
            workbook.Save("CustomPivotGlobalization.xlsx");
        }
    }
}