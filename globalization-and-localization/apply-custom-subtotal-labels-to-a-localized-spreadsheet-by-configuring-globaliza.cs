using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Settings;

namespace AsposeCellsCustomSubtotalDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for the pivot table
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Region");
            cells["C1"].PutValue("Sales");
            cells["A2"].PutValue("Bikes");
            cells["B2"].PutValue("North");
            cells["C2"].PutValue(1200);
            cells["A3"].PutValue("Bikes");
            cells["B3"].PutValue("South");
            cells["C3"].PutValue(800);
            cells["A4"].PutValue("Cars");
            cells["B4"].PutValue("North");
            cells["C4"].PutValue(2500);
            cells["A5"].PutValue("Cars");
            cells["B5"].PutValue("South");
            cells["C5"].PutValue(3000);

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:C5", "E1", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure pivot fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);      // Category
            pivotTable.AddFieldToArea(PivotFieldType.Column, 1);   // Region
            int dataFieldPos = pivotTable.AddFieldToArea(PivotFieldType.Data, 2); // Sales
            PivotField dataField = pivotTable.DataFields[dataFieldPos];
            dataField.Function = ConsolidationFunction.Sum; // Use Sum as the aggregation

            // Create a SettablePivotGlobalizationSettings instance to customize subtotal texts
            SettablePivotGlobalizationSettings pivotGlobalSettings = new SettablePivotGlobalizationSettings();

            // Set custom labels for various subtotal types
            pivotGlobalSettings.SetTextOfSubTotal(PivotFieldSubtotalType.Sum, "Custom Sum Subtotal");
            pivotGlobalSettings.SetTextOfSubTotal(PivotFieldSubtotalType.Average, "Custom Avg Subtotal");
            pivotGlobalSettings.SetTextOfSubTotal(PivotFieldSubtotalType.Count, "Custom Count Subtotal");
            pivotGlobalSettings.SetTextOfSubTotal(PivotFieldSubtotalType.Max, "Custom Max Subtotal");
            pivotGlobalSettings.SetTextOfSubTotal(PivotFieldSubtotalType.Min, "Custom Min Subtotal");

            // Assign the customized pivot globalization settings to the workbook
            workbook.Settings.GlobalizationSettings.PivotSettings = pivotGlobalSettings;

            // Refresh and calculate the pivot table so that the new labels take effect
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook to an XLSX file
            workbook.Save("CustomSubtotalLabels.xlsx");
        }
    }
}