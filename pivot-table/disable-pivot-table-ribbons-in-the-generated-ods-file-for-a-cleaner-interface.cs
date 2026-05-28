using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class DisablePivotTableRibbons
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Amount");
        sheet.Cells["A2"].PutValue("Food");
        sheet.Cells["B2"].PutValue(120);
        sheet.Cells["A3"].PutValue("Clothing");
        sheet.Cells["B3"].PutValue(80);
        sheet.Cells["A4"].PutValue("Travel");
        sheet.Cells["B4"].PutValue(200);

        // Add a pivot table to the worksheet
        int pivotIndex = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
        PivotTable pivotTable = sheet.PivotTables[pivotIndex];
        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Category as row field
        pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Amount as data field

        // Hide the pivot field list (ribbon UI) for a cleaner ODS interface
        workbook.Settings.HidePivotFieldList = true;

        // Save the workbook as ODS using default OdsSaveOptions
        OdsSaveOptions saveOptions = new OdsSaveOptions();
        workbook.Save("CleanPivot.ods", saveOptions);
    }
}