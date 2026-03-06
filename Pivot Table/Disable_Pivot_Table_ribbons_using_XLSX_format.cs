using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["B1"].PutValue("Sales");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["B2"].PutValue(1000);
        sheet.Cells["A3"].PutValue("Orange");
        sheet.Cells["B3"].PutValue(1500);
        sheet.Cells["A4"].PutValue("Banana");
        sheet.Cells["B4"].PutValue(2000);

        // Add a pivot table to the worksheet
        int pivotIndex = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
        PivotTable pivotTable = sheet.PivotTables[pivotIndex];
        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Product column as row field
        pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Sales column as data field

        // Disable the PivotTable wizard and field list (these are part of the ribbon UI)
        pivotTable.EnableWizard = false;
        pivotTable.EnableFieldList = false;

        // Provide custom Ribbon XML that hides the PivotTable related tabs
        string ribbonXml =
            "<customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\">" +
            "  <ribbon>" +
            "    <tabs>" +
            "      <tab idMso=\"TabPivotTableTools\" visible=\"false\" />" +   // Hide PivotTable Tools tab
            "      <tab idMso=\"TabPivotChartTools\" visible=\"false\" />" +   // Hide PivotChart Tools tab
            "    </tabs>" +
            "  </ribbon>" +
            "</customUI>";

        workbook.RibbonXml = ribbonXml;

        // Save the workbook in XLSX format
        workbook.Save("PivotTable_RibbonDisabled.xlsx");
    }
}