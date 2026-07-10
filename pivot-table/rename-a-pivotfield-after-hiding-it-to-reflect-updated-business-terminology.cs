using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class RenamePivotFieldAfterHide
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        worksheet.Cells["A1"].Value = "Category";
        worksheet.Cells["B1"].Value = "Amount";
        worksheet.Cells["A2"].Value = "Alpha";
        worksheet.Cells["B2"].Value = 100;
        worksheet.Cells["A3"].Value = "Beta";
        worksheet.Cells["B3"].Value = 200;
        worksheet.Cells["A4"].Value = "Gamma";
        worksheet.Cells["B4"].Value = 300;
        worksheet.Cells["A5"].Value = "Alpha";
        worksheet.Cells["B5"].Value = 150;

        // Add a pivot table covering the data range
        int pivotIndex = worksheet.PivotTables.Add("A1:B5", "D3", "PivotTable1");
        PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

        // Add the "Category" field to the row area and "Amount" to the data area
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

        // Hide all row items except the one named "Alpha"
        PivotField rowField = pivotTable.RowFields[0];
        for (int i = 0; i < rowField.ItemCount; i++)
        {
            // HideItem(int index, bool isHidden) – hide if the item is not "Alpha"
            rowField.HideItem(i, rowField.Items[i] != "Alpha");
        }

        // Rename the pivot field to reflect new business terminology
        rowField.Name = "ProductGroup";

        // Refresh and calculate the pivot table to apply changes
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Save the workbook
        workbook.Save("RenamedPivotField.xlsx");
    }
}