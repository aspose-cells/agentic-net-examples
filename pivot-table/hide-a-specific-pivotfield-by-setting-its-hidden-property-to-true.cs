using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotHideFieldDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].Value = "Category";
            sheet.Cells["A2"].Value = "Apple";
            sheet.Cells["A3"].Value = "Orange";
            sheet.Cells["A4"].Value = "Apple";
            sheet.Cells["A5"].Value = "Banana";

            sheet.Cells["B1"].Value = "Quantity";
            sheet.Cells["B2"].Value = 10;
            sheet.Cells["B3"].Value = 15;
            sheet.Cells["B4"].Value = 20;
            sheet.Cells["B5"].Value = 5;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add the "Category" field to the row area
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");

            // Add the "Quantity" field to the data area
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Quantity");

            // Get the row field (Category)
            PivotField categoryField = pivotTable.RowFields[0];

            // Hide a specific item within the field, e.g., hide "Apple"
            // This uses the HideItem(string, bool) method to set the item as hidden
            categoryField.HideItem("Apple", true);

            // Refresh and calculate the pivot table to apply changes
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook with the hidden pivot item
            workbook.Save("PivotFieldItemHiddenDemo.xlsx");
        }
    }
}