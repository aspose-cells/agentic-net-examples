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
            Cells cells = sheet.Cells;

            // Populate sample data for the pivot table
            cells["A1"].Value = "Product";
            cells["A2"].Value = "Apple";
            cells["A3"].Value = "Banana";
            cells["A4"].Value = "Apple";
            cells["A5"].Value = "Orange";

            cells["B1"].Value = "Sales";
            cells["B2"].Value = 1200;
            cells["B3"].Value = 800;
            cells["B4"].Value = 1500;
            cells["B5"].Value = 2000;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add the "Product" field to the row area and "Sales" to the data area
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Retrieve the row field we want to hide
            PivotField productField = pivotTable.RowFields[0];

            // Hide all items of the field, effectively making the field invisible in the report area
            for (int i = 0; i < productField.ItemCount; i++)
            {
                productField.HideItem(i, true);
            }

            // Refresh and calculate the pivot table after modifications
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("PivotFieldHiddenDemo.xlsx");
        }
    }
}