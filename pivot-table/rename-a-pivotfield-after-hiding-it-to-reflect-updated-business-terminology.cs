using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotRenameDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data
            cells["A1"].Value = "Product";
            cells["A2"].Value = "Laptop";
            cells["A3"].Value = "Tablet";
            cells["A4"].Value = "Smartphone";
            cells["A5"].Value = "Laptop";

            cells["B1"].Value = "Sales";
            cells["B2"].Value = 1500;
            cells["B3"].Value = 800;
            cells["B4"].Value = 1200;
            cells["B5"].Value = 1700;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add the "Product" field to the row area
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");

            // Add the "Sales" field to the data area
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Hide all products except "Laptop"
            PivotField productField = pivotTable.RowFields[0];
            for (int i = 0; i < productField.ItemCount; i++)
            {
                // Hide the item if its name is not "Laptop"
                productField.HideItem(i, productField.Items[i] != "Laptop");
            }

            // After hiding, rename the pivot field to match new business terminology
            productField.Name = "DeviceCategory";

            // Refresh and calculate the pivot table to apply changes
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("PivotFieldRenamedAfterHide.xlsx");
        }
    }
}