using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class HideSpecificPivotItemsDemo
    {
        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["A2"].PutValue("PO-23-05");
            sheet.Cells["A3"].PutValue("PO-23-06");
            sheet.Cells["A4"].PutValue("PO-23-05");
            sheet.Cells["A5"].PutValue("PO-23-07");

            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["B2"].PutValue(1000);
            sheet.Cells["B3"].PutValue(2000);
            sheet.Cells["B4"].PutValue(1500);
            sheet.Cells["B5"].PutValue(3000);

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "C3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add the "Product" field as a row field and "Sales" as a data field
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Get the row field that contains the product items
            PivotField productField = pivotTable.RowFields[0];

            // Hide all items that are NOT "PO-23-05"
            for (int i = 0; i < productField.ItemCount; i++)
            {
                string itemValue = productField.Items[i];
                bool hide = itemValue != "PO-23-05";
                productField.HideItem(itemValue, hide);
            }

            // Refresh the pivot table after modifying visibility
            pivotTable.CalculateData();

            // Save the workbook in XLSX format
            workbook.Save("HiddenPivotItemsDemo.xlsx");
        }

        public static void Main(string[] args)
        {
            Run();
        }
    }
}