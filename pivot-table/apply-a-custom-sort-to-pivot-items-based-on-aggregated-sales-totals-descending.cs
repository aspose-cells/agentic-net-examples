using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotCustomSort
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data: Product and Sales
            cells["A1"].PutValue("Product");
            cells["B1"].PutValue("Sales");
            cells["A2"].PutValue("Apple");
            cells["A3"].PutValue("Banana");
            cells["A4"].PutValue("Cherry");
            cells["A5"].PutValue("Date");
            cells["B2"].PutValue(1200);
            cells["B3"].PutValue(800);
            cells["B4"].PutValue(1500);
            cells["B5"].PutValue(600);

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add the row field (Product) and the data field (Sales)
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Apply custom sort: sort the row field by the aggregated Sales totals in descending order
            // -1 indicates sorting by the data labels (i.e., the calculated totals) of this field
            PivotField rowField = pivotTable.RowFields[0];
            rowField.SortBy(SortOrder.Descending, -1);

            // Refresh and calculate the pivot table to apply sorting
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("CustomSortedPivot.xlsx");
        }
    }
}