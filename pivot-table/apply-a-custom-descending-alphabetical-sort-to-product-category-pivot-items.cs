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

            // Populate sample data: Product Category and Sales
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["A2"].PutValue("Electronics");
            sheet.Cells["A3"].PutValue("Furniture");
            sheet.Cells["A4"].PutValue("Clothing");
            sheet.Cells["A5"].PutValue("Books");
            sheet.Cells["B2"].PutValue(1200);
            sheet.Cells["B3"].PutValue(800);
            sheet.Cells["B4"].PutValue(450);
            sheet.Cells["B5"].PutValue(300);

            // Add a pivot table based on the data range
            int pivotIdx = sheet.PivotTables.Add("A1:B5", "D3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIdx];

            // Add the Category field to the Row area
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");

            // Add the Sales field to the Data area
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Apply custom descending alphabetical sort to the Category pivot field
            // -1 indicates sorting by the field's own labels (alphabetical)
            PivotField categoryField = pivotTable.RowFields[0];
            categoryField.SortBy(SortOrder.Descending, -1);

            // Refresh and calculate the pivot table to apply sorting
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("PivotTable_CustomDescendingAlphabeticalSort.xlsx");
        }
    }
}