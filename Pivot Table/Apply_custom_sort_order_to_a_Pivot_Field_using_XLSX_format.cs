using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class CustomPivotSort
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data for the pivot table
        cells["A1"].PutValue("Category");
        cells["B1"].PutValue("Item");
        cells["C1"].PutValue("Sales");

        string[,] data = new string[,]
        {
            {"Fruit","Apple","120"},
            {"Fruit","Banana","80"},
            {"Fruit","Orange","150"},
            {"Vegetable","Carrot","90"},
            {"Vegetable","Broccoli","110"},
            {"Vegetable","Spinach","70"}
        };

        for (int i = 0; i < data.GetLength(0); i++)
        {
            cells[i + 1, 0].PutValue(data[i, 0]);
            cells[i + 1, 1].PutValue(data[i, 1]);
            cells[i + 1, 2].PutValue(double.Parse(data[i, 2]));
        }

        // Create a pivot table based on the data range
        int ptIndex = sheet.PivotTables.Add("A1:C7", "E3", "PivotTable1");
        PivotTable pivotTable = sheet.PivotTables[ptIndex];

        // Add fields: Category as row, Item as column, Sales as data
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
        pivotTable.AddFieldToArea(PivotFieldType.Column, "Item");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

        // Enable custom list sorting so that Excel's custom list order is respected
        pivotTable.CustomListSort = true;

        // Apply custom sort to the row field.
        // Sort by the field's own labels (fieldSortedBy = -1) in ascending order.
        // With CustomListSort enabled, this will follow any custom list defined in Excel.
        PivotField rowField = pivotTable.RowFields[0];
        rowField.SortBy(SortOrder.Ascending, -1);

        // Refresh and calculate the pivot table to apply sorting
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Save the workbook in XLSX format
        workbook.Save("CustomPivotSort.xlsx", SaveFormat.Xlsx);
    }
}