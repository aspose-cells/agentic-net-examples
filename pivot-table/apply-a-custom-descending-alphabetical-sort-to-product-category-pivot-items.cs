using Aspose.Cells;
using Aspose.Cells.Pivot;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data: Product Category and Sales
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["B1"].PutValue("Sales");
        worksheet.Cells["A2"].PutValue("Electronics");
        worksheet.Cells["A3"].PutValue("Furniture");
        worksheet.Cells["A4"].PutValue("Clothing");
        worksheet.Cells["B2"].PutValue(1200);
        worksheet.Cells["B3"].PutValue(800);
        worksheet.Cells["B4"].PutValue(500);

        // Add a pivot table based on the data range
        int pivotIndex = worksheet.PivotTables.Add("A1:B4", "D3", "PivotTable1");
        PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

        // Add the Category field as a row field
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
        // Add the Sales field as a data field
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

        // Retrieve the Category pivot field
        PivotField categoryField = pivotTable.RowFields[0];

        // Apply descending alphabetical sort (sort by labels, descending)
        categoryField.SortBy(SortOrder.Descending, -1);

        // Refresh and calculate the pivot table to apply sorting
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Save the workbook with the sorted pivot table
        workbook.Save("PivotCategoryDescAlphabetical.xlsx");
    }
}