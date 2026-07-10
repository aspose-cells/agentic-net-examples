using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class CreateMultiplePivotTables
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet for data
        Workbook workbook = new Workbook();
        Worksheet dataSheet = workbook.Worksheets[0];
        dataSheet.Name = "Data";

        // Populate sample data
        Cells cells = dataSheet.Cells;
        cells["A1"].PutValue("Region");
        cells["B1"].PutValue("Product");
        cells["C1"].PutValue("Sales");
        cells["D1"].PutValue("Quantity");

        string[] regions = { "North", "South", "East", "West" };
        string[] products = { "Apple", "Banana", "Cherry" };
        Random rnd = new Random();
        int currentRow = 2;

        for (int i = 0; i < 20; i++)
        {
            cells[currentRow, 0].PutValue(regions[rnd.Next(regions.Length)]);
            cells[currentRow, 1].PutValue(products[rnd.Next(products.Length)]);
            cells[currentRow, 2].PutValue(rnd.Next(1000, 5000));   // Sales
            cells[currentRow, 3].PutValue(rnd.Next(10, 100));     // Quantity
            currentRow++;
        }

        // Define the source data range for the pivot tables
        string sourceData = $"=Data!A1:D{currentRow - 1}";

        // Add a new worksheet to hold the pivot tables
        Worksheet pivotSheet = workbook.Worksheets.Add("PivotTables");
        PivotTableCollection pivotTables = pivotSheet.PivotTables;

        // -------------------------------------------------
        // First PivotTable: Sales summarized by Region and Product
        // -------------------------------------------------
        int pivotIndex1 = pivotTables.Add(sourceData, "A3", "SalesByRegion");
        PivotTable salesPivot = pivotTables[pivotIndex1];
        salesPivot.AddFieldToArea(PivotFieldType.Row, "Region");
        salesPivot.AddFieldToArea(PivotFieldType.Column, "Product");
        salesPivot.AddFieldToArea(PivotFieldType.Data, "Sales");
        salesPivot.ShowInTabularForm();

        // -------------------------------------------------
        // Second PivotTable: Quantity summarized by Product
        // -------------------------------------------------
        // Use overload that specifies row and column indices for placement
        int pivotIndex2 = pivotTables.Add(sourceData, 20, 0, "QuantityByProduct");
        PivotTable qtyPivot = pivotTables[pivotIndex2];
        qtyPivot.AddFieldToArea(PivotFieldType.Row, "Product");
        qtyPivot.AddFieldToArea(PivotFieldType.Data, "Quantity");
        qtyPivot.ShowInOutlineForm();

        // Refresh all pivot tables to ensure they reflect the source data
        workbook.Worksheets.RefreshPivotTables();

        // Save the workbook
        workbook.Save("MultiplePivotTables.xlsx");
    }
}