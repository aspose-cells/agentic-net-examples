using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace PivotTableExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet and add sample data
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "SourceData";

            // Populate sample data
            dataSheet.Cells["A1"].PutValue("Category");
            dataSheet.Cells["B1"].PutValue("Item");
            dataSheet.Cells["C1"].PutValue("Amount");

            dataSheet.Cells["A2"].PutValue("Fruit");
            dataSheet.Cells["B2"].PutValue("Apple");
            dataSheet.Cells["C2"].PutValue(120);

            dataSheet.Cells["A3"].PutValue("Fruit");
            dataSheet.Cells["B3"].PutValue("Banana");
            dataSheet.Cells["C3"].PutValue(80);

            dataSheet.Cells["A4"].PutValue("Vegetable");
            dataSheet.Cells["B4"].PutValue("Carrot");
            dataSheet.Cells["C4"].PutValue(150);

            dataSheet.Cells["A5"].PutValue("Vegetable");
            dataSheet.Cells["B5"].PutValue("Broccoli");
            dataSheet.Cells["C5"].PutValue(90);

            // Create a new worksheet for the pivot table
            Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

            // Define the source data range in A1:C5 of the source sheet
            string sourceData = $"=SourceData!{dataSheet.Cells.MaxDisplayRange.Address}";

            // Add a pivot table at cell A1 of the pivot sheet
            int pivotIndex = pivotSheet.PivotTables.Add(sourceData, "A1", "SalesPivot");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // Configure the pivot table fields
            // Row field: Category
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            // Column field: Item
            pivotTable.AddFieldToArea(PivotFieldType.Column, "Item");
            // Data field: Amount (sum)
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Optional: set a built‑in style for better appearance
            pivotTable.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;

            // Refresh and calculate the pivot data
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook in XLSX format
            workbook.Save("PivotTableExample.xlsx");
        }
    }
}