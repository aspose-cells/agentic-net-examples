using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

public class PivotTableExample
{
    public static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add a worksheet that will contain the source data
        Worksheet dataSheet = workbook.Worksheets[0];
        dataSheet.Name = "Data";

        // Populate sample data
        dataSheet.Cells["A1"].PutValue("Category");
        dataSheet.Cells["B1"].PutValue("Product");
        dataSheet.Cells["C1"].PutValue("Sales");

        dataSheet.Cells["A2"].PutValue("Electronics");
        dataSheet.Cells["B2"].PutValue("Laptop");
        dataSheet.Cells["C2"].PutValue(1200);

        dataSheet.Cells["A3"].PutValue("Electronics");
        dataSheet.Cells["B3"].PutValue("Phone");
        dataSheet.Cells["C3"].PutValue(800);

        dataSheet.Cells["A4"].PutValue("Furniture");
        dataSheet.Cells["B4"].PutValue("Chair");
        dataSheet.Cells["C4"].PutValue(150);

        dataSheet.Cells["A5"].PutValue("Furniture");
        dataSheet.Cells["B5"].PutValue("Table");
        dataSheet.Cells["C5"].PutValue(300);

        // Add a new worksheet that will host the pivot table
        Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

        // Define the source data range (including sheet name)
        string sourceData = "=Data!A1:C5";

        // Add the pivot table at cell A3 of the pivot sheet
        // Using the Add(string sourceData, string destCellName, string tableName) overload
        int pivotIndex = pivotSheet.PivotTables.Add(sourceData, "A3", "SalesPivot");

        // Retrieve the created pivot table
        PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

        // Configure the pivot table fields
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");   // Row field
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");    // Additional row field
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");     // Data field

        // Optional: display the pivot table in tabular form
        pivotTable.ShowInTabularForm();

        // Populate the pivot table with calculated data
        pivotTable.CalculateData();

        // Save the workbook in XLSX format
        workbook.Save("PivotTableExample.xlsx");
    }
}