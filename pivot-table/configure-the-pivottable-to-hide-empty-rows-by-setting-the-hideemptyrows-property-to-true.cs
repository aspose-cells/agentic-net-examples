using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;

class HideEmptyRowsPivotDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet dataSheet = workbook.Worksheets[0];

        // Populate sample data with some empty rows
        dataSheet.Cells["A1"].PutValue("Category");
        dataSheet.Cells["B1"].PutValue("Product");
        dataSheet.Cells["C1"].PutValue("Sales");

        dataSheet.Cells["A2"].PutValue("Electronics");
        dataSheet.Cells["B2"].PutValue("TV");
        dataSheet.Cells["C2"].PutValue(1000);

        dataSheet.Cells["A3"].PutValue("Electronics");
        // Empty product and sales rows
        dataSheet.Cells["B3"].PutValue("");
        dataSheet.Cells["C3"].PutValue("");

        dataSheet.Cells["A4"].PutValue("Furniture");
        dataSheet.Cells["B4"].PutValue("Chair");
        dataSheet.Cells["C4"].PutValue(500);

        dataSheet.Cells["A5"].PutValue("Furniture");
        // Empty product and sales rows
        dataSheet.Cells["B5"].PutValue("");
        dataSheet.Cells["C5"].PutValue("");

        // Add a pivot table on the same sheet
        PivotTableCollection pivots = dataSheet.PivotTables;
        int pivotIndex = pivots.Add("A1:C5", "E3", "PivotTable1");
        PivotTable pivot = pivots[pivotIndex];

        // Add fields to the pivot table
        pivot.AddFieldToArea(PivotFieldType.Row, "Category");
        pivot.AddFieldToArea(PivotFieldType.Row, "Product");
        pivot.AddFieldToArea(PivotFieldType.Data, "Sales");

        // Hide empty rows by disabling inclusion of empty rows
        pivot.ShowEmptyRow = false;

        // Calculate the pivot data
        pivot.CalculateData();

        // Save the workbook
        workbook.Save("PivotTableHideEmptyRows.xlsx");
    }
}