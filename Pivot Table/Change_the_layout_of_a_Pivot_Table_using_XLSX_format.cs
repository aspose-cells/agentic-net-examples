using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

public class Program
{
    public static void Main()
    {
        ChangePivotTableLayout.Run();
    }
}

public class ChangePivotTableLayout
{
    public static void Run()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add a worksheet that will hold the source data
        Worksheet dataSheet = workbook.Worksheets[0];
        dataSheet.Name = "Data";

        // Populate sample data for the pivot table
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

        // Add a separate worksheet for the pivot table
        Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

        // Create the pivot table using the data range from the Data sheet
        int pivotIndex = pivotSheet.PivotTables.Add("=Data!A1:C5", "A3", "SalesPivot");
        PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

        // Add fields to the pivot table
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

        // Change the layout to Tabular form
        pivotTable.DisplayInTabularForm = true;

        // Refresh and calculate the pivot table data
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Save the workbook in XLSX format
        workbook.Save("PivotTable_TabularLayout.xlsx");
    }
}