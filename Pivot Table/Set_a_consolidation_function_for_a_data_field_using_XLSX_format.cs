using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

public class Program
{
    public static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        // Protect the worksheet with empty passwords to satisfy API requirements
        sheet.Protect(ProtectionType.All, "", "");

        Cells cells = sheet.Cells;

        // Populate sample data for the pivot table
        cells["A1"].PutValue("Category");
        cells["B1"].PutValue("Amount");
        cells["A2"].PutValue("Food");
        cells["B2"].PutValue(100);
        cells["A3"].PutValue("Food");
        cells["B3"].PutValue(150);
        cells["A4"].PutValue("Drink");
        cells["B4"].PutValue(80);
        cells["A5"].PutValue("Drink");
        cells["B5"].PutValue(120);

        // Add a pivot table that uses the data range A1:B5
        int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "PivotTable1");
        PivotTable pivotTable = sheet.PivotTables[pivotIndex];

        // Add the "Category" column as a row field
        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);

        // Add the "Amount" column as a data field
        int dataFieldIdx = pivotTable.AddFieldToArea(PivotFieldType.Data, 1);
        PivotField dataField = pivotTable.DataFields[dataFieldIdx];

        // Set the consolidation function for the data field (e.g., Average)
        dataField.Function = ConsolidationFunction.Average;

        // Refresh and calculate the pivot table to apply the function
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Save the workbook in XLSX format
        workbook.Save("ConsolidationFunctionPivot.xlsx");
    }
}