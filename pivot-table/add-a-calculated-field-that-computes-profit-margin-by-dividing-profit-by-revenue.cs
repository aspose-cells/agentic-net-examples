using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data: Product, Revenue, Profit
        cells["A1"].PutValue("Product");
        cells["B1"].PutValue("Revenue");
        cells["C1"].PutValue("Profit");

        cells["A2"].PutValue("A"); cells["B2"].PutValue(1000); cells["C2"].PutValue(200);
        cells["A3"].PutValue("B"); cells["B3"].PutValue(1500); cells["C3"].PutValue(300);
        cells["A4"].PutValue("C"); cells["B4"].PutValue(2000); cells["C4"].PutValue(400);

        // Add a pivot table based on the data range
        int pivotIndex = sheet.PivotTables.Add("A1:C4", "E3", "PivotTable1");
        PivotTable pivot = sheet.PivotTables[pivotIndex];

        // Add fields to the pivot table
        pivot.AddFieldToArea(PivotFieldType.Row, "Product");   // Row field
        pivot.AddFieldToArea(PivotFieldType.Data, "Revenue"); // Data field
        pivot.AddFieldToArea(PivotFieldType.Data, "Profit");  // Data field

        // Add a calculated field that computes profit margin = Profit / Revenue
        // This method automatically drags the calculated field to the data area
        pivot.AddCalculatedField("ProfitMargin", "=Profit/Revenue");

        // Retrieve the newly added calculated field and format it as a percentage
        PivotField marginField = pivot.DataFields[pivot.DataFields.Count - 1];
        marginField.NumberFormat = "0.00%";

        // Refresh the pivot table data and calculate the results
        pivot.RefreshData();
        pivot.CalculateData();

        // Save the workbook to a file
        workbook.Save("PivotTable_With_ProfitMargin.xlsx");
    }
}