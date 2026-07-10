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
        cells["A1"].Value = "Product";
        cells["B1"].Value = "Revenue";
        cells["C1"].Value = "Profit";

        cells["A2"].Value = "A";
        cells["B2"].Value = 1000;
        cells["C2"].Value = 200;

        cells["A3"].Value = "B";
        cells["B3"].Value = 1500;
        cells["C3"].Value = 300;

        cells["A4"].Value = "C";
        cells["B4"].Value = 2000;
        cells["C4"].Value = 400;

        // Add a pivot table that uses the data range A1:C4 and place it at E3
        int pivotIndex = sheet.PivotTables.Add("A1:C4", "E3", "PivotTable1");
        PivotTable pivotTable = sheet.PivotTables[pivotIndex];

        // Add fields to the pivot table
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");          // Row field
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Revenue");        // Data field
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Profit");         // Data field

        // Add a calculated field named "ProfitMargin" with the formula =Profit/Revenue
        // The third argument 'true' drags the field to the data area automatically
        pivotTable.AddCalculatedField("ProfitMargin", "=Profit/Revenue", true);

        // Format the calculated field as a percentage
        PivotField profitMarginField = pivotTable.DataFields[pivotTable.DataFields.Count - 1];
        profitMarginField.NumberFormat = "0.00%";

        // Refresh the pivot table data and calculate the results
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Save the workbook to a file
        workbook.Save("PivotTable_With_ProfitMargin.xlsx");
    }
}