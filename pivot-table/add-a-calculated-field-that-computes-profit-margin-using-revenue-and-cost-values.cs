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

        // Populate sample data: Product, Revenue, Cost
        cells["A1"].Value = "Product";
        cells["B1"].Value = "Revenue";
        cells["C1"].Value = "Cost";

        cells["A2"].Value = "A";
        cells["B2"].Value = 5000;
        cells["C2"].Value = 3000;

        cells["A3"].Value = "B";
        cells["B3"].Value = 7000;
        cells["C3"].Value = 4200;

        cells["A4"].Value = "C";
        cells["B4"].Value = 6000;
        cells["C4"].Value = 3600;

        // Add a pivot table based on the data range
        int pivotIndex = sheet.PivotTables.Add("A1:C4", "E3", "SalesPivot");
        PivotTable pivotTable = sheet.PivotTables[pivotIndex];

        // Add fields to the pivot table
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Revenue");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Cost");

        // Add a calculated field for profit margin: (Revenue - Cost) / Revenue
        // The formula references the source field names exactly.
        pivotTable.AddCalculatedField("ProfitMargin", "=(Revenue-Cost)/Revenue", true);

        // Format the calculated field as a percentage
        PivotField profitMarginField = pivotTable.DataFields[pivotTable.DataFields.Count - 1];
        profitMarginField.NumberFormat = "0.00%";

        // Refresh and calculate the pivot table
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Save the workbook
        workbook.Save("PivotTable_With_ProfitMargin.xlsx");
    }
}