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

        // Populate sample data with Revenue and Cost columns
        cells["A1"].Value = "Product";
        cells["B1"].Value = "Region";
        cells["C1"].Value = "Revenue";
        cells["D1"].Value = "Cost";

        string[] products = { "A", "B", "A", "B" };
        string[] regions = { "North", "North", "South", "South" };
        double[] revenues = { 1000, 1500, 1200, 1800 };
        double[] costs = { 600, 900, 700, 1100 };

        for (int i = 0; i < products.Length; i++)
        {
            int row = i + 2;
            cells[$"A{row}"].Value = products[i];
            cells[$"B{row}"].Value = regions[i];
            cells[$"C{row}"].Value = revenues[i];
            cells[$"D{row}"].Value = costs[i];
        }

        // Add a pivot table covering the data range
        int pivotIndex = sheet.PivotTables.Add("A1:D5", "F3", "SalesPivot");
        PivotTable pivotTable = sheet.PivotTables[pivotIndex];

        // Add fields to the pivot table areas
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
        pivotTable.AddFieldToArea(PivotFieldType.Column, "Region");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Revenue");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Cost");

        // Add a calculated field that computes profit margin: (Revenue - Cost) / Revenue
        pivotTable.AddCalculatedField("ProfitMargin", "=(Revenue-Cost)/Revenue", true);

        // Format the calculated field as a percentage
        PivotField profitMarginField = pivotTable.DataFields[pivotTable.DataFields.Count - 1];
        profitMarginField.NumberFormat = "0.00%";

        // Refresh and calculate the pivot table
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Save the workbook
        workbook.Save("PivotWithProfitMargin.xlsx");
    }
}