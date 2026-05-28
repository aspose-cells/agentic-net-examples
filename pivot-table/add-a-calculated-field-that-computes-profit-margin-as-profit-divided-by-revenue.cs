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

        // Populate sample data: Region, Product, Revenue, Profit
        cells["A1"].Value = "Region";
        cells["B1"].Value = "Product";
        cells["C1"].Value = "Revenue";
        cells["D1"].Value = "Profit";

        // Sample rows
        object[,] data = {
            { "North", "A", 1000, 200 },
            { "North", "B", 1500, 300 },
            { "South", "A", 1200, 240 },
            { "South", "B", 1800, 360 }
        };

        for (int i = 0; i < data.GetLength(0); i++)
        {
            cells[i + 2, 0].Value = data[i, 0];
            cells[i + 2, 1].Value = data[i, 1];
            cells[i + 2, 2].Value = data[i, 2];
            cells[i + 2, 3].Value = data[i, 3];
        }

        // Add a pivot table based on the data range
        int pivotIndex = sheet.PivotTables.Add("A1:D5", "F3", "SalesPivot");
        PivotTable pivot = sheet.PivotTables[pivotIndex];

        // Add fields to the pivot layout
        pivot.AddFieldToArea(PivotFieldType.Row, "Region");
        pivot.AddFieldToArea(PivotFieldType.Column, "Product");
        pivot.AddFieldToArea(PivotFieldType.Data, "Revenue");
        pivot.AddFieldToArea(PivotFieldType.Data, "Profit");

        // Add a calculated field "ProfitMargin" = Profit / Revenue and place it in the data area
        pivot.AddCalculatedField("ProfitMargin", "=Profit/Revenue", true);

        // Format the calculated field as a percentage
        PivotField marginField = pivot.DataFields[pivot.DataFields.Count - 1];
        marginField.NumberFormat = "0.00%";

        // Refresh and calculate the pivot table
        pivot.RefreshData();
        pivot.CalculateData();

        // Save the workbook
        workbook.Save("PivotWithProfitMargin.xlsx");
    }
}