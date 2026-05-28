using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data (Product, Region, Sales, Quantity)
                cells["A1"].Value = "Product";
                cells["B1"].Value = "Region";
                cells["C1"].Value = "Sales";
                cells["D1"].Value = "Quantity";

                string[] products = { "Bike", "Car", "Bike", "Car", "Bike", "Car" };
                string[] regions = { "North", "North", "South", "South", "East", "East" };
                double[] sales = { 1000, 2000, 1500, 2500, 1200, 2200 };
                int[] qty = { 10, 20, 15, 25, 12, 22 };

                for (int i = 0; i < products.Length; i++)
                {
                    int row = i + 2;
                    cells[$"A{row}"].Value = products[i];
                    cells[$"B{row}"].Value = regions[i];
                    cells[$"C{row}"].Value = sales[i];
                    cells[$"D{row}"].Value = qty[i];
                }

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:D7", "F3", "SalesPivot");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Configure pivot fields
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
                pivotTable.AddFieldToArea(PivotFieldType.Column, "Region");

                // Add two data fields: Sales and Quantity
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Quantity");

                // Aspose.Cells does not expose a DataFieldSeparator property.
                // Custom separators can be applied via formatting after the pivot is generated if required.

                // Refresh and calculate the pivot table to apply changes
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Save the workbook
                string outputPath = "PivotTable_With_Custom_DataFieldSeparator.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}