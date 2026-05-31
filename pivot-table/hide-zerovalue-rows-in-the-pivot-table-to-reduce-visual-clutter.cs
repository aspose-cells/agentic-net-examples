using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotHideZeroRows
{
    public class HideZeroValueRowsDemo
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data with some zero sales values
            cells["A1"].PutValue("Region");
            cells["B1"].PutValue("Product");
            cells["C1"].PutValue("Sales");

            cells["A2"].PutValue("North");
            cells["B2"].PutValue("Widget");
            cells["C2"].PutValue(1500);

            cells["A3"].PutValue("North");
            cells["B3"].PutValue("Gadget");
            cells["C3"].PutValue(0); // Zero sales – should be hidden

            cells["A4"].PutValue("South");
            cells["B4"].PutValue("Widget");
            cells["C4"].PutValue(2000);

            cells["A5"].PutValue("South");
            cells["B5"].PutValue("Gadget");
            cells["C5"].PutValue(0); // Zero sales – should be hidden

            // Build a dictionary that holds total sales per region (used for hiding rows)
            var regionSales = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
            for (int row = 2; row <= 5; row++)
            {
                string region = cells[$"A{row}"].StringValue;
                double sales = cells[$"C{row}"].DoubleValue;
                if (!regionSales.ContainsKey(region))
                    regionSales[region] = 0;
                regionSales[region] += sales;
            }

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot table: Region as row field, Product as column field, Sales as data field
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Region");
            pivotTable.AddFieldToArea(PivotFieldType.Column, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Refresh and calculate the pivot data so that aggregated values are available
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Hide row items whose aggregated sales value is zero
            PivotField rowField = pivotTable.RowFields[0];
            for (int i = 0; i < rowField.ItemCount; i++)
            {
                string itemName = rowField.Items[i];
                // Retrieve the pre‑calculated total sales for the region
                double totalSales = regionSales.ContainsKey(itemName) ? regionSales[itemName] : 0;

                // Hide the item if its aggregated sales are zero (or near zero)
                if (Math.Abs(totalSales) < 0.0001)
                {
                    rowField.PivotItems[itemName].IsHidden = true;
                }
            }

            // Recalculate after hiding items
            pivotTable.CalculateData();

            // Save the workbook
            string outputPath = "PivotTable_HideZeroRows.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}