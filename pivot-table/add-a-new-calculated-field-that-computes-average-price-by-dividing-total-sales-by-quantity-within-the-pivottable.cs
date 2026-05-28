using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class AddAveragePriceCalculatedField
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                // Columns: Product, Sales, Quantity
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["C1"].PutValue("Quantity");

                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["B2"].PutValue(1200);
                sheet.Cells["C2"].PutValue(30);

                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["B3"].PutValue(800);
                sheet.Cells["C3"].PutValue(20);

                sheet.Cells["A4"].PutValue("Orange");
                sheet.Cells["B4"].PutValue(1500);
                sheet.Cells["C4"].PutValue(50);

                // Add a pivot table based on the data range A1:C4, place it at E3
                int pivotIndex = sheet.PivotTables.Add("A1:C4", "E3", "SalesPivot");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Add Product as a row field
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");

                // Add Sales and Quantity as data fields
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Quantity");

                // Add a calculated field that computes average price = Sales / Quantity
                pivotTable.AddCalculatedField("AveragePrice", "=Sales/Quantity");

                // Refresh and calculate the pivot table to apply the new field
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Save the workbook with the pivot table and calculated field
                workbook.Save("PivotTable_With_AveragePrice.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            AddAveragePriceCalculatedField.Run();
        }
    }
}