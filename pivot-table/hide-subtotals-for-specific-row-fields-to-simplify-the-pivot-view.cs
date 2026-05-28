using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotHideSubtotals
{
    public class HideSubtotalsDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data
                // Columns: Category, Product, Sales
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Product");
                sheet.Cells["C1"].PutValue("Sales");

                sheet.Cells["A2"].PutValue("Electronics");
                sheet.Cells["B2"].PutValue("TV");
                sheet.Cells["C2"].PutValue(1200);

                sheet.Cells["A3"].PutValue("Electronics");
                sheet.Cells["B3"].PutValue("Radio");
                sheet.Cells["C3"].PutValue(800);

                sheet.Cells["A4"].PutValue("Clothing");
                sheet.Cells["B4"].PutValue("Shirt");
                sheet.Cells["C4"].PutValue(500);

                sheet.Cells["A5"].PutValue("Clothing");
                sheet.Cells["B5"].PutValue("Pants");
                sheet.Cells["C5"].PutValue(700);

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Add row fields: Category and Product
                int categoryRowIndex = pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                int productRowIndex = pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");

                // Add data field: Sales
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Retrieve the PivotField for "Category" (first row field)
                PivotField categoryField = pivotTable.RowFields[categoryRowIndex];

                // Hide all automatic subtotals for the "Category" field
                categoryField.IsAutoSubtotals = false;

                // Ensure specific subtotal types are not shown
                categoryField.SetSubtotals(PivotFieldSubtotalType.Sum, false);
                categoryField.SetSubtotals(PivotFieldSubtotalType.Average, false);
                categoryField.SetSubtotals(PivotFieldSubtotalType.Count, false);
                categoryField.SetSubtotals(PivotFieldSubtotalType.Max, false);
                categoryField.SetSubtotals(PivotFieldSubtotalType.Min, false);
                categoryField.SetSubtotals(PivotFieldSubtotalType.Product, false);
                // StdDev and Var types are not available in this version; omitted.

                // Refresh and calculate the pivot table to apply changes
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Save the workbook
                string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "PivotHideSubtotalsDemo.xlsx");
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Entry point for the console application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}