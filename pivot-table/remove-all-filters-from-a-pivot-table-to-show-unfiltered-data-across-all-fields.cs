using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class RemoveAllPivotFiltersDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Product");
                sheet.Cells["C1"].PutValue("Sales");
                sheet.Cells["A2"].PutValue("Fruit");
                sheet.Cells["B2"].PutValue("Apple");
                sheet.Cells["C2"].PutValue(120);
                sheet.Cells["A3"].PutValue("Fruit");
                sheet.Cells["B3"].PutValue("Banana");
                sheet.Cells["C3"].PutValue(80);
                sheet.Cells["A4"].PutValue("Vegetable");
                sheet.Cells["B4"].PutValue("Carrot");
                sheet.Cells["C4"].PutValue(150);
                sheet.Cells["A5"].PutValue("Vegetable");
                sheet.Cells["B5"].PutValue("Broccoli");
                sheet.Cells["C5"].PutValue(130);

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "SalesPivot");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Configure the pivot table: rows = Category, columns = Product, data = Sum of Sales
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Column, "Product");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Refresh to calculate initial data
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Apply a sample filter: show only "Fruit" in Category row field
                PivotField rowField = pivotTable.RowFields["Category"];
                rowField.FilterByLabel(PivotFilterType.CaptionEqual, "Fruit", null);

                // Refresh after applying filter
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Remove all filters from the pivot table
                foreach (PivotField field in pivotTable.RowFields)
                {
                    field.ClearFilter();
                }
                foreach (PivotField field in pivotTable.ColumnFields)
                {
                    field.ClearFilter();
                }
                foreach (PivotField field in pivotTable.PageFields)
                {
                    field.ClearFilter();
                }
                foreach (PivotField field in pivotTable.DataFields)
                {
                    field.ClearFilter();
                }

                // Refresh to reflect removal of filters
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Save the workbook
                string outputPath = "PivotTable_NoFilters.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            RemoveAllPivotFiltersDemo.Run();
        }
    }
}