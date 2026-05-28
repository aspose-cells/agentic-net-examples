using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class PivotValueFilterDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data: Product and Sales
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["A2"].PutValue("WidgetA");
                sheet.Cells["B2"].PutValue(150);
                sheet.Cells["A3"].PutValue("WidgetB");
                sheet.Cells["B3"].PutValue(80);
                sheet.Cells["A4"].PutValue("WidgetC");
                sheet.Cells["B4"].PutValue(200);
                sheet.Cells["A5"].PutValue("WidgetD");
                sheet.Cells["B5"].PutValue(45);

                // Create a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "SalesPivot");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Add Product as row field and Sales as data field
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Define the sales threshold
                double salesThreshold = 100.0;

                // Apply a value filter to show only items with Sales > threshold
                // baseFieldIndex = 0 (Product field), valueFieldIndex = 1 (Sales data field)
                PivotFilterCollection filters = pivotTable.PivotFilters;
                filters.AddValueFilter(
                    baseFieldIndex: 0,
                    valueFieldIndex: 1,
                    type: PivotFilterType.ValueGreaterThan,
                    value1: salesThreshold,
                    value2: 0); // value2 is ignored for ValueGreaterThan

                // Refresh the pivot table to apply the filter
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Save the workbook
                string outputPath = "PivotValueFilterDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
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
            PivotValueFilterDemo.Run();
        }
    }
}