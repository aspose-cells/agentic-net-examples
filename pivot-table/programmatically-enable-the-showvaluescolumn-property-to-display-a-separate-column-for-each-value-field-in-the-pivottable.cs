using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class EnableShowValuesColumnDemo
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
                sheet.Cells["D1"].PutValue("Quantity");

                sheet.Cells["A2"].PutValue("Fruit");
                sheet.Cells["B2"].PutValue("Apple");
                sheet.Cells["C2"].PutValue(1200);
                sheet.Cells["D2"].PutValue(10);

                sheet.Cells["A3"].PutValue("Fruit");
                sheet.Cells["B3"].PutValue("Orange");
                sheet.Cells["C3"].PutValue(800);
                sheet.Cells["D3"].PutValue(15);

                sheet.Cells["A4"].PutValue("Vegetable");
                sheet.Cells["B4"].PutValue("Carrot");
                sheet.Cells["C4"].PutValue(500);
                sheet.Cells["D4"].PutValue(20);

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:D4", "F2", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Configure the pivot table: Category as row, Product as column,
                // and both Sales and Quantity as data fields
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Column, "Product");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Quantity");

                // NOTE: The ShowValuesColumn property is not available in older Aspose.Cells versions.
                // If supported, it can be enabled as shown below:
                // pivotTable.ShowValuesColumn = true;

                // Refresh and calculate the pivot table data
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Define output file path
                string outputPath = "EnableShowValuesColumnDemo.xlsx";

                // Save the workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    internal class Program
    {
        private static void Main(string[] args)
        {
            try
            {
                EnableShowValuesColumnDemo.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}