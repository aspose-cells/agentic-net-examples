using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class CollapsePivotItemsDemo
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
                sheet.Cells["B1"].PutValue("Amount");
                sheet.Cells["A2"].PutValue("Fruit");
                sheet.Cells["B2"].PutValue(100);
                sheet.Cells["A3"].PutValue("Fruit");
                sheet.Cells["B3"].PutValue(150);
                sheet.Cells["A4"].PutValue("Vegetable");
                sheet.Cells["B4"].PutValue(200);
                sheet.Cells["A5"].PutValue("Vegetable");
                sheet.Cells["B5"].PutValue(250);

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Add the "Category" field as a row field and "Amount" as a data field
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

                // Refresh and calculate the pivot table to populate items
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Collapse each row item by hiding its detail
                PivotField rowField = pivotTable.RowFields[0];
                foreach (PivotItem item in rowField.PivotItems)
                {
                    item.IsDetailHidden = true;
                }

                // Recalculate after changing item states
                pivotTable.CalculateData();

                // Save the workbook
                string outputPath = "CollapsedPivotItemsDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            CollapsePivotItemsDemo.Run();
        }
    }
}