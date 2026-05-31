using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class PivotItemMoveUpwardDemo
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
                sheet.Cells["A2"].PutValue("Fruit");
                sheet.Cells["A3"].PutValue("Vegetable");
                sheet.Cells["A4"].PutValue("Fruit");
                sheet.Cells["A5"].PutValue("Vegetable");

                sheet.Cells["B1"].PutValue("Amount");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(80);
                sheet.Cells["B4"].PutValue(150);
                sheet.Cells["B5"].PutValue(70);

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Add the "Category" field to the row area
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");

                // Add the "Amount" field to the data area
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

                // Refresh and calculate to populate the pivot table
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Get the PivotItemCollection for the row field
                PivotItemCollection items = pivotTable.RowFields["Category"].PivotItems;

                // Move the second item ("Vegetable") up by one position if possible
                if (items.Count > 1)
                {
                    items[1].Move(-1, true);
                }

                // Recalculate after moving the item
                pivotTable.CalculateData();

                // Save the workbook
                string outputPath = "PivotItemMoveUpwardDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
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
            PivotItemMoveUpwardDemo.Run();
        }
    }
}