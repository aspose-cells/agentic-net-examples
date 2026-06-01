using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class PivotRowFieldTop10Filter
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                sheet.Cells["A1"].Value = "Category";
                sheet.Cells["B1"].Value = "SubCategory";
                sheet.Cells["C1"].Value = "Sales";

                string[] categories = { "Fruit", "Fruit", "Fruit", "Vegetable", "Vegetable", "Dairy", "Dairy", "Dairy" };
                string[] subCategories = { "Apple", "Orange", "Banana", "Carrot", "Broccoli", "Milk", "Cheese", "Yogurt" };
                int[] sales = { 120, 80, 150, 60, 40, 200, 180, 90 };

                for (int i = 0; i < categories.Length; i++)
                {
                    sheet.Cells[i + 1, 0].Value = categories[i];
                    sheet.Cells[i + 1, 1].Value = subCategories[i];
                    sheet.Cells[i + 1, 2].Value = sales[i];
                }

                // Create a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:C9", "E3", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Add Category as a row field
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");

                // Add Sales as a data field (sum)
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Apply a Top 10 filter on the row field to show only the highest values
                pivotTable.BaseFields[0].FilterTop10(
                    valueFieldIndex: 0,
                    type: PivotFilterType.Sum,
                    isTop: true,
                    itemCount: 5);

                // Save the workbook
                string outputPath = "PivotRowFieldTop10Filter.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
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
            PivotRowFieldTop10Filter.Run();
        }
    }
}