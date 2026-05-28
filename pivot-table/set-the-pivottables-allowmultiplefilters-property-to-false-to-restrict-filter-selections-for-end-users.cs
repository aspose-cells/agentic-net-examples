using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class PivotTableAllowMultipleFiltersDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                Cells cells = sheet.Cells;
                cells["A1"].Value = "Category";
                cells["B1"].Value = "Amount";
                cells["A2"].Value = "Fruit";
                cells["B2"].Value = 100;
                cells["A3"].Value = "Vegetable";
                cells["B3"].Value = 150;
                cells["A4"].Value = "Fruit";
                cells["B4"].Value = 200;
                cells["A5"].Value = "Vegetable";
                cells["B5"].Value = 250;

                // Add a pivot table to the worksheet
                int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Add fields to the pivot table
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

                // Restrict multiple filters per field for end users
                pivotTable.AllowMultipleFiltersPerField = false;

                // Save the workbook with the configured pivot table
                string outputPath = "PivotTable_AllowMultipleFiltersFalse.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point required for console application
    public class Program
    {
        public static void Main(string[] args)
        {
            PivotTableAllowMultipleFiltersDemo.Run();
        }
    }
}