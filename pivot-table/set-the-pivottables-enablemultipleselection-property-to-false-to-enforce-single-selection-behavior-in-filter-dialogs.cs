using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class PivotTableEnableMultipleSelectionDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Category");
                sheet.Cells["C1"].PutValue("Sales");

                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["B2"].PutValue("Fruit");
                sheet.Cells["C2"].PutValue(1000);

                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["B3"].PutValue("Fruit");
                sheet.Cells["C3"].PutValue(1500);

                sheet.Cells["A4"].PutValue("Carrot");
                sheet.Cells["B4"].PutValue("Vegetable");
                sheet.Cells["C4"].PutValue(800);

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:C4", "E3", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Add fields to the pivot table
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
                pivotTable.AddFieldToArea(PivotFieldType.Column, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Disable multiple selection in filter dialogs
                pivotTable.AllowMultipleFiltersPerField = false;

                // Refresh and calculate the pivot table to apply changes
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Save the workbook
                string outputPath = "PivotTable_EnableMultipleSelection.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
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
            PivotTableEnableMultipleSelectionDemo.Run();
        }
    }
}