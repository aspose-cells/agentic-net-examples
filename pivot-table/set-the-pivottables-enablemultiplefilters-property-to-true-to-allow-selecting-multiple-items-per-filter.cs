using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotMultipleFiltersDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            Cells cells = sheet.Cells;
            cells["A1"].Value = "Category";
            cells["B1"].Value = "Amount";
            cells["A2"].Value = "Fruit";
            cells["B2"].Value = 120;
            cells["A3"].Value = "Vegetable";
            cells["B3"].Value = 80;
            cells["A4"].Value = "Fruit";
            cells["B4"].Value = 150;
            cells["A5"].Value = "Grain";
            cells["B5"].Value = 60;
            cells["A6"].Value = "Vegetable";
            cells["B6"].Value = 90;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B6", "D3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot table fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Enable multiple filters per field (allows selecting multiple items in a filter)
            pivotTable.AllowMultipleFiltersPerField = true;

            // Save the workbook
            workbook.Save("PivotTable_MultipleFilters.xlsx");

            Console.WriteLine("Pivot table created with AllowMultipleFiltersPerField set to true.");
        }
    }
}