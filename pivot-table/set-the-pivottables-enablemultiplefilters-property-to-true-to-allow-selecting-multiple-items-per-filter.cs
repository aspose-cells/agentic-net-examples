using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotMultipleFiltersDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            Cells cells = sheet.Cells;
            cells["A1"].Value = "Category";
            cells["B1"].Value = "Amount";
            cells["A2"].Value = "Food";
            cells["B2"].Value = 120;
            cells["A3"].Value = "Drink";
            cells["B3"].Value = 80;
            cells["A4"].Value = "Food";
            cells["B4"].Value = 150;
            cells["A5"].Value = "Drink";
            cells["B5"].Value = 60;

            // Add a pivot table based on the data range A1:B5, place it at C3, and give it a name
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "C3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot table fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");   // Row field
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");   // Data field

            // Enable multiple filters per field for the pivot table
            pivotTable.AllowMultipleFiltersPerField = true;

            // Save the workbook to a file
            workbook.Save("PivotTable_EnableMultipleFilters.xlsx");

            // Optional: output confirmation
            Console.WriteLine("PivotTable created with AllowMultipleFiltersPerField = true.");
        }
    }
}