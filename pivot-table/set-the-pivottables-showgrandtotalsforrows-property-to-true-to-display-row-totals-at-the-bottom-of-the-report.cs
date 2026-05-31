using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class ShowGrandTotalsForRowsDemo
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            Cells cells = sheet.Cells;
            cells["A1"].Value = "Category";
            cells["B1"].Value = "Item";
            cells["C1"].Value = "Amount";

            cells["A2"].Value = "Food";
            cells["B2"].Value = "Apple";
            cells["C2"].Value = 120;

            cells["A3"].Value = "Food";
            cells["B3"].Value = "Banana";
            cells["C3"].Value = 80;

            cells["A4"].Value = "Beverage";
            cells["B4"].Value = "Tea";
            cells["C4"].Value = 150;

            cells["A5"].Value = "Beverage";
            cells["B5"].Value = "Coffee";
            cells["C5"].Value = 200;

            // Add a pivot table based on the data range
            PivotTableCollection pivotTables = sheet.PivotTables;
            int pivotIndex = pivotTables.Add("A1:C5", "E3", "PivotTable1");
            PivotTable pivotTable = pivotTables[pivotIndex];

            // Add fields to the pivot table
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Category as row field
            pivotTable.AddFieldToArea(PivotFieldType.Row, 1);   // Item as row field
            pivotTable.AddFieldToArea(PivotFieldType.Data, 2);  // Amount as data field

            // Enable grand totals for rows (display row totals at the bottom)
            pivotTable.ShowRowGrandTotals = true;

            // Save the workbook
            workbook.Save("ShowGrandTotalsForRowsDemo.xlsx");
        }
    }
}