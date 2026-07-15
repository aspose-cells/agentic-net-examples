using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotSaveDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            Cells cells = sheet.Cells;
            cells["A1"].Value = "Category";
            cells["B1"].Value = "Amount";
            cells["A2"].Value = "Food";
            cells["B2"].Value = 120;
            cells["A3"].Value = "Food";
            cells["B3"].Value = 80;
            cells["A4"].Value = "Beverage";
            cells["B4"].Value = 150;
            cells["A5"].Value = "Beverage";
            cells["B5"].Value = 200;

            // Add a pivot table based on the data range A1:B5, place it at C3
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "C3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot table: Category as row field, Amount as data field
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Save the workbook to XLSX format using default save options
            workbook.Save("SalesPivotTable.xlsx");
        }
    }
}