using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotOdsDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for the pivot table
            cells["A1"].Value = "Category";
            cells["B1"].Value = "Amount";
            cells["A2"].Value = "Food";
            cells["B2"].Value = 1200;
            cells["A3"].Value = "Food";
            cells["B3"].Value = 800;
            cells["A4"].Value = "Travel";
            cells["B4"].Value = 1500;
            cells["A5"].Value = "Travel";
            cells["B5"].Value = 700;

            // Add a pivot table based on the data range A1:B5, place it at D1
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "D1", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot table: Category as row field, Amount as data field
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Column 0 -> Category
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Column 1 -> Amount

            // Save the workbook as an ODS file using SaveFormat.ODS
            workbook.Save("SalesPivotTable.ods", SaveFormat.Ods);
        }
    }
}