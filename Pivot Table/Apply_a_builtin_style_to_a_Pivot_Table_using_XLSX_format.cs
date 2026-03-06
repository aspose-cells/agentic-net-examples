using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotStyleDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["A4"].PutValue("Orange");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["B2"].PutValue(1000);
            sheet.Cells["B3"].PutValue(2000);
            sheet.Cells["B4"].PutValue(3000);

            // Add a pivot table based on the data range A1:B4, place it at E3
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "E3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot table: Product as row field, Sales as data field
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Column index 0 -> Product
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Column index 1 -> Sales

            // Apply a built‑in pivot table style by name (e.g., "PivotStyleLight16")
            pivotTable.PivotTableStyleName = "PivotStyleLight16";

            // Save the workbook in XLSX format
            workbook.Save("PivotTableWithBuiltInStyle.xlsx", SaveFormat.Xlsx);
        }
    }
}