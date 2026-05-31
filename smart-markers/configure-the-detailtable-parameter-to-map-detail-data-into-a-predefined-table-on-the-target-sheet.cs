using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Tables;   // Required for ListObject

namespace AsposeCellsDetailTableDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // ---------- Populate source data ----------
                // Header
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Item");
                sheet.Cells["C1"].PutValue("Quantity");
                // Data rows
                sheet.Cells["A2"].PutValue("Fruit");
                sheet.Cells["B2"].PutValue("Apple");
                sheet.Cells["C2"].PutValue(10);
                sheet.Cells["A3"].PutValue("Fruit");
                sheet.Cells["B3"].PutValue("Banana");
                sheet.Cells["C3"].PutValue(15);
                sheet.Cells["A4"].PutValue("Vegetable");
                sheet.Cells["B4"].PutValue("Carrot");
                sheet.Cells["C4"].PutValue(20);
                sheet.Cells["A5"].PutValue("Vegetable");
                sheet.Cells["B5"].PutValue("Tomato");
                sheet.Cells["C5"].PutValue(25);

                // ---------- Create a pivot table ----------
                // The pivot will be placed starting at E3
                int pivotIdx = sheet.PivotTables.Add("A1:C5", "E3", "SalesPivot");
                PivotTable pivot = sheet.PivotTables[pivotIdx];

                // Row field: Category
                pivot.AddFieldToArea(PivotFieldType.Row, "Category");
                // Data field: Sum of Quantity
                pivot.AddFieldToArea(PivotFieldType.Data, "Quantity");

                // ---------- Prepare the target table (DetailTable) ----------
                // We'll place the detail table starting at A10 on the same sheet
                // Create a ListObject (Excel table) with enough rows/columns
                // Here we create a 5x3 table (adjust as needed)
                int tableIdx = sheet.ListObjects.Add(9, 0, 13, 2, true); // rows 10‑14, cols A‑C (zero‑based)
                ListObject detailTable = sheet.ListObjects[tableIdx];
                detailTable.DisplayName = "DetailTable";

                // Optional: set column headers for the detail table
                sheet.Cells["A10"].PutValue("Category");
                sheet.Cells["B10"].PutValue("Item");
                sheet.Cells["C10"].PutValue("Quantity");

                // ---------- Show detail of a pivot item into the predefined table ----------
                // rowOffset and columnOffset refer to the position of the first data row/column
                // in the pivot's data region (excluding headers). For this simple pivot,
                // the first data row is at offset 1 (row index 1 relative to the pivot's top‑left cell).
                // We'll map the detail of the first row item ("Fruit") into the table.
                int rowOffset = 1;      // first data row in pivot
                int columnOffset = 0;   // first data column in pivot
                bool newSheet = false;  // place detail on the same worksheet
                int destRow = 9;        // zero‑based row index where the table starts (A10)
                int destColumn = 0;     // zero‑based column index (A)

                // Use the ShowDetail method (rule‑based implementation)
                pivot.ShowDetail(rowOffset, columnOffset, newSheet, destRow, destColumn);

                // ---------- Save the workbook ----------
                string outputPath = "DetailTableMapped.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}