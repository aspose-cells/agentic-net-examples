using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotBatchDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add multiple worksheets and populate each with sample data
            for (int i = 0; i < 3; i++)
            {
                // Create or get worksheet
                Worksheet sheet = i == 0 ? workbook.Worksheets[0] : workbook.Worksheets.Add($"Sheet{i + 1}");

                // Populate sample data (A1:C10)
                Cells cells = sheet.Cells;
                cells["A1"].PutValue("Category");
                cells["B1"].PutValue("Item");
                cells["C1"].PutValue("Amount");

                for (int row = 2; row <= 10; row++)
                {
                    cells[$"A{row}"].PutValue($"Cat{(row % 3) + 1}");
                    cells[$"B{row}"].PutValue($"Item{row}");
                    cells[$"C{row}"].PutValue(row * 10);
                }

                // Define source data range string (including sheet name)
                string sourceData = $"={sheet.Name}!A1:C10";

                // Destination cell for the pivot table
                string destCell = "E1";

                // Unique pivot table name
                string pivotName = $"Pivot_{sheet.Name}";

                // Add the pivot table to the current worksheet
                int pivotIndex = sheet.PivotTables.Add(sourceData, destCell, pivotName);
                PivotTable pivot = sheet.PivotTables[pivotIndex];

                // Configure the pivot table (Category as row, Amount as data)
                pivot.AddFieldToArea(PivotFieldType.Row, "Category");
                pivot.AddFieldToArea(PivotFieldType.Data, "Amount");
                pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;
            }

            // Refresh all pivot tables in the workbook (optional but ensures data is up‑to‑date)
            workbook.Worksheets.RefreshPivotTables();

            // Save the workbook
            workbook.Save("BatchPivotTables.xlsx");
        }
    }
}