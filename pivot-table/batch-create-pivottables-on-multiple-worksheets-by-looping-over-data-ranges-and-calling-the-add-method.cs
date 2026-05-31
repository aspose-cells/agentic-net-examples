using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class BatchPivotTables
{
    static void Main()
    {
        try
        {
            // Create a new workbook and ensure it has exactly three worksheets
            Workbook workbook = new Workbook();
            WorksheetCollection sheets = workbook.Worksheets;
            sheets.Clear(); // remove the default sheet

            for (int i = 0; i < 3; i++)
            {
                // Add a new worksheet with a specific name
                Worksheet sheet = sheets.Add($"DataSheet{i + 1}");

                // Header
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Item");
                sheet.Cells["C1"].PutValue("Amount");

                // Sample rows
                for (int row = 2; row <= 5; row++)
                {
                    sheet.Cells[$"A{row}"].PutValue($"Cat{(row % 3) + 1}");
                    sheet.Cells[$"B{row}"].PutValue($"Item{row}");
                    sheet.Cells[$"C{row}"].PutValue(row * 10);
                }
            }

            // Create a pivot table on each worksheet
            foreach (Worksheet ws in workbook.Worksheets)
            {
                string sourceData = "A1:C5";
                string destCell = "E1";
                string pivotName = $"Pivot_{ws.Name}";

                // Add the pivot table (returns the index of the new pivot table)
                int pivotIndex = ws.PivotTables.Add(sourceData, destCell, pivotName);
                PivotTable pivot = ws.PivotTables[pivotIndex];

                // Configure fields
                pivot.AddFieldToArea(PivotFieldType.Row, "Category");
                // Additional configuration (e.g., column, data fields) can be added here
            }

            // Save the workbook
            string outputPath = "BatchPivotTables.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}