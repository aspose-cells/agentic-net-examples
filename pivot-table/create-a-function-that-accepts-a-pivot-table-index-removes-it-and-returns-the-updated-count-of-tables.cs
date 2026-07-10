using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotDemo
{
    public static class PivotTableHelper
    {
        /// <summary>
        /// Creates a workbook with sample data and three pivot tables,
        /// removes the pivot table at the specified index,
        /// and returns the updated count of pivot tables on the first worksheet.
        /// </summary>
        /// <param name="removeIndex">Zero‑based index of the pivot table to remove.</param>
        /// <returns>Number of remaining pivot tables after removal.</returns>
        public static int RemovePivotTableAtIndex(int removeIndex)
        {
            // ---------- Create ----------
            // Initialize a new workbook and get the first worksheet.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for pivot tables.
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["A4"].PutValue("Apple");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(200);

            // Add three pivot tables to the worksheet.
            sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
            sheet.PivotTables.Add("A1:B4", "D10", "PivotTable2");
            sheet.PivotTables.Add("A1:B4", "D20", "PivotTable3");

            // ---------- Remove ----------
            // Validate the index before attempting removal.
            if (removeIndex < 0 || removeIndex >= sheet.PivotTables.Count)
                throw new ArgumentOutOfRangeException(nameof(removeIndex), "Invalid pivot table index.");

            // Use the RemoveAt method from PivotTableCollection to delete the pivot table.
            sheet.PivotTables.RemoveAt(removeIndex);

            // ---------- Return ----------
            // Return the updated count of pivot tables.
            int remainingCount = sheet.PivotTables.Count;

            // ---------- Save (optional) ----------
            // Save the workbook to verify the result if needed.
            // workbook.Save("PivotTableRemovalResult.xlsx");

            return remainingCount;
        }

        // Example usage
        public static void Main()
        {
            int indexToRemove = 1; // remove the second pivot table (zero‑based)
            int remaining = RemovePivotTableAtIndex(indexToRemove);
            Console.WriteLine($"Remaining pivot tables after removal: {remaining}");
        }
    }
}