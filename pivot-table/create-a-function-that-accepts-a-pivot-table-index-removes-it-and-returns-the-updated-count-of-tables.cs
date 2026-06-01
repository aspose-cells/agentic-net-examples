using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotDemo
{
    public class PivotTableHelper
    {
        /// <summary>
        /// Removes the pivot table at the specified index from the given worksheet
        /// and returns the updated number of pivot tables on that worksheet.
        /// </summary>
        /// <param name="worksheet">Worksheet containing the pivot tables.</param>
        /// <param name="index">Zero‑based index of the pivot table to remove.</param>
        /// <returns>Count of remaining pivot tables after removal.</returns>
        public static int RemovePivotTableAtIndex(Worksheet worksheet, int index)
        {
            // Validate the index against the current collection count.
            if (index < 0 || index >= worksheet.PivotTables.Count)
                throw new ArgumentOutOfRangeException(nameof(index), "Invalid pivot table index.");

            // Use the documented RemoveAt(int) method of PivotTableCollection.
            worksheet.PivotTables.RemoveAt(index);

            // Return the new count.
            return worksheet.PivotTables.Count;
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook.
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

            // Add three pivot tables.
            sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
            sheet.PivotTables.Add("A1:B4", "D10", "PivotTable2");
            sheet.PivotTables.Add("A1:B4", "D20", "PivotTable3");

            // Remove the second pivot table (index 1) and get the new count.
            int remainingCount = PivotTableHelper.RemovePivotTableAtIndex(sheet, 1);

            Console.WriteLine("Remaining pivot tables count: " + remainingCount);

            // Save the workbook (uses the standard Save method).
            workbook.Save("PivotTableRemovalResult.xlsx");
        }
    }
}