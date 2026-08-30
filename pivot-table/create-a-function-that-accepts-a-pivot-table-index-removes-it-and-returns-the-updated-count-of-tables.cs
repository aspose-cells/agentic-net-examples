// Title: C# method to remove a pivot table by index and get remaining pivot table count with Aspose.Cells
// AI Prompts: Write a static C# function that accepts a Workbook and a zero‑based pivot table index, validates the index, removes the pivot table from the first worksheet using Worksheet.PivotTables.RemoveAt, and returns the updated PivotTables.Count. | Show how to invoke the removal function, catch an ArgumentOutOfRangeException for an invalid index, and save the workbook after the pivot table has been deleted.
// Common Searches: C# Aspose.Cells how to delete a specific pivot table using its index | example of removing a pivot table and checking remaining count in Aspose.Cells | Aspose.Cells PivotTables.RemoveAt method usage in .NET
// Tags: Aspose.Cells Worksheet.PivotTables.RemoveAt example | C# delete pivot table by index | pivot table count after removal Aspose.Cells | handle invalid pivot table index ArgumentOutOfRangeException | save workbook after pivot table deletion Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Provides a static helper that validates a zero‑based pivot table index, removes the corresponding pivot table from the first worksheet via Worksheet.PivotTables.RemoveAt, and returns the new PivotTables.Count. A demo creates a workbook, adds three pivot tables, removes the second one, prints the remaining count, and saves the file.
    public static class PivotTableHelper
    {
        /// <param name="workbook">The workbook containing the pivot tables.</param>
        /// <param name="pivotIndex">Zero‑based index of the pivot table to remove.</param>
        /// <returns>Number of remaining pivot tables after removal.</returns>
        public static int RemovePivotTableAtIndex(Workbook workbook, int pivotIndex)
        {
            // Access the first worksheet (adjust as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the index is within range
            if (pivotIndex < 0 || pivotIndex >= sheet.PivotTables.Count)
                throw new ArgumentOutOfRangeException(nameof(pivotIndex), "Invalid pivot table index.");

            // Remove the pivot table using the provided RemoveAt method
            sheet.PivotTables.RemoveAt(pivotIndex);

            // Return the updated count
            return sheet.PivotTables.Count;
        }
    }

    public class Demo
    {
        public static void Run()
        {
            try
            {
                // Create a workbook with sample data and multiple pivot tables
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Sample data
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["B2"].PutValue(100);
                sheet.Cells["B3"].PutValue(200);
                sheet.Cells["B4"].PutValue(300);

                // Add three pivot tables
                sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
                sheet.PivotTables.Add("A1:B4", "D10", "PivotTable2");
                sheet.PivotTables.Add("A1:B4", "D20", "PivotTable3");

                // Remove the second pivot table (index 1) and get the new count
                int newCount = PivotTableHelper.RemovePivotTableAtIndex(workbook, 1);
                Console.WriteLine("Remaining pivot tables count: " + newCount);

                // Save the workbook
                workbook.Save("PivotTableRemoved.xlsx");
                Console.WriteLine("Workbook saved as PivotTableRemoved.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
