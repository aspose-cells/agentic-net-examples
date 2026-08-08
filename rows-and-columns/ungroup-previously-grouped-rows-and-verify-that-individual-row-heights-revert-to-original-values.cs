// Title: Ungroup rows and restore original row heights with Aspose.Cells for .NET
// Description: This C# example creates a workbook, assigns custom heights to rows 2‑4, groups the rows, then uses Cells.UngroupRows to remove the group and verifies that each row's height matches the values saved before grouping. The result is printed to the console and the workbook is saved.
// Keywords: Aspose.Cells | C# ungroup rows | row height restoration | Cells.UngroupRows | group rows Aspose.Cells | verify row height | Excel outline rows | Aspose.Cells .NET example
// Common Searches: Aspose.Cells ungroup rows C# | restore row height after grouping Aspose.Cells | check row height after ungrouping Excel | C# code to group and ungroup rows Aspose.Cells | verify row heights in Aspose.Cells workbook
// Developer Intent: The developer needs to remove a previously created row group and confirm that each row’s height reverts to the original values set before grouping.
// Use Cases: Apply custom heights to a set of rows, group them for outline display, then ungroup while preserving the original formatting. | Generate an Excel report where rows are temporarily collapsed for presentation and later expanded without altering row dimensions. | Automate a regression test to ensure that grouping and ungrouping operations do not modify row height properties in generated files.
// AI Prompts: Write a C# snippet using Aspose.Cells that groups rows 5‑10, then ungroups them and asserts that the row heights are identical to the pre‑group values. | Explain how Aspose.Cells manages row height preservation during grouping and ungrouping, and outline best practices for validating this behavior. | Create an MSTest unit test that sets distinct heights for rows, groups and ungroups them with Aspose.Cells, and verifies that the heights are restored.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // This C# example creates a workbook, assigns custom heights to rows 2‑4, groups the rows, then uses Cells.UngroupRows to remove the group and verifies that each row's height matches the values saved before grouping. The result is printed to the console and the workbook is saved.
    public class UngroupRowsVerificationDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Define the rows to work with (zero‑based indices 1 to 3 correspond to Excel rows 2‑4)
            int firstRow = 1;
            int lastRow = 3;

            // Set distinct heights for each row and store the original heights
            double[] originalHeights = new double[lastRow - firstRow + 1];
            double[] newHeights = new double[] { 20.0, 25.0, 30.0 };

            for (int i = firstRow; i <= lastRow; i++)
            {
                int idx = i - firstRow;
                cells.SetRowHeight(i, newHeights[idx]);          // Apply a custom height
                originalHeights[idx] = cells.GetRowHeight(i);   // Store the height (should be the same as newHeights)
                cells[i, 0].PutValue($"Row {i + 1}");           // Add some data for visibility
            }

            // Group the rows (the grouping itself does not alter row heights)
            cells.GroupRows(firstRow, lastRow, false);

            // Ungroup the rows using the UngroupRows method
            cells.UngroupRows(firstRow, lastRow);

            // Verify that each row's height matches the original height after ungrouping
            bool allMatch = true;
            for (int i = firstRow; i <= lastRow; i++)
            {
                int idx = i - firstRow;
                double currentHeight = cells.GetRowHeight(i);
                if (Math.Abs(currentHeight - originalHeights[idx]) > 0.001)
                {
                    allMatch = false;
                    Console.WriteLine($"Row {i + 1} height mismatch. Expected: {originalHeights[idx]}, Actual: {currentHeight}");
                }
                else
                {
                    Console.WriteLine($"Row {i + 1} height correctly restored to {currentHeight}");
                }
            }

            Console.WriteLine(allMatch
                ? "All row heights reverted to original values after ungrouping."
                : "Some row heights did not revert to original values.");

            // Save the workbook to demonstrate the final state (optional)
            workbook.Save("UngroupRowsVerificationDemo.xlsx");
        }
    }
}
