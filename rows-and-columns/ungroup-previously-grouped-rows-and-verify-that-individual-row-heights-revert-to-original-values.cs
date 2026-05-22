using System;
using Aspose.Cells;

namespace AsposeCellsUngroupRowsDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Define the rows to work with (rows 2 to 4, zero‑based indexes 1‑3)
            int firstRow = 1;
            int lastRow = 3;
            int rowCount = lastRow - firstRow + 1;

            // Set custom heights for the rows and store the original heights
            double[] originalHeights = new double[rowCount];
            for (int i = 0; i < rowCount; i++)
            {
                int rowIndex = firstRow + i;
                // Assign a distinct height for demonstration
                cells.SetRowHeight(rowIndex, 20 + i * 5);
                // Store the height value
                originalHeights[i] = cells.GetRowHeight(rowIndex);
            }

            // Group the rows (the group will be hidden for visual effect)
            cells.GroupRows(firstRow, lastRow, true);

            // At this point the rows are hidden; now ungroup them
            // Using the overload that removes only the outer group info
            cells.UngroupRows(firstRow, lastRow);

            // Verify that each row's height matches the original value
            bool allMatch = true;
            for (int i = 0; i < rowCount; i++)
            {
                int rowIndex = firstRow + i;
                double currentHeight = cells.GetRowHeight(rowIndex);
                if (Math.Abs(currentHeight - originalHeights[i]) > 0.001)
                {
                    allMatch = false;
                    Console.WriteLine($"Row {rowIndex + 1} height mismatch. Expected: {originalHeights[i]}, Actual: {currentHeight}");
                }
            }

            if (allMatch)
            {
                Console.WriteLine("All row heights reverted to their original values after ungrouping.");
            }

            // Save the workbook to demonstrate the final state
            workbook.Save("UngroupRowsResult.xlsx");
        }
    }
}