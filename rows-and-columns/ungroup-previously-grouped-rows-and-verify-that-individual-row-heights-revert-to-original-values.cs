// Title: How to ungroup rows in Aspose.Cells for .NET and verify that each row’s original height is restored
// AI Prompts: Create a C# program that groups rows 2‑5 in a worksheet, stores each row’s height, then ungroups the rows and checks that the heights match the stored values using Aspose.Cells. | Write code that uses Cells.GroupRows and Cells.UngroupRows to hide rows, iterates through the affected rows, compares GetRowHeight with previously saved heights, and reports any mismatches.
// Common Searches: Aspose.Cells C# ungroup rows and keep original row heights | verify row height after ungrouping rows in a .NET workbook | how to restore individual row heights after grouping rows with Aspose.Cells
// Tags: Aspose.Cells group rows C# | Aspose.Cells ungroup rows C# | row height restoration after grouping Aspose.Cells | verify individual row heights .NET workbook | Cells.SetRowHeight usage with grouping

using System;
using Aspose.Cells;

// Demonstrates assigning distinct heights to rows, grouping them, ungrouping them, and programmatically confirming that each row’s height returns to its original value before saving the workbook.
public class UngroupRowsDemo
{
    public static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Define the range of rows to group and later ungroup
        int firstRow = 2;   // zero‑based index
        int lastRow = 5;    // zero‑based index

        // Set distinct heights for each row and store the original heights
        double[] originalHeights = new double[lastRow - firstRow + 1];
        for (int i = firstRow; i <= lastRow; i++)
        {
            double height = 15 + i;               // arbitrary unique height
            cells.SetRowHeight(i, height);        // apply height
            originalHeights[i - firstRow] = cells.GetRowHeight(i); // store original
        }

        // Group the rows (hide them to simulate typical grouping)
        cells.GroupRows(firstRow, lastRow, true);

        // Ungroup the rows (remove the outer group information)
        cells.UngroupRows(firstRow, lastRow);

        // Verify that each row's height has reverted to its original value
        bool allMatch = true;
        for (int i = firstRow; i <= lastRow; i++)
        {
            double currentHeight = cells.GetRowHeight(i);
            double originalHeight = originalHeights[i - firstRow];
            if (Math.Abs(currentHeight - originalHeight) > 0.001)
            {
                allMatch = false;
                Console.WriteLine($"Row {i} height mismatch. Original: {originalHeight}, Current: {currentHeight}");
            }
        }

        Console.WriteLine(allMatch
            ? "All row heights restored correctly after ungrouping."
            : "Row height verification failed.");

        // Save the workbook to demonstrate the final state
        workbook.Save("UngroupRowsResult.xlsx");
    }
}
