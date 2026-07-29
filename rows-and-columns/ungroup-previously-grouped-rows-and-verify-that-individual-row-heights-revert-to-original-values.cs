// Title: C# Example: Ungroup Rows and Restore Original Row Heights with Aspose.Cells for .NET
// Description: Demonstrates how to set custom heights for rows, group a range, ungroup it, and programmatically verify that each row's height returns to the value assigned before grouping, then saves the workbook.
// Keywords: Aspose.Cells ungroup rows C# | restore row height after grouping | verify row height Aspose.Cells .NET | group and ungroup rows example | Excel row height validation C#
// Common Searches: how to ungroup rows with Aspose.Cells and keep original height | check row height after ungrouping in C# | Aspose.Cells restore row heights after grouping rows | C# code to verify row heights after ungrouping Excel rows | Aspose.Cells example ungroup rows and validate heights
// Developer Intent: The developer needs to ungroup previously grouped rows and confirm that each row's height matches the original value set before grouping.
// Use Cases: Create a spreadsheet, assign distinct heights to rows, temporarily collapse a subset for printing, then expand and ensure heights are unchanged. | Automate quality checks in an Excel processing pipeline that groups and ungroups rows while preserving layout integrity. | Generate reports where rows are grouped for summary views and later ungrouped, requiring verification that visual formatting remains consistent.
// AI Prompts: Write C# code using Aspose.Cells to set row heights, group rows 2‑4, ungroup them, and assert that the heights are restored to the original values. | Provide a reusable method that takes a worksheet and a row range, groups the rows, ungroups them, and returns true if the original heights are preserved. | Explain Aspose.Cells' handling of row height metadata during grouping and ungrouping, and suggest best practices for testing this behavior in automated scripts.

using System;
using Aspose.Cells;

// Demonstrates how to set custom heights for rows, group a range, ungroup it, and programmatically verify that each row's height returns to the value assigned before grouping, then saves the workbook.
class UngroupRowsDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate rows 1‑5 with sample data
        for (int i = 0; i < 5; i++)
        {
            cells[i, 0].PutValue($"Row {i + 1}");
        }

        // Assign distinct heights to rows 1‑5 and store the original values
        double[] originalHeights = new double[5];
        for (int i = 0; i < 5; i++)
        {
            double height = 15 + i * 2; // Example heights: 15, 17, 19, 21, 23
            cells.SetRowHeight(i, height);
            originalHeights[i] = height;
        }

        // Group rows 2‑4 (zero‑based indices 1‑3) and hide them
        cells.GroupRows(1, 3, true);

        // Ungroup the same rows, removing all grouping levels
        cells.UngroupRows(1, 3, true);

        // Verify that each row's height matches the original value
        bool heightsMatch = true;
        for (int i = 0; i < 5; i++)
        {
            double currentHeight = cells.GetRowHeight(i);
            if (Math.Abs(currentHeight - originalHeights[i]) > 0.001)
            {
                heightsMatch = false;
                Console.WriteLine($"Row {i + 1} height mismatch. Original: {originalHeights[i]}, Current: {currentHeight}");
            }
        }

        Console.WriteLine(heightsMatch
            ? "All row heights restored to original values."
            : "Row heights differ after ungrouping.");

        // Save the workbook
        workbook.Save("UngroupRowsResult.xlsx");
    }
}
