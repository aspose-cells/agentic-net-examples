// Title: C# – Ungroup Rows and Restore Original Row Heights with Aspose.Cells for .NET
// Description: Demonstrates how to set custom heights for rows, group them, ungroup them, and verify that each row returns to its original height using Aspose.Cells. The workbook is saved after the validation.
// Keywords: Aspose.Cells ungroup rows C# | restore row height after grouping | verify row height Aspose.Cells | GroupRows UngroupRows .NET | row height preservation Excel | Aspose.Cells Workbook example | C# Excel row formatting | Aspose.Cells API row height | UngroupRowsDemo GitHub
// Common Searches: how to ungroup rows with Aspose.Cells .NET | check row height after ungrouping in C# | Aspose.Cells restore original row height | group and ungroup rows Excel using Aspose.Cells | C# code to verify row heights after ungrouping
// Developer Intent: Remove grouping from a range of rows and confirm that their heights revert to the values set before grouping.
// Use Cases: Create a worksheet, assign distinct heights to rows, temporarily group them for printing, then ungroup while ensuring formatting is unchanged. | Automate a data‑processing routine that collapses rows for batch operations and validates that ungrouping does not modify row height settings. | Generate Excel reports where rows are collapsed for layout purposes and later expanded without losing custom row height definitions.
// AI Prompts: Write C# code using Aspose.Cells to set custom heights for rows 2‑4, group them, ungroup them, and assert that each row height matches the pre‑grouping value. | Provide a reusable method that accepts a worksheet and a row range, groups the rows, ungroups them, and returns a boolean indicating whether all row heights were restored. | Explain the behavior of Aspose.Cells when rows are grouped and later ungrouped, focusing on how original row heights are preserved.

using System;
using Aspose.Cells;

namespace AsposeCellsUngroupRowsDemo
{
    // Demonstrates how to set custom heights for rows, group them, ungroup them, and verify that each row returns to its original height using Aspose.Cells. The workbook is saved after the validation.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate rows 1 to 5 with sample data
            for (int i = 0; i < 5; i++)
            {
                cells[i, 0].PutValue($"Row {i + 1}");
            }

            // Set distinct heights for rows 2, 3, and 4 (zero‑based indices 1‑3)
            double[] originalHeights = new double[3];
            for (int i = 1; i <= 3; i++)
            {
                double height = 15 + i * 2;               // example heights: 17, 19, 21
                cells.SetRowHeight(i, height);
                originalHeights[i - 1] = height;          // store for later verification
            }

            // Group rows 2‑4 (indices 1‑3) and hide them
            cells.GroupRows(1, 3, true);

            // Ungroup the same rows, removing all grouping levels
            cells.UngroupRows(1, 3, true);

            // Verify that each row's height has reverted to the original value
            bool allMatch = true;
            for (int i = 1; i <= 3; i++)
            {
                double currentHeight = cells.GetRowHeight(i);
                double expectedHeight = originalHeights[i - 1];
                Console.WriteLine($"Row {i + 1}: Expected Height = {expectedHeight}, Current Height = {currentHeight}");
                if (Math.Abs(currentHeight - expectedHeight) > 0.001)
                {
                    allMatch = false;
                }
            }

            Console.WriteLine(allMatch
                ? "All row heights have reverted to their original values after ungrouping."
                : "Row heights differ after ungrouping.");

            // Save the workbook to demonstrate the final state
            workbook.Save("UngroupRowsResult.xlsx");
        }
    }
}
