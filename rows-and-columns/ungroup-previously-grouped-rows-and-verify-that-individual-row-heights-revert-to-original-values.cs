// Title: Ungroup rows and confirm original row heights with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# ungroup rows and retain original row heights | How to verify row height after ungrouping rows in an Excel workbook using Aspose.Cells | C# example to ungroup rows and check each row's height with Debug.Assert | Restore individual row heights after grouping rows with Aspose.Cells for .NET | Aspose.Cells ungroup rows code sample with height validation in C#
// Tags: aspocells ungroup rows preserve heights | row height validation using Debug.Assert C# | excel workbook row height after ungrouping | c# aspocells row grouping operations | set distinct row heights Aspose.Cells

using System;
using System.Diagnostics;
using Aspose.Cells;

// The example creates a workbook, assigns unique heights to rows 1‑5, groups rows 2‑4, then ungroups them and uses Debug.Assert to ensure each row’s height returns to its original value before saving the file.
class UngroupRowsExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define rows to work with (rows 1 to 5, zero‑based indices 0‑4)
            int startRow = 0;
            int rowCount = 5;

            // Set distinct heights for each row and store the original heights
            double[] originalHeights = new double[rowCount];
            for (int i = 0; i < rowCount; i++)
            {
                // Assign a unique height (e.g., 15 + i * 2 points)
                double height = 15 + i * 2;
                sheet.Cells.Rows[startRow + i].Height = height;
                originalHeights[i] = height;
            }

            // Group rows 2‑4 (zero‑based indices 1‑3)
            sheet.Cells.GroupRows(1, 3); // groups rows 2,3,4

            // Ungroup the same rows
            sheet.Cells.UngroupRows(1, 3); // ungroup rows 2,3,4

            // Verify that each row's height has reverted to its original value
            for (int i = 0; i < rowCount; i++)
            {
                double currentHeight = sheet.Cells.Rows[startRow + i].Height;
                double expectedHeight = originalHeights[i];

                // Use Debug.Assert for verification; in production you might throw an exception
                Debug.Assert(Math.Abs(currentHeight - expectedHeight) < 0.001,
                    $"Row {startRow + i + 1} height mismatch. Expected: {expectedHeight}, Actual: {currentHeight}");
            }

            // Save the workbook
            workbook.Save("UngroupRowsResult.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
