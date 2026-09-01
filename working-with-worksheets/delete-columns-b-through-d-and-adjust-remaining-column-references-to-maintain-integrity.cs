// Title: Delete columns B‑D in an Excel worksheet and automatically adjust formulas using Aspose.Cells for .NET
// AI Prompts: Remove the contiguous column range B through D from a worksheet and have all dependent formulas automatically re‑reference the new cells with Aspose.Cells DeleteOptions.UpdateReference in C#. | Programmatically delete columns B‑D in a .NET workbook while preserving the integrity of SUM or other formulas that originally pointed to those columns.
// Common Searches: aspnet cells delete columns b-d and keep formula references updated | c# remove multiple columns in excel workbook using Aspose.Cells DeleteOptions | how to preserve SUM formula after deleting columns with Aspose.Cells | update cell references automatically after column deletion Aspose.Cells .NET | delete column range B to D in Excel using Aspose.Cells C# example
// Tags: delete columns with reference update Aspose.Cells | Aspose.Cells DeleteOptions.UpdateReference | adjust formulas after column removal .NET | C# remove column range B-D Excel workbook | preserve formula integrity Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsColumnDeletionDemo
{
    // The example creates a workbook, populates columns A‑E, adds a SUM formula in E2 that references B2:D2, then deletes columns B‑D using DeleteOptions.UpdateReference so the formula automatically shifts to the new range (C2), and finally saves the workbook as ColumnDeletionResult.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data in columns A through E
            for (int col = 0; col < 5; col++)
            {
                cells[0, col].PutValue($"Header {(char)('A' + col)}");
                cells[1, col].PutValue($"Data{col + 1}");
            }

            // Add a formula that references columns B through D
            // Example: sum of B2:D2
            cells["E2"].Formula = "=SUM(B2:D2)";

            // Set up DeleteOptions to update references after deletion
            DeleteOptions options = new DeleteOptions
            {
                UpdateReference = true
            };

            // Delete columns B (index 1) through D (index 3) – total of 3 columns
            cells.DeleteColumns(1, 3, options);

            // After deletion, the formula in E2 should automatically adjust to reference the new range
            Console.WriteLine("Formula after column deletion: " + cells["C2"].Formula); // Formerly E2 shifts left

            // Save the modified workbook
            workbook.Save("ColumnDeletionResult.xlsx");
        }
    }
}
