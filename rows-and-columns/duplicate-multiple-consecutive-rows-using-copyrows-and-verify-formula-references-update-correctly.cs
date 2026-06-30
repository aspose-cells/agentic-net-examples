using System;
using Aspose.Cells;

namespace AsposeCellsRowDuplicationDemo
{
    // Author: Aspose.Cells .NET example – duplicate rows and verify formula updates
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data in column A
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].PutValue(30);
            sheet.Cells["A4"].PutValue(40);

            // Add formulas in column B that reference the same row in column A
            sheet.Cells["B1"].Formula = "=A1*2";
            sheet.Cells["B2"].Formula = "=A2*2";
            sheet.Cells["B3"].Formula = "=A3*2";
            sheet.Cells["B4"].Formula = "=A4*2";

            // Insert two blank rows after the second row (index 2, zero‑based)
            // updateReference = true ensures existing formulas adjust to the insertion
            sheet.Cells.InsertRows(2, 2, true);

            // Copy rows 0 and 1 (first two rows) to the newly inserted rows starting at index 2
            // This duplicates the rows while preserving relative formula references
            sheet.Cells.CopyRows(sheet.Cells, 0, 2, 2);

            // Verify that formulas in the copied rows have been updated correctly
            Console.WriteLine("Original formulas:");
            Console.WriteLine($"B1: {sheet.Cells["B1"].Formula}");
            Console.WriteLine($"B2: {sheet.Cells["B2"].Formula}");

            Console.WriteLine("\nCopied formulas after duplication:");
            Console.WriteLine($"B3: {sheet.Cells["B3"].Formula}");
            Console.WriteLine($"B4: {sheet.Cells["B4"].Formula}");

            // Expected output:
            // B3 formula should be "=A3*2"
            // B4 formula should be "=A4*2"

            // Save the workbook to verify visually if needed
            workbook.Save("RowDuplicationResult.xlsx");
        }
    }
}