// Title: Generate a cell‑by‑cell difference report between two Excel worksheets using Aspose.Cells enumerators in C#
// Common Searches: C# compare two Excel worksheets cell by cell with Aspose.Cells and export mismatches to a new workbook | How to create a diff report of missing and changed cells between two .xlsx files using Aspose.Cells enumerator | Aspose.Cells C# enumerate worksheet cells to find value differences and generate diff report | Generate Excel diff report for two workbooks using Aspose.Cells IEnumerator in .NET | C# Aspose.Cells compare first worksheets of two workbooks and list cell address mismatches
// Tags: Aspose.Cells cell enumerator comparison | Excel worksheet diff report generation | C# detect missing cells in workbooks | compare worksheet values with Aspose.Cells | enumerator based Excel diff in .NET

using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example loads two .xlsx files, uses IEnumerator to walk through the cells of the first worksheet in each workbook, compares cell addresses and values, records mismatches and cells that exist only in one sheet, and writes the differences to a new workbook (DiffReport.xlsx) with columns for cell reference, value from Workbook1, and value from Workbook2.
class WorksheetDiff
{
    static void Main()
    {
        try
        {
            const string file1 = "Workbook1.xlsx";
            const string file2 = "Workbook2.xlsx";
            const string diffFile = "DiffReport.xlsx";

            // Verify input files exist to avoid FileNotFoundException
            if (!File.Exists(file1))
                throw new FileNotFoundException($"Input file not found: {file1}");
            if (!File.Exists(file2))
                throw new FileNotFoundException($"Input file not found: {file2}");

            // Load the two workbooks to compare
            Workbook wb1 = new Workbook(file1);
            Workbook wb2 = new Workbook(file2);

            // Assume we compare the first worksheet of each workbook
            Worksheet ws1 = wb1.Worksheets[0];
            Worksheet ws2 = wb2.Worksheets[0];

            // Create a new workbook that will hold the diff report
            Workbook diffWb = new Workbook();
            Worksheet diffWs = diffWb.Worksheets[0];
            diffWs.Name = "DiffReport";

            // Write header row for the report
            diffWs.Cells["A1"].PutValue("Cell");
            diffWs.Cells["B1"].PutValue("Value in Workbook1");
            diffWs.Cells["C1"].PutValue("Value in Workbook2");

            int diffRow = 1; // zero‑based index (row 2 in Excel)

            // Enumerators to iterate through cells of each worksheet
            IEnumerator enum1 = ws1.Cells.GetEnumerator();
            IEnumerator enum2 = ws2.Cells.GetEnumerator();

            bool hasCell1 = enum1.MoveNext();
            bool hasCell2 = enum2.MoveNext();

            // Walk through both enumerators simultaneously
            while (hasCell1 && hasCell2)
            {
                Cell cell1 = (Cell)enum1.Current;
                Cell cell2 = (Cell)enum2.Current;

                string addr1 = cell1.Name; // e.g. "A1"
                string addr2 = cell2.Name;

                int cmp = string.Compare(addr1, addr2, StringComparison.Ordinal);
                if (cmp == 0)
                {
                    // Same address – compare values
                    object val1 = cell1.Value;
                    object val2 = cell2.Value;

                    if (!Object.Equals(val1, val2))
                    {
                        diffWs.Cells[diffRow, 0].PutValue(addr1);
                        diffWs.Cells[diffRow, 1].PutValue(val1);
                        diffWs.Cells[diffRow, 2].PutValue(val2);
                        diffRow++;
                    }

                    hasCell1 = enum1.MoveNext();
                    hasCell2 = enum2.MoveNext();
                }
                else if (cmp < 0)
                {
                    // Cell exists only in ws1
                    diffWs.Cells[diffRow, 0].PutValue(addr1);
                    diffWs.Cells[diffRow, 1].PutValue(cell1.Value);
                    diffWs.Cells[diffRow, 2].PutValue("<<Missing>>");
                    diffRow++;
                    hasCell1 = enum1.MoveNext();
                }
                else // cmp > 0
                {
                    // Cell exists only in ws2
                    diffWs.Cells[diffRow, 0].PutValue(addr2);
                    diffWs.Cells[diffRow, 1].PutValue("<<Missing>>");
                    diffWs.Cells[diffRow, 2].PutValue(cell2.Value);
                    diffRow++;
                    hasCell2 = enum2.MoveNext();
                }
            }

            // Remaining cells in ws1 (no counterpart in ws2)
            while (hasCell1)
            {
                Cell cell1 = (Cell)enum1.Current;
                diffWs.Cells[diffRow, 0].PutValue(cell1.Name);
                diffWs.Cells[diffRow, 1].PutValue(cell1.Value);
                diffWs.Cells[diffRow, 2].PutValue("<<Missing>>");
                diffRow++;
                hasCell1 = enum1.MoveNext();
            }

            // Remaining cells in ws2 (no counterpart in ws1)
            while (hasCell2)
            {
                Cell cell2 = (Cell)enum2.Current;
                diffWs.Cells[diffRow, 0].PutValue(cell2.Name);
                diffWs.Cells[diffRow, 1].PutValue("<<Missing>>");
                diffWs.Cells[diffRow, 2].PutValue(cell2.Value);
                diffRow++;
                hasCell2 = enum2.MoveNext();
            }

            // Save the diff report workbook
            diffWb.Save(diffFile);
            Console.WriteLine($"Diff report saved to '{diffFile}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
