// Title: Create a cell‑by‑cell diff report for two Excel worksheets with Aspose.Cells in C#
// AI Prompts: Write C# code that uses Aspose.Cells enumerators to walk through every cell of two worksheets, compare their values, and collect mismatched addresses. | Add logic that records cells present only in one worksheet as NULL in the diff output. | Encapsulate the comparison into a reusable method that returns a list of (address, valueFromFirst, valueFromSecond) tuples and writes the list to a new worksheet named DiffReport.
// Common Searches: aspocells c# compare two worksheets cell by cell | how to generate an Excel diff report using Aspose.Cells | enumerate worksheet cells with Aspose.Cells and find differences | c# create diff worksheet that shows missing or changed cells between workbooks | Aspose.Cells diff two workbooks and export mismatched cells to new sheet
// Tags: cell enumeration with Aspose.Cells | worksheet value comparison Aspose.Cells | Excel diff report generation C# | detect missing cells between workbooks Aspose.Cells | write diff results to new worksheet Aspose.Cells | compare two .xlsx files using Aspose.Cells

using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDiffReport
{
    // The program ensures two sample .xlsx files exist, loads them with Aspose.Cells, enumerates every cell of the first worksheet in each workbook, compares values address‑by‑address, captures mismatches and cells that exist only in one sheet, writes the differences (address, value from workbook1, value from workbook2) to a newly added "DiffReport" worksheet, and saves the result as DiffReport.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Ensure input workbooks exist; create sample files if they are missing.
                string file1 = "Workbook1.xlsx";
                string file2 = "Workbook2.xlsx";
                EnsureWorkbook(file1, new Dictionary<string, object>
                {
                    { "A1", "Name" },
                    { "B1", "Age" },
                    { "A2", "Alice" },
                    { "B2", 30 }
                });
                EnsureWorkbook(file2, new Dictionary<string, object>
                {
                    { "A1", "Name" },
                    { "B1", "Age" },
                    { "A2", "Alice" },
                    { "B2", 31 }, // Different value to demonstrate diff
                    { "A3", "Bob" } // Extra row in Workbook2
                });

                // Load the two workbooks to be compared
                Workbook wb1 = new Workbook(file1);
                Workbook wb2 = new Workbook(file2);

                // Assume we compare the first worksheet of each workbook
                Worksheet ws1 = wb1.Worksheets[0];
                Worksheet ws2 = wb2.Worksheets[0];

                // Store all cells from the first worksheet in a dictionary (address -> value)
                var sheet1Cells = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
                IEnumerator enum1 = ws1.Cells.GetEnumerator();
                while (enum1.MoveNext())
                {
                    Cell cell = (Cell)enum1.Current;
                    sheet1Cells[cell.Name] = cell.Value;
                }

                // Prepare a list to hold differences
                var diffs = new List<(string Address, object Value1, object Value2)>();
                // Keep track of addresses that have been processed from sheet1
                var processed = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                // Enumerate cells of the second worksheet and compare with sheet1
                IEnumerator enum2 = ws2.Cells.GetEnumerator();
                while (enum2.MoveNext())
                {
                    Cell cell = (Cell)enum2.Current;
                    string address = cell.Name;
                    object value2 = cell.Value;

                    if (sheet1Cells.TryGetValue(address, out object value1))
                    {
                        // Mark as processed
                        processed.Add(address);

                        // Compare values (handle nulls)
                        bool equal = (value1 == null && value2 == null) ||
                                     (value1 != null && value1.Equals(value2));

                        if (!equal)
                        {
                            diffs.Add((address, value1, value2));
                        }
                    }
                    else
                    {
                        // Cell exists only in sheet2
                        diffs.Add((address, null, value2));
                    }
                }

                // Any cells that exist only in sheet1 (not visited in sheet2)
                foreach (var kvp in sheet1Cells)
                {
                    if (!processed.Contains(kvp.Key))
                    {
                        diffs.Add((kvp.Key, kvp.Value, null));
                    }
                }

                // Create a new worksheet to hold the diff report
                Worksheet diffSheet = wb1.Worksheets.Add("DiffReport");
                // Write header
                diffSheet.Cells["A1"].PutValue("Cell Address");
                diffSheet.Cells["B1"].PutValue("Workbook1 Value");
                diffSheet.Cells["C1"].PutValue("Workbook2 Value");

                // Populate diff rows
                int rowIndex = 1; // zero‑based index; start after header
                foreach (var diff in diffs)
                {
                    diffSheet.Cells[rowIndex, 0].PutValue(diff.Address);
                    diffSheet.Cells[rowIndex, 1].PutValue(diff.Value1?.ToString() ?? "NULL");
                    diffSheet.Cells[rowIndex, 2].PutValue(diff.Value2?.ToString() ?? "NULL");
                    rowIndex++;
                }

                // Save the workbook containing the diff report
                string reportFile = "DiffReport.xlsx";
                wb1.Save(reportFile);
                Console.WriteLine($"Diff report saved to '{reportFile}'.");
            }
            catch (CellsException ex)
            {
                Console.WriteLine($"Aspose.Cells error: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        private static void EnsureWorkbook(string path, Dictionary<string, object> cellData)
        {
            if (File.Exists(path))
                return;

            var wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            foreach (var kvp in cellData)
            {
                ws.Cells[kvp.Key].PutValue(kvp.Value);
            }

            wb.Save(path);
        }
    }
}
