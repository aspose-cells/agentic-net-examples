// Title: C# Aspose.Cells: Cell‑by‑Cell Worksheet Comparison with Diff Report
// Description: Loads two Excel workbooks (creates empty workbooks if files are missing), enumerates every populated cell in the first worksheet, compares each cell's value with the matching cell in the second worksheet, records mismatches and cells that exist only in one sheet, and saves the results to a DiffReport.xlsx workbook.
// Keywords: Aspose.Cells compare worksheets | cell level diff C# | enumerator cell comparison | Excel diff report .NET | missing file handling Aspose | worksheet mismatch detection
// Common Searches: compare two Excel worksheets cell by cell Aspose.Cells | generate diff report for Excel files C# | enumerate cells Aspose.Cells to find differences | create diff workbook when one file is missing
// Developer Intent: Produce a workbook that lists every cell where two worksheets differ, including cells present in only one of the files.
// Use Cases: Validate data consistency between two versions of a financial statement before release. | Verify that a data migration copied all cell values correctly by comparing source and target spreadsheets. | Audit configuration changes after an automated update by generating a mismatch report.
// AI Prompts: Write C# code with Aspose.Cells to compare two worksheets cell by cell and output mismatched cells to a new workbook. | Refactor the diff report program to use foreach loops instead of IEnumerator while keeping the same functionality. | Explain how to extend the diff report to capture differences in cell formatting such as font color or background.

using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDiffReport
{
    // Loads two Excel workbooks (creates empty workbooks if files are missing), enumerates every populated cell in the first worksheet, compares each cell's value with the matching cell in the second worksheet, records mismatches and cells that exist only in one sheet, and saves the results to a DiffReport.xlsx workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                // Paths to the source workbooks
                string sourcePath1 = "File1.xlsx";
                string sourcePath2 = "File2.xlsx";

                // Load the two workbooks, create empty ones if files are missing
                Workbook wb1 = File.Exists(sourcePath1) ? new Workbook(sourcePath1) : CreateEmptyWorkbook(sourcePath1);
                Workbook wb2 = File.Exists(sourcePath2) ? new Workbook(sourcePath2) : CreateEmptyWorkbook(sourcePath2);

                // Access the first worksheet of each workbook
                Worksheet ws1 = wb1.Worksheets[0];
                Worksheet ws2 = wb2.Worksheets[0];

                // Create a new workbook that will hold the diff report
                Workbook diffWb = new Workbook();
                Worksheet diffSheet = diffWb.Worksheets[0];
                diffSheet.Name = "DiffReport";

                // Write header row
                diffSheet.Cells[0, 0].PutValue("Cell");
                diffSheet.Cells[0, 1].PutValue("Value in File1");
                diffSheet.Cells[0, 2].PutValue("Value in File2");

                int diffRowIndex = 1; // start after header

                // ---------- Compare cells that exist in the first worksheet ----------
                IEnumerator enum1 = ws1.Cells.GetEnumerator(); // get enumerator for cells collection
                while (enum1.MoveNext())
                {
                    Cell cell1 = (Cell)enum1.Current;
                    // Get the counterpart cell from the second worksheet using same row/column
                    Cell cell2 = ws2.Cells[cell1.Row, cell1.Column];

                    // Convert values to string for comparison (handle nulls)
                    string val1 = cell1.Value?.ToString() ?? string.Empty;
                    string val2 = cell2?.Value?.ToString() ?? string.Empty;

                    // If values differ, record the mismatch
                    if (!val1.Equals(val2))
                    {
                        diffSheet.Cells[diffRowIndex, 0].PutValue(cell1.Name);
                        diffSheet.Cells[diffRowIndex, 1].PutValue(val1);
                        diffSheet.Cells[diffRowIndex, 2].PutValue(val2);
                        diffRowIndex++;
                    }
                }

                // ---------- Find cells that exist only in the second worksheet ----------
                // Use a hash set to remember cells already processed from the first sheet
                var processed = new HashSet<string>();
                IEnumerator enumProcessed = ws1.Cells.GetEnumerator();
                while (enumProcessed.MoveNext())
                {
                    Cell c = (Cell)enumProcessed.Current;
                    processed.Add(c.Name);
                }

                IEnumerator enum2 = ws2.Cells.GetEnumerator();
                while (enum2.MoveNext())
                {
                    Cell cell2 = (Cell)enum2.Current;
                    // Skip cells already compared
                    if (processed.Contains(cell2.Name))
                        continue;

                    // Cell exists only in second worksheet
                    string val2 = cell2.Value?.ToString() ?? string.Empty;
                    diffSheet.Cells[diffRowIndex, 0].PutValue(cell2.Name);
                    diffSheet.Cells[diffRowIndex, 1].PutValue(string.Empty); // no value in first file
                    diffSheet.Cells[diffRowIndex, 2].PutValue(val2);
                    diffRowIndex++;
                }

                // Save the diff report
                diffWb.Save("DiffReport.xlsx");
                Console.WriteLine("Diff report generated successfully: DiffReport.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Helper method to create an empty workbook and optionally inform the user
        private static Workbook CreateEmptyWorkbook(string missingFilePath)
        {
            Console.WriteLine($"File not found: {missingFilePath}. An empty workbook will be used.");
            return new Workbook();
        }
    }
}
