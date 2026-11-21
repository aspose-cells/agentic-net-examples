// Title: Generate a differences report for named ranges in two Excel workbooks using Aspose.Cells for .NET
// Common Searches: C# compare named ranges in two Excel files with Aspose.Cells and export differences to a new workbook | how to detect missing named ranges between two workbooks using Aspose.Cells .NET | Aspose.Cells compare range dimensions and cell values across Excel workbooks in C# | create a named range discrepancy report with Aspose.Cells for .NET | C# code to list named range mismatches and generate Excel report using Aspose.Cells
// Tags: Aspose.Cells named range comparison | Excel workbook difference reporting .NET | detect missing named ranges Aspose.Cells | range dimension mismatch detection C# | cell value mismatch reporting Aspose.Cells | generate named range differences report

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace NamedRangeComparison
{
    // The sample loads two Excel workbooks, extracts all defined named ranges into dictionaries, and then iterates over the combined set of range names. It identifies missing ranges, dimension mismatches, and cell‑by‑cell value differences, records each discrepancy in a DTO, and writes the results to a new workbook (NamedRangeDifferencesReport.xlsx) with columns for the named range, cell address, values from each source workbook, and the type of difference.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Paths to the source workbooks
                string workbookPath1 = "Workbook1.xlsx";
                string workbookPath2 = "Workbook2.xlsx";

                // Verify input files exist
                if (!File.Exists(workbookPath1))
                {
                    Console.WriteLine($"Error: File not found - {workbookPath1}");
                    return;
                }
                if (!File.Exists(workbookPath2))
                {
                    Console.WriteLine($"Error: File not found - {workbookPath2}");
                    return;
                }

                // Load the two workbooks
                Workbook wb1 = new Workbook(workbookPath1);
                Workbook wb2 = new Workbook(workbookPath2);

                // Dictionaries to hold named ranges for quick lookup
                Dictionary<string, AsposeRange> ranges1 = GetNamedRanges(wb1);
                Dictionary<string, AsposeRange> ranges2 = GetNamedRanges(wb2);

                // List to store difference records
                List<DifferenceRecord> differences = new List<DifferenceRecord>();

                // Union of all named range names
                HashSet<string> allNames = new HashSet<string>(ranges1.Keys, StringComparer.OrdinalIgnoreCase);
                allNames.UnionWith(ranges2.Keys);

                foreach (string name in allNames)
                {
                    bool hasInWb1 = ranges1.TryGetValue(name, out AsposeRange range1);
                    bool hasInWb2 = ranges2.TryGetValue(name, out AsposeRange range2);

                    if (!hasInWb1)
                    {
                        differences.Add(new DifferenceRecord
                        {
                            NamedRange = name,
                            CellAddress = "",
                            ValueInWorkbook1 = "Missing",
                            ValueInWorkbook2 = "Exists",
                            DifferenceType = "Named range missing in Workbook1"
                        });
                        continue;
                    }

                    if (!hasInWb2)
                    {
                        differences.Add(new DifferenceRecord
                        {
                            NamedRange = name,
                            CellAddress = "",
                            ValueInWorkbook1 = "Exists",
                            ValueInWorkbook2 = "Missing",
                            DifferenceType = "Named range missing in Workbook2"
                        });
                        continue;
                    }

                    // Both ranges exist – compare dimensions
                    if (range1.RowCount != range2.RowCount || range1.ColumnCount != range2.ColumnCount)
                    {
                        differences.Add(new DifferenceRecord
                        {
                            NamedRange = name,
                            CellAddress = "",
                            ValueInWorkbook1 = $"{range1.RowCount}x{range1.ColumnCount}",
                            ValueInWorkbook2 = $"{range2.RowCount}x{range2.ColumnCount}",
                            DifferenceType = "Dimension mismatch"
                        });
                        continue;
                    }

                    // Compare cell by cell
                    for (int i = 0; i < range1.RowCount; i++)
                    {
                        for (int j = 0; j < range1.ColumnCount; j++)
                        {
                            Cell cell1 = range1[i, j];
                            Cell cell2 = range2[i, j];

                            // Use string representation for comparison to handle different data types uniformly
                            string val1 = cell1?.StringValue ?? "";
                            string val2 = cell2?.StringValue ?? "";

                            if (!val1.Equals(val2))
                            {
                                // Build absolute address (e.g., Sheet1!A1)
                                string address = $"{range1.Worksheet?.Name ?? "Unknown"}!{cell1?.Name ?? ""}";

                                differences.Add(new DifferenceRecord
                                {
                                    NamedRange = name,
                                    CellAddress = address,
                                    ValueInWorkbook1 = val1,
                                    ValueInWorkbook2 = val2,
                                    DifferenceType = "Value mismatch"
                                });
                            }
                        }
                    }
                }

                // Create a report workbook
                Workbook reportWb = new Workbook();
                Worksheet reportWs = reportWb.Worksheets[0];
                reportWs.Name = "Differences";

                // Write header
                reportWs.Cells[0, 0].PutValue("Named Range");
                reportWs.Cells[0, 1].PutValue("Cell Address");
                reportWs.Cells[0, 2].PutValue("Value in Workbook1");
                reportWs.Cells[0, 3].PutValue("Value in Workbook2");
                reportWs.Cells[0, 4].PutValue("Difference Type");

                // Populate rows
                for (int row = 0; row < differences.Count; row++)
                {
                    DifferenceRecord dr = differences[row];
                    int r = row + 1; // offset for header
                    reportWs.Cells[r, 0].PutValue(dr.NamedRange);
                    reportWs.Cells[r, 1].PutValue(dr.CellAddress);
                    reportWs.Cells[r, 2].PutValue(dr.ValueInWorkbook1);
                    reportWs.Cells[r, 3].PutValue(dr.ValueInWorkbook2);
                    reportWs.Cells[r, 4].PutValue(dr.DifferenceType);
                }

                // Save the report
                string reportPath = "NamedRangeDifferencesReport.xlsx";
                try
                {
                    reportWb.Save(reportPath);
                    Console.WriteLine($"Report saved to {reportPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to save report: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Helper method to extract all named ranges from a workbook into a dictionary
        private static Dictionary<string, AsposeRange> GetNamedRanges(Workbook wb)
        {
            var dict = new Dictionary<string, AsposeRange>(StringComparer.OrdinalIgnoreCase);
            try
            {
                foreach (Name name in wb.Worksheets.Names)
                {
                    // Retrieve the range by name; may return null if the name does not refer to a range
                    AsposeRange rng = wb.Worksheets.GetRangeByName(name.Text);
                    if (rng != null)
                    {
                        dict[name.Text] = rng;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error while extracting named ranges: {ex.Message}");
            }
            return dict;
        }

        // Simple DTO to hold a difference entry
        private class DifferenceRecord
        {
            public string NamedRange { get; set; } = string.Empty;
            public string CellAddress { get; set; } = string.Empty;
            public string ValueInWorkbook1 { get; set; } = string.Empty;
            public string ValueInWorkbook2 { get; set; } = string.Empty;
            public string DifferenceType { get; set; } = string.Empty;
        }
    }
}
