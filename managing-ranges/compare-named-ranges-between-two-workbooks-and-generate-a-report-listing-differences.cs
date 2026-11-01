// Title: Generate a named‑range differences report by comparing two Excel workbooks with Aspose.Cells for .NET
// Common Searches: c# compare named ranges in two Excel files using Aspose.Cells | how to create a report of missing or mismatched named ranges with Aspose.Cells .NET | Aspose.Cells example for detecting dimension differences in named ranges across workbooks | C# code to list cell‑by‑cell value differences for matching named ranges in Excel | generate Excel report of named range discrepancies using Aspose.Cells library
// Tags: Aspose.Cells named range diff analysis | detect absent named ranges in Excel C# | named range size mismatch detection | cell‑level value comparison for named ranges | auto‑generate differences workbook with Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The sample loads two Excel workbooks with Aspose.Cells, builds case‑insensitive dictionaries of their named ranges, and iterates the combined set to identify missing ranges, dimension mismatches, or cell‑by‑cell value differences. Each discrepancy is recorded in a new workbook that lists the range name, cell address, values from both source files, and a comment, then auto‑fits columns and saves the report.
class NamedRangeComparer
{
    static void Main()
    {
        try
        {
            // Paths to the source workbooks and the report workbook
            string workbookPath1 = "Workbook1.xlsx";
            string workbookPath2 = "Workbook2.xlsx";
            string reportPath = "NamedRangeDifferencesReport.xlsx";

            // Verify source files exist to avoid FileNotFoundException
            if (!File.Exists(workbookPath1))
                throw new FileNotFoundException($"Source workbook not found: {workbookPath1}");
            if (!File.Exists(workbookPath2))
                throw new FileNotFoundException($"Source workbook not found: {workbookPath2}");

            // Load the two workbooks
            Workbook wb1 = new Workbook(workbookPath1);
            Workbook wb2 = new Workbook(workbookPath2);

            // Create a new workbook for the report
            Workbook reportWb = new Workbook();
            Worksheet reportSheet = reportWb.Worksheets[0];
            reportSheet.Name = "Differences";

            // Write header row
            reportSheet.Cells[0, 0].PutValue("Named Range");
            reportSheet.Cells[0, 1].PutValue("Cell Address");
            reportSheet.Cells[0, 2].PutValue("Value in Workbook 1");
            reportSheet.Cells[0, 3].PutValue("Value in Workbook 2");
            reportSheet.Cells[0, 4].PutValue("Comment");

            int reportRow = 1;

            // Build dictionaries of named ranges for quick lookup
            var ranges1 = new Dictionary<string, AsposeRange>(StringComparer.OrdinalIgnoreCase);
            foreach (Name name in wb1.Worksheets.Names)
            {
                AsposeRange rng = name.GetRange();
                if (rng != null)
                    ranges1[name.Text] = rng;
            }

            var ranges2 = new Dictionary<string, AsposeRange>(StringComparer.OrdinalIgnoreCase);
            foreach (Name name in wb2.Worksheets.Names)
            {
                AsposeRange rng = name.GetRange();
                if (rng != null)
                    ranges2[name.Text] = rng;
            }

            // Union of all named range keys
            var allNames = new HashSet<string>(ranges1.Keys, StringComparer.OrdinalIgnoreCase);
            allNames.UnionWith(ranges2.Keys);

            foreach (string rangeName in allNames)
            {
                bool inFirst = ranges1.TryGetValue(rangeName, out AsposeRange range1);
                bool inSecond = ranges2.TryGetValue(rangeName, out AsposeRange range2);

                if (!inFirst)
                {
                    // Named range exists only in workbook 2
                    reportSheet.Cells[reportRow, 0].PutValue(rangeName);
                    reportSheet.Cells[reportRow, 4].PutValue("Missing in Workbook 1");
                    reportRow++;
                    continue;
                }

                if (!inSecond)
                {
                    // Named range exists only in workbook 1
                    reportSheet.Cells[reportRow, 0].PutValue(rangeName);
                    reportSheet.Cells[reportRow, 4].PutValue("Missing in Workbook 2");
                    reportRow++;
                    continue;
                }

                // Both workbooks contain the named range – compare dimensions first
                if (range1.RowCount != range2.RowCount || range1.ColumnCount != range2.ColumnCount)
                {
                    reportSheet.Cells[reportRow, 0].PutValue(rangeName);
                    reportSheet.Cells[reportRow, 4].PutValue("Dimension mismatch");
                    reportRow++;
                    continue;
                }

                // Compare each cell inside the range
                for (int i = 0; i < range1.RowCount; i++)
                {
                    for (int j = 0; j < range1.ColumnCount; j++)
                    {
                        Cell cell1 = range1[i, j];
                        Cell cell2 = range2[i, j];

                        string val1 = cell1.Value?.ToString() ?? string.Empty;
                        string val2 = cell2.Value?.ToString() ?? string.Empty;

                        if (!val1.Equals(val2, StringComparison.Ordinal))
                        {
                            // Record the difference
                            reportSheet.Cells[reportRow, 0].PutValue(rangeName);
                            string address = $"{cell1.Worksheet.Name}!{cell1.Name}";
                            reportSheet.Cells[reportRow, 1].PutValue(address);
                            reportSheet.Cells[reportRow, 2].PutValue(val1);
                            reportSheet.Cells[reportRow, 3].PutValue(val2);
                            reportSheet.Cells[reportRow, 4].PutValue("Value mismatch");
                            reportRow++;
                        }
                    }
                }
            }

            // Auto‑fit columns for readability
            reportSheet.AutoFitColumns();

            // Ensure the directory for the report exists
            string reportDir = Path.GetDirectoryName(reportPath);
            if (!string.IsNullOrEmpty(reportDir) && !Directory.Exists(reportDir))
                Directory.CreateDirectory(reportDir);

            // Save the report workbook
            reportWb.Save(reportPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
