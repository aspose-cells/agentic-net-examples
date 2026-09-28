// Title: Create a C# Aspose.Cells script to compare named ranges in two Excel workbooks and export a differences report
// AI Prompts: Write C# code using Aspose.Cells that loads two .xlsx files, builds case‑insensitive dictionaries of defined names, compares each named range cell‑by‑cell, and saves an Excel workbook reporting missing ranges, size mismatches, and value differences. | Adapt the comparison program to ignore whitespace and apply a numeric tolerance when evaluating cell values, then output the mismatch list as a CSV file instead of an Excel workbook. | Enhance the script to detect formula changes inside matching named ranges and add a "Formula Difference" column to the generated report.
// Common Searches: aspnet compare defined names between two Excel workbooks using Aspose.Cells | c# generate report of named range mismatches in .xlsx files | how to list missing named ranges when comparing two spreadsheets with Aspose.Cells | detect size differences in named ranges across workbooks C# Aspose | compare cell values of matching named ranges in two workbooks programmatically
// Tags: Aspose.Cells compare named ranges | C# generate named range differences report | Excel workbook named range size mismatch detection | Aspose.Cells resolve defined name to range | auto‑fit columns in Aspose.Cells report workbook

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The example validates two input .xlsx files, loads them with Aspose.Cells, creates case‑insensitive dictionaries of defined names, and builds a union of all named range identifiers. It then generates a new workbook, writes a header row, and iterates through each name to report missing ranges, unresolved definitions, size mismatches, and cell‑by‑cell value differences. Columns are auto‑fitted and the report is saved as NamedRangeDifferencesReport.xlsx, with comprehensive exception handling throughout.
class NamedRangeComparer
{
    static void Main()
    {
        try
        {
            // Verify input files exist
            const string file1 = "Workbook1.xlsx";
            const string file2 = "Workbook2.xlsx";

            if (!File.Exists(file1))
                throw new FileNotFoundException($"Input file not found: {file1}");
            if (!File.Exists(file2))
                throw new FileNotFoundException($"Input file not found: {file2}");

            // Load the two workbooks to compare
            Workbook wb1 = new Workbook(file1);
            Workbook wb2 = new Workbook(file2);

            // Create a new workbook for the report
            Workbook reportWb = new Workbook();
            Worksheet reportSheet = reportWb.Worksheets[0];
            reportSheet.Name = "Differences";

            // Write header row
            reportSheet.Cells[0, 0].PutValue("Named Range");
            reportSheet.Cells[0, 1].PutValue("Cell Address");
            reportSheet.Cells[0, 2].PutValue("Workbook1 Value");
            reportSheet.Cells[0, 3].PutValue("Workbook2 Value");
            reportSheet.Cells[0, 4].PutValue("Status");

            int reportRow = 1;

            // Build dictionaries of named ranges for quick lookup
            var names1 = new Dictionary<string, Name>(StringComparer.OrdinalIgnoreCase);
            foreach (Name n in wb1.Worksheets.Names)
                names1[n.Text] = n;

            var names2 = new Dictionary<string, Name>(StringComparer.OrdinalIgnoreCase);
            foreach (Name n in wb2.Worksheets.Names)
                names2[n.Text] = n;

            // Union of all named range names
            var allNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            allNames.UnionWith(names1.Keys);
            allNames.UnionWith(names2.Keys);

            foreach (string rangeName in allNames)
            {
                bool inWb1 = names1.TryGetValue(rangeName, out Name name1);
                bool inWb2 = names2.TryGetValue(rangeName, out Name name2);

                if (!inWb1)
                {
                    // Named range exists only in Workbook2
                    reportSheet.Cells[reportRow, 0].PutValue(rangeName);
                    reportSheet.Cells[reportRow, 1].PutValue("-");
                    reportSheet.Cells[reportRow, 2].PutValue("-");
                    reportSheet.Cells[reportRow, 3].PutValue("-");
                    reportSheet.Cells[reportRow, 4].PutValue("Missing in Workbook1");
                    reportRow++;
                    continue;
                }

                if (!inWb2)
                {
                    // Named range exists only in Workbook1
                    reportSheet.Cells[reportRow, 0].PutValue(rangeName);
                    reportSheet.Cells[reportRow, 1].PutValue("-");
                    reportSheet.Cells[reportRow, 2].PutValue("-");
                    reportSheet.Cells[reportRow, 3].PutValue("-");
                    reportSheet.Cells[reportRow, 4].PutValue("Missing in Workbook2");
                    reportRow++;
                    continue;
                }

                // Both workbooks contain the named range – compare cell by cell
                AsposeRange? range1 = null;
                AsposeRange? range2 = null;

                try
                {
                    range1 = name1.GetRange();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Unable to resolve range '{rangeName}' in Workbook1: {ex.Message}");
                }

                try
                {
                    range2 = name2.GetRange();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Unable to resolve range '{rangeName}' in Workbook2: {ex.Message}");
                }

                if (range1 == null || range2 == null)
                {
                    // Could not resolve one of the ranges – report and skip
                    reportSheet.Cells[reportRow, 0].PutValue(rangeName);
                    reportSheet.Cells[reportRow, 1].PutValue("N/A");
                    reportSheet.Cells[reportRow, 2].PutValue(range1 == null ? "Unresolved" : "Resolved");
                    reportSheet.Cells[reportRow, 3].PutValue(range2 == null ? "Unresolved" : "Resolved");
                    reportSheet.Cells[reportRow, 4].PutValue("Resolution Failure");
                    reportRow++;
                    continue;
                }

                // Ensure the ranges have the same dimensions
                int rows1 = range1.RowCount;
                int cols1 = range1.ColumnCount;
                int rows2 = range2.RowCount;
                int cols2 = range2.ColumnCount;

                if (rows1 != rows2 || cols1 != cols2)
                {
                    // Different size – report as a whole
                    reportSheet.Cells[reportRow, 0].PutValue(rangeName);
                    reportSheet.Cells[reportRow, 1].PutValue("Entire Range");
                    reportSheet.Cells[reportRow, 2].PutValue($"{rows1}x{cols1}");
                    reportSheet.Cells[reportRow, 3].PutValue($"{rows2}x{cols2}");
                    reportSheet.Cells[reportRow, 4].PutValue("Size Mismatch");
                    reportRow++;
                    continue;
                }

                // Iterate through each cell in the range
                for (int r = 0; r < rows1; r++)
                {
                    for (int c = 0; c < cols1; c++)
                    {
                        int row1 = range1.FirstRow + r;
                        int col1 = range1.FirstColumn + c;
                        int row2 = range2.FirstRow + r;
                        int col2 = range2.FirstColumn + c;

                        Cell cell1 = range1.Worksheet.Cells[row1, col1];
                        Cell cell2 = range2.Worksheet.Cells[row2, col2];

                        string val1 = cell1.Value?.ToString() ?? string.Empty;
                        string val2 = cell2.Value?.ToString() ?? string.Empty;

                        if (!string.Equals(val1, val2, StringComparison.Ordinal))
                        {
                            // Record the difference
                            string address = cell1.Name; // address relative to its sheet
                            reportSheet.Cells[reportRow, 0].PutValue(rangeName);
                            reportSheet.Cells[reportRow, 1].PutValue(address);
                            reportSheet.Cells[reportRow, 2].PutValue(val1);
                            reportSheet.Cells[reportRow, 3].PutValue(val2);
                            reportSheet.Cells[reportRow, 4].PutValue("Different");
                            reportRow++;
                        }
                    }
                }
            }

            // Auto-fit columns for readability
            reportSheet.AutoFitColumns();

            // Save the report workbook
            const string reportFile = "NamedRangeDifferencesReport.xlsx";
            reportWb.Save(reportFile);
            Console.WriteLine($"Report saved to {reportFile}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
