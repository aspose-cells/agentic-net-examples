// Title: Detect formulas that reference cells outside the used range in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells find formulas that reference cells beyond the worksheet used range | how to list Excel formulas with out‑of‑range cell references using Aspose.Cells | detect invalid range references in a .xlsx file with Aspose.Cells C# example | scan workbook for formulas pointing to non‑existent rows or columns using Aspose.Cells | identify formulas that reference cells outside MaxDataRow or MaxDataColumn in C#
// Tags: out-of-range formula detection Aspose.Cells | Excel formula reference validation C# | used range analysis for formulas Aspose.Cells | regex cell address parsing C# | scan worksheet formulas for non‑existent cells | identify invalid range references workbook

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

// The example loads an Excel workbook with Aspose.Cells, determines each worksheet's used range via MaxDataRow and MaxDataColumn, iterates all formula cells, extracts single‑cell and range references using a regular expression, converts the addresses to zero‑based indices, and reports any reference whose row or column exceeds the used range, effectively detecting out‑of‑range formula references.
class DetectOutOfRangeFormulas
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: File \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook wb = new Workbook(inputPath);

            // Store information about problematic formulas
            List<string> issues = new List<string>();

            // Regular expression to capture cell references and ranges (e.g., A1, $B$2, C3:D10)
            Regex refRegex = new Regex(@"(\$?[A-Za-z]{1,3}\$?\d+)(:\$?[A-Za-z]{1,3}\$?\d+)?", RegexOptions.Compiled);

            // Iterate through each worksheet
            foreach (Worksheet sheet in wb.Worksheets)
            {
                // Determine the used range of the sheet (zero‑based indices)
                int maxRow = sheet.Cells.MaxDataRow;
                int maxCol = sheet.Cells.MaxDataColumn;

                // Scan all cells that contain formulas
                foreach (Cell cell in sheet.Cells)
                {
                    if (string.IsNullOrEmpty(cell.Formula))
                        continue;

                    // Find all references inside the formula
                    foreach (Match match in refRegex.Matches(cell.Formula))
                    {
                        // Single cell reference
                        if (string.IsNullOrEmpty(match.Groups[2].Value))
                        {
                            var (row, col) = CellAddressToIndices(match.Groups[1].Value);
                            if (row > maxRow || col > maxCol)
                            {
                                issues.Add($"{sheet.Name}!{cell.Name} references {match.Groups[1].Value}");
                                break; // No need to check further matches for this cell
                            }
                        }
                        else // Range reference (e.g., A1:B10)
                        {
                            var (firstRow, firstCol) = CellAddressToIndices(match.Groups[1].Value);
                            var (lastRow, lastCol) = CellAddressToIndices(match.Groups[2].Value.TrimStart(':'));

                            if (firstRow > maxRow || firstCol > maxCol ||
                                lastRow > maxRow || lastCol > maxCol)
                            {
                                issues.Add($"{sheet.Name}!{cell.Name} references range {match.Groups[1].Value}:{match.Groups[2].Value.TrimStart(':')}");
                                break;
                            }
                        }
                    }
                }
            }

            // Output the detection results
            if (issues.Count == 0)
            {
                Console.WriteLine("No out-of-range formula references found.");
            }
            else
            {
                Console.WriteLine("Out-of-range formula references detected:");
                foreach (string issue in issues)
                {
                    Console.WriteLine(issue);
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }

    // Converts an Excel cell address (e.g., "C5") to zero‑based row and column indices.
    private static (int row, int col) CellAddressToIndices(string address)
    {
        // Remove any absolute reference symbols ('$')
        string clean = address.Replace("$", "");

        // Split letters and numbers
        int i = 0;
        while (i < clean.Length && Char.IsLetter(clean[i]))
            i++;

        string colPart = clean.Substring(0, i);
        string rowPart = clean.Substring(i);

        int col = ColumnNameToIndex(colPart);
        int row = int.Parse(rowPart) - 1; // Excel rows are 1‑based

        return (row, col);
    }

    // Converts column letters (e.g., "AB") to a zero‑based column index.
    private static int ColumnNameToIndex(string columnName)
    {
        int sum = 0;
        foreach (char c in columnName.ToUpper())
        {
            sum *= 26;
            sum += (c - 'A' + 1);
        }
        return sum - 1; // Zero‑based
    }
}
