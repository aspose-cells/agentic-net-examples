// Title: How to merge specific cell ranges in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells merge cells A1:D1 and B3:C4 example | write merged ranges to Excel file with Aspose.Cells .NET | preserve complex table layout when populating Excel workbook using Aspose.Cells C# | apply multiple merged cell areas programmatically in Aspose.Cells | Aspose.Cells C# merge cells after filling data rows
// Tags: merge cells Aspose.Cells C# | apply merged ranges Excel .NET | populate worksheet then merge cells | preserve table structure Aspose.Cells | write merged cells large workbook C#

using System;
using System.Collections.Generic;
using Aspose.Cells;

// The sample defines a MergedRange class, builds a two‑dimensional data list, creates a Workbook, fills the first worksheet with the data, merges the specified ranges (A1:D1 and B3:C4), and saves the file as MergedCellsStreaming.xlsx using Aspose.Cells for .NET.
public class MergedRange
{
    public int FirstRow { get; set; }
    public int FirstColumn { get; set; }
    public int TotalRows { get; set; }
    public int TotalColumns { get; set; }

    public MergedRange(int firstRow, int firstColumn, int totalRows, int totalColumns)
    {
        FirstRow = firstRow;
        FirstColumn = firstColumn;
        TotalRows = totalRows;
        TotalColumns = totalColumns;
    }
}

public class Program
{
    public static void Main()
    {
        try
        {
            // Sample data: each inner array represents a row
            var rows = new List<object[]>
            {
                new object[] { "Header1", "Header2", "Header3", "Header4" },
                new object[] { "A1", "B1", "C1", "D1" },
                new object[] { "A2", "B2", "C2", "D2" },
                new object[] { "A3", "B3", "C3", "D3" }
            };

            // Define merged ranges (same as original example)
            var mergedRanges = new List<MergedRange>
            {
                new MergedRange(0, 0, 1, 4), // Merge cells A1:D1
                new MergedRange(2, 1, 2, 2)  // Merge cells B3:C4
            };

            // Create a new workbook and get the first worksheet
            var workbook = new Workbook();
            var sheet = workbook.Worksheets[0];

            // Populate cells with data
            for (int r = 0; r < rows.Count; r++)
            {
                var row = rows[r];
                for (int c = 0; c < row.Length; c++)
                {
                    sheet.Cells[r, c].PutValue(row[c]);
                }
            }

            // Apply merged ranges
            foreach (var range in mergedRanges)
            {
                sheet.Cells.Merge(range.FirstRow, range.FirstColumn, range.TotalRows, range.TotalColumns);
            }

            // Save the workbook
            workbook.Save("MergedCellsStreaming.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
