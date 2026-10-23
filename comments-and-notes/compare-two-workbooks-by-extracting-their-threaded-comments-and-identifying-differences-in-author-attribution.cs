// Title: How to compare threaded comment authors between two Excel workbooks using Aspose.Cells for .NET (C#)
// Common Searches: C# code to extract threaded comment authors from an Excel workbook with Aspose.Cells | compare comment author lists of two .xlsx files using Aspose.Cells .NET | find cells with different comment authors in two workbooks using Aspose.Cells | Aspose.Cells example for detecting comment author differences across worksheets | threaded comment author comparison between Workbook1.xlsx and Workbook2.xlsx in C#
// Tags: extract threaded comment authors Aspose.Cells | compare comment author sets .NET | cell-level comment author difference detection | Aspose.Cells workbook comment comparison | C# threaded comment author analysis

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The sample loads two Excel workbooks, iterates each worksheet to collect the authors of threaded comments per cell, stores the data in a dictionary keyed by sheet name and cell address, and then compares the dictionaries to report cells where the set of comment authors differs between the two files.
class ThreadedCommentComparer
{
    static void Main()
    {
        const string file1 = "Workbook1.xlsx";
        const string file2 = "Workbook2.xlsx";

        // Verify that the input files exist to avoid FileNotFoundException
        if (!File.Exists(file1))
        {
            Console.WriteLine($"Error: File \"{file1}\" not found.");
            return;
        }

        if (!File.Exists(file2))
        {
            Console.WriteLine($"Error: File \"{file2}\" not found.");
            return;
        }

        Workbook wb1;
        Workbook wb2;

        try
        {
            // Load the two workbooks to compare
            wb1 = new Workbook(file1);
            wb2 = new Workbook(file2);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbooks: {ex.Message}");
            return;
        }

        Dictionary<CellKey, HashSet<string>> commentsInfo1;
        Dictionary<CellKey, HashSet<string>> commentsInfo2;

        try
        {
            // Extract comment author information from each workbook
            commentsInfo1 = ExtractComments(wb1);
            commentsInfo2 = ExtractComments(wb2);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error extracting comments: {ex.Message}");
            return;
        }

        // Compare the extracted information and report differences in author attribution
        Console.WriteLine("Differences in comment authors:");
        foreach (var key in commentsInfo1.Keys)
        {
            if (commentsInfo2.ContainsKey(key))
            {
                var authors1 = commentsInfo1[key];
                var authors2 = commentsInfo2[key];

                // Find authors that are present in one workbook but not the other
                var onlyInFirst = new HashSet<string>(authors1, StringComparer.OrdinalIgnoreCase);
                onlyInFirst.ExceptWith(authors2);

                var onlyInSecond = new HashSet<string>(authors2, StringComparer.OrdinalIgnoreCase);
                onlyInSecond.ExceptWith(authors1);

                if (onlyInFirst.Count > 0 || onlyInSecond.Count > 0)
                {
                    Console.WriteLine($"Sheet: {key.SheetName}, Cell: {key.CellName}");
                    if (onlyInFirst.Count > 0)
                        Console.WriteLine($"  Authors only in {file1}: {string.Join(", ", onlyInFirst)}");
                    if (onlyInSecond.Count > 0)
                        Console.WriteLine($"  Authors only in {file2}: {string.Join(", ", onlyInSecond)}");
                }
            }
        }
    }

    // Helper method to extract comment authors keyed by sheet and cell address
    private static Dictionary<CellKey, HashSet<string>> ExtractComments(Workbook workbook)
    {
        var result = new Dictionary<CellKey, HashSet<string>>();

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Access the collection of comments for the current worksheet
            CommentCollection comments = sheet.Comments;

            // Iterate through each comment
            foreach (Comment comment in comments)
            {
                // Determine the cell address the comment is attached to
                int row = comment.Row;       // zero‑based row index
                int column = comment.Column; // zero‑based column index
                string cellName = CellsHelper.CellIndexToName(row, column); // e.g., "B5"

                // Create a key that uniquely identifies the comment location
                var key = new CellKey(sheet.Name, cellName);

                // Ensure the dictionary entry exists
                if (!result.ContainsKey(key))
                {
                    result[key] = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                }

                // Add the author of the comment to the set (ignore null/empty authors)
                if (!string.IsNullOrEmpty(comment.Author))
                {
                    result[key].Add(comment.Author);
                }
            }
        }

        return result;
    }

    // Simple struct to act as a composite key for sheet name + cell address
    private struct CellKey : IEquatable<CellKey>
    {
        public string SheetName { get; }
        public string CellName { get; }

        public CellKey(string sheetName, string cellName)
        {
            SheetName = sheetName;
            CellName = cellName;
        }

        public bool Equals(CellKey other)
        {
            return string.Equals(SheetName, other.SheetName, StringComparison.OrdinalIgnoreCase) &&
                   string.Equals(CellName, other.CellName, StringComparison.OrdinalIgnoreCase);
        }

        public override bool Equals(object obj) => obj is CellKey other && Equals(other);

        public override int GetHashCode()
        {
            unchecked
            {
                int hashSheet = SheetName != null ? StringComparer.OrdinalIgnoreCase.GetHashCode(SheetName) : 0;
                int hashCell = CellName != null ? StringComparer.OrdinalIgnoreCase.GetHashCode(CellName) : 0;
                return (hashSheet * 397) ^ hashCell;
            }
        }
    }
}
