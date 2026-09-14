// Title: Extract and compare threaded comment authors across two Excel workbooks using Aspose.Cells for .NET
// AI Prompts: Load two .xlsx files, build a dictionary of Sheet!Cell → ordered list of threaded comment authors, and output cells where the author sequences differ. | Write a C# routine that iterates through all worksheets, gathers each threaded comment's author name, and returns a map suitable for cross‑workbook comparison. | Enhance the comparer to also list cells that contain threaded comments in only one of the two workbooks.
// Common Searches: how to extract threaded comment authors from an Excel file using Aspose.Cells C# | compare comment author lists between two .xlsx workbooks in .NET | C# code to find cells with different threaded comment authors in two spreadsheets | list cells that have comments in one workbook but not the other using Aspose.Cells
// Tags: extract threaded comments Aspose.Cells C# | compare comment authors between workbooks | map sheet cell to comment author list | detect mismatched threaded comment authors .NET | identify exclusive comments in one workbook Aspose.Cells

using System;
using System.Collections.Generic;
using Aspose.Cells;

// The sample loads two Excel workbooks, extracts every threaded comment together with its author name per cell, stores the data in dictionaries keyed by "Sheet!Cell", and then compares the dictionaries to report cells where the ordered author lists differ or where comments exist only in one of the workbooks.
class ThreadedCommentComparer
{
    static void Main(string[] args)
    {
        // Expect two workbook file paths as arguments
        if (args.Length < 2)
        {
            Console.WriteLine("Usage: ThreadedCommentComparer <workbook1.xlsx> <workbook2.xlsx>");
            return;
        }

        string filePath1 = args[0];
        string filePath2 = args[1];

        // Load the workbooks (lifecycle rule)
        Workbook wb1 = new Workbook(filePath1);
        Workbook wb2 = new Workbook(filePath2);

        // Extract threaded comments with their authors from each workbook
        var commentsMap1 = ExtractThreadedComments(wb1);
        var commentsMap2 = ExtractThreadedComments(wb2);

        // Compare the two dictionaries and report differences in author attribution
        foreach (var kvp in commentsMap1)
        {
            string cellKey = kvp.Key;                     // e.g., Sheet1!A1
            List<string> authors1 = kvp.Value;

            if (commentsMap2.TryGetValue(cellKey, out var authors2))
            {
                if (!AreAuthorListsEqual(authors1, authors2))
                {
                    Console.WriteLine($"Author difference in {cellKey}:");
                    Console.WriteLine($"  Workbook1: {string.Join(", ", authors1)}");
                    Console.WriteLine($"  Workbook2: {string.Join(", ", authors2)}");
                }
            }
            else
            {
                Console.WriteLine($"Threaded comments exist only in Workbook1 for {cellKey}.");
            }
        }

        // Cells that have threaded comments only in Workbook2
        foreach (var cellKey in commentsMap2.Keys)
        {
            if (!commentsMap1.ContainsKey(cellKey))
            {
                Console.WriteLine($"Threaded comments exist only in Workbook2 for {cellKey}.");
            }
        }
    }

    // Extracts a map: "SheetName!CellAddress" -> list of author names (ordered)
    static Dictionary<string, List<string>> ExtractThreadedComments(Workbook workbook)
    {
        var map = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);

        foreach (Worksheet sheet in workbook.Worksheets)
        {
            CommentCollection comments = sheet.Comments;

            // Iterate through all comments in the worksheet
            for (int i = 0; i < comments.Count; i++)
            {
                Comment comment = comments[i];

                // Each comment may contain a collection of threaded comments
                foreach (ThreadedComment tc in comment.ThreadedComments)
                {
                    // Convert row/column indices to Excel cell name (e.g., A1)
                    string cellAddress = CellsHelper.CellIndexToName(tc.Row, tc.Column);
                    string key = $"{sheet.Name}!{cellAddress}";

                    if (!map.TryGetValue(key, out var authorList))
                    {
                        authorList = new List<string>();
                        map[key] = authorList;
                    }

                    // Store the author name; fallback to "Unknown" if null
                    authorList.Add(tc.Author?.Name ?? "Unknown");
                }
            }
        }

        return map;
    }

    // Simple ordered comparison of two author name lists
    static bool AreAuthorListsEqual(List<string> list1, List<string> list2)
    {
        if (list1.Count != list2.Count) return false;
        for (int i = 0; i < list1.Count; i++)
        {
            if (!string.Equals(list1[i], list2[i], StringComparison.Ordinal))
                return false;
        }
        return true;
    }
}
