// Title: Compare Threaded Comment Authors in Two Excel Workbooks with Aspose.Cells (C#)
// Description: Loads two .xlsx files using Aspose.Cells, extracts every threaded comment author per cell, builds dictionaries of cell‑to‑author sets, and reports cells where the author lists differ or where a cell exists only in one workbook.
// Keywords: Aspose.Cells | threaded comments | comment author extraction | excel workbook comparison | C# example | compare Excel files | author differences | cell comment analysis
// Common Searches: how to compare threaded comment authors with Aspose.Cells | extract comment authors from Excel using C# | find cells with different comment authors in two workbooks | Aspose.Cells example for comment author comparison | C# code to diff Excel comment authors
// Developer Intent: Find cells whose threaded comment author sets are not identical between two Excel workbooks.
// Use Cases: Audit review trails by detecting added or removed comment authors after a document revision. | Generate a discrepancy report when merging spreadsheets from different contributors. | Validate that comment authors in a batch of workbooks conform to an approved list.
// AI Prompts: Refactor the ThreadedCommentComparer to export the comparison results to a CSV file. | Extend the sample to also compare the text content of threaded comments, not just the authors. | Show how to load each workbook with a distinct password using Aspose.Cells LoadOptions.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// Loads two .xlsx files using Aspose.Cells, extracts every threaded comment author per cell, builds dictionaries of cell‑to‑author sets, and reports cells where the author lists differ or where a cell exists only in one workbook.
class ThreadedCommentComparer
{
    static void Main(string[] args)
    {
        // Paths to the two workbooks to compare
        string file1 = "Workbook1.xlsx";
        string file2 = "Workbook2.xlsx";

        // Load the workbooks safely
        Workbook wb1 = null;
        Workbook wb2 = null;

        try
        {
            if (!File.Exists(file1))
                throw new FileNotFoundException($"File not found: {file1}");
            if (!File.Exists(file2))
                throw new FileNotFoundException($"File not found: {file2}");

            // Use LoadOptions in case the workbook is password‑protected; an empty password will be tried.
            var loadOptions = new LoadOptions { Password = "" };

            wb1 = new Workbook(file1, loadOptions);
            wb2 = new Workbook(file2, loadOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbooks: {ex.Message}");
            return;
        }

        try
        {
            // Extract threaded comment authors per cell for each workbook
            var comments1 = ExtractThreadedComments(wb1);
            var comments2 = ExtractThreadedComments(wb2);

            // Compare the extracted data and output differences in author attribution
            Console.WriteLine("Differences in threaded comment authors:");

            // Cells present in the first workbook
            foreach (var kvp in comments1)
            {
                string cell = kvp.Key;
                var authors1 = kvp.Value;

                if (comments2.TryGetValue(cell, out var authors2))
                {
                    // Authors that exist only in one of the workbooks
                    var onlyIn1 = new HashSet<string>(authors1, StringComparer.OrdinalIgnoreCase);
                    onlyIn1.ExceptWith(authors2);
                    var onlyIn2 = new HashSet<string>(authors2, StringComparer.OrdinalIgnoreCase);
                    onlyIn2.ExceptWith(authors1);

                    if (onlyIn1.Count > 0 || onlyIn2.Count > 0)
                    {
                        Console.WriteLine($"Cell {cell}:");
                        if (onlyIn1.Count > 0)
                            Console.WriteLine($"  Authors only in {file1}: {string.Join(", ", onlyIn1)}");
                        if (onlyIn2.Count > 0)
                            Console.WriteLine($"  Authors only in {file2}: {string.Join(", ", onlyIn2)}");
                    }
                }
                else
                {
                    Console.WriteLine($"Cell {cell} exists only in {file1} with authors: {string.Join(", ", authors1)}");
                }
            }

            // Cells present only in the second workbook
            foreach (var kvp in comments2)
            {
                if (!comments1.ContainsKey(kvp.Key))
                {
                    Console.WriteLine($"Cell {kvp.Key} exists only in {file2} with authors: {string.Join(", ", kvp.Value)}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during comparison: {ex.Message}");
        }
    }

    // Extracts a dictionary where the key is the cell address (e.g., "B2")
    // and the value is the set of author names that have threaded comments on that cell.
    static Dictionary<string, HashSet<string>> ExtractThreadedComments(Workbook wb)
    {
        var result = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in wb.Worksheets)
        {
            // If the worksheet has no comments, skip it
            if (sheet.Comments == null || sheet.Comments.Count == 0)
                continue;

            // Iterate through all comments in the worksheet
            foreach (Comment comment in sheet.Comments)
            {
                int row = comment.Row;
                int col = comment.Column;
                string cellName = sheet.Cells[row, col].Name;

                // Ensure a set exists for this cell
                if (!result.TryGetValue(cellName, out var authorSet))
                {
                    authorSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    result[cellName] = authorSet;
                }

                // Retrieve threaded comments for the cell (may be null)
                ThreadedCommentCollection threaded = sheet.Comments.GetThreadedComments(row, col);
                if (threaded == null)
                    continue;

                // Add each author name to the set
                foreach (ThreadedComment tc in threaded)
                {
                    if (tc?.Author?.Name != null && tc.Author.Name.Length > 0)
                    {
                        authorSet.Add(tc.Author.Name);
                    }
                }
            }
        }

        return result;
    }
}
