// Title: C# – Compare Threaded Comment Authors in Two Excel Workbooks Using Aspose.Cells
// Description: Loads two .xlsx files with Aspose.Cells for .NET, scans each worksheet’s used range, extracts every threaded comment author per cell, builds a SheetName!CellAddress → author list map, and then highlights cells that exist only in one workbook or have mismatched author sets. Results are printed to the console.
// Keywords: Aspose.Cells threaded comments C# | compare Excel comment authors | extract threaded comment authors .NET | Excel workbook comment comparison | Aspose.Cells get threaded comments
// Common Searches: how to compare threaded comment authors with Aspose.Cells | list comment authors per cell C# Aspose.Cells | find differences in Excel comment authors between two files | extract and compare Excel threaded comments programmatically
// Developer Intent: Retrieve author names from threaded comments in two workbooks and report cells with missing or differing authors.
// Use Cases: Verify that reviewers' comments in a master spreadsheet match those in a submitted version before release. | Audit author changes in comment threads across spreadsheet revisions for compliance. | Identify cells that contain comments in only one of two compared workbooks.
// AI Prompts: Generate a C# method that returns a Dictionary<string, List<string>> mapping "Sheet!Cell" to distinct threaded comment author names using Aspose.Cells. | Write a function that accepts two such dictionaries and outputs cells with mismatched author sets or comments present in a single workbook. | Suggest ways to speed up threaded comment extraction for very large worksheets when using Aspose.Cells.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// Loads two .xlsx files with Aspose.Cells for .NET, scans each worksheet’s used range, extracts every threaded comment author per cell, builds a SheetName!CellAddress → author list map, and then highlights cells that exist only in one workbook or have mismatched author sets. Results are printed to the console.
class ThreadedCommentComparer
{
    // Extracts threaded comment authors per cell address from a workbook
    static Dictionary<string, List<string>> GetCommentAuthors(Workbook wb)
    {
        var result = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);

        // Iterate through all worksheets
        for (int wsIndex = 0; wsIndex < wb.Worksheets.Count; wsIndex++)
        {
            Worksheet ws = wb.Worksheets[wsIndex];
            CommentCollection comments = ws.Comments;

            // Determine the used range to limit scanning
            int maxRow = ws.Cells.MaxDataRow;
            int maxCol = ws.Cells.MaxDataColumn;

            // Scan each cell in the used range
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    // Retrieve threaded comments for the current cell
                    ThreadedCommentCollection tcCollection = comments.GetThreadedComments(row, col);
                    if (tcCollection == null || tcCollection.Count == 0)
                        continue; // No threaded comments in this cell

                    // Build a unique key: SheetName!CellName (e.g., Sheet1!B2)
                    string cellName = ws.Cells[row, col].Name;
                    string key = $"{ws.Name}!{cellName}";

                    // Collect author names (distinct)
                    var authors = new List<string>();
                    foreach (ThreadedComment tc in tcCollection)
                    {
                        if (tc?.Author != null && !string.IsNullOrEmpty(tc.Author.Name))
                            authors.Add(tc.Author.Name);
                    }

                    // Store the list (keep duplicates if needed)
                    result[key] = authors;
                }
            }
        }

        return result;
    }

    // Compares two dictionaries and prints differences in author attribution
    static void CompareAndReport(
        Dictionary<string, List<string>> dict1,
        Dictionary<string, List<string>> dict2,
        string wb1Name,
        string wb2Name)
    {
        // Cells present in workbook1
        foreach (var kvp in dict1)
        {
            string cell = kvp.Key;
            List<string> authors1 = kvp.Value;

            if (!dict2.TryGetValue(cell, out List<string> authors2))
            {
                Console.WriteLine($"Cell {cell} has threaded comments only in {wb1Name} (authors: {string.Join(", ", authors1)})");
                continue;
            }

            // Compare author sets (order‑insensitive)
            var set1 = new HashSet<string>(authors1);
            var set2 = new HashSet<string>(authors2);

            if (!set1.SetEquals(set2))
            {
                Console.WriteLine($"Cell {cell} author mismatch:");
                Console.WriteLine($"  {wb1Name}: {string.Join(", ", set1)}");
                Console.WriteLine($"  {wb2Name}: {string.Join(", ", set2)}");
            }
        }

        // Cells present only in workbook2
        foreach (var kvp in dict2)
        {
            if (!dict1.ContainsKey(kvp.Key))
            {
                Console.WriteLine($"Cell {kvp.Key} has threaded comments only in {wb2Name} (authors: {string.Join(", ", kvp.Value)})");
            }
        }
    }

    static void Main()
    {
        const string file1 = "Workbook1.xlsx";
        const string file2 = "Workbook2.xlsx";

        try
        {
            // Verify files exist before loading
            if (!File.Exists(file1))
            {
                Console.WriteLine($"Error: File '{file1}' not found.");
                return;
            }

            if (!File.Exists(file2))
            {
                Console.WriteLine($"Error: File '{file2}' not found.");
                return;
            }

            // Load the two workbooks to compare
            Workbook wb1 = new Workbook(file1);
            Workbook wb2 = new Workbook(file2);

            // Extract threaded comment author information
            var commentsInfo1 = GetCommentAuthors(wb1);
            var commentsInfo2 = GetCommentAuthors(wb2);

            // Report differences
            CompareAndReport(commentsInfo1, commentsInfo2, file1, file2);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
