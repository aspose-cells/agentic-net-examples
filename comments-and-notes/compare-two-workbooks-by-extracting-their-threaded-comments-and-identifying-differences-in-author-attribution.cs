// Title: Compare comment authors between two Excel workbooks using Aspose.Cells in C#
// Common Searches: Aspose.Cells C# compare comment authors between two .xlsx files | how to get comment author from a cell using Aspose.Cells | C# extract Excel comment author by cell address with Aspose.Cells | detect mismatched comment authors across worksheets using Aspose.Cells | compare workbook comments author metadata programmatically in .NET
// Tags: Aspose.Cells extract comment author | Aspose.Cells compare workbook comment metadata | C# map Excel comments by cell address | detect comment author differences Aspose.Cells | compare worksheet comment authors .NET

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The sample loads two Excel workbooks, iterates matching worksheets, builds dictionaries of comment authors keyed by cell address, and prints cells where a comment exists only in one workbook or where the author differs between the workbooks.
class WorkbookCommentComparer
{
    static void Main()
    {
        try
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

            // Load the two workbooks to compare
            Workbook wb1 = new Workbook(file1);
            Workbook wb2 = new Workbook(file2);

            // Iterate through each worksheet (by index) present in both workbooks
            int sheetCount = Math.Min(wb1.Worksheets.Count, wb2.Worksheets.Count);
            for (int sheetIndex = 0; sheetIndex < sheetCount; sheetIndex++)
            {
                Worksheet ws1 = wb1.Worksheets[sheetIndex];
                Worksheet ws2 = wb2.Worksheets[sheetIndex];
                string sheetName = ws1.Name;

                // Build dictionaries of comments keyed by cell address (e.g., "A1")
                var comments1 = GetCommentsByAddress(ws1);
                var comments2 = GetCommentsByAddress(ws2);

                // Union of all addresses that have comments in either workbook
                var allAddresses = new HashSet<string>(comments1.Keys, StringComparer.OrdinalIgnoreCase);
                allAddresses.UnionWith(comments2.Keys);

                // Compare author attribution for each address
                foreach (string address in allAddresses)
                {
                    comments1.TryGetValue(address, out string author1);
                    comments2.TryGetValue(address, out string author2);

                    if (author1 == null && author2 != null)
                    {
                        Console.WriteLine($"Comment only in Workbook2 at {sheetName}!{address}: Author = \"{author2}\"");
                    }
                    else if (author1 != null && author2 == null)
                    {
                        Console.WriteLine($"Comment only in Workbook1 at {sheetName}!{address}: Author = \"{author1}\"");
                    }
                    else if (!string.Equals(author1, author2, StringComparison.Ordinal))
                    {
                        Console.WriteLine($"Author difference at {sheetName}!{address}: Workbook1 Author = \"{author1}\", Workbook2 Author = \"{author2}\"");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }

    // Helper method to extract regular comments from a worksheet and map them by cell address
    private static Dictionary<string, string> GetCommentsByAddress(Worksheet worksheet)
    {
        var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        try
        {
            // Regular comments are stored in the Comments collection of the worksheet
            CommentCollection comments = worksheet.Comments;
            foreach (Comment comment in comments)
            {
                // Use row/column to build the cell address (e.g., "A1")
                string address = CellsHelper.CellIndexToName(comment.Row, comment.Column);
                string author = comment.Author;

                // If multiple comments exist for the same cell, keep the first encountered author
                if (!dict.ContainsKey(address))
                {
                    dict[address] = author;
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error while extracting comments: {ex.Message}");
        }
        return dict;
    }
}
