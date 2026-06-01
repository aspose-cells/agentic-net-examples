using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Cells;

namespace ThreadedCommentComparison
{
    class Program
    {
        static void Main(string[] args)
        {
            // Paths to the two workbooks to compare
            string workbookPath1 = "Workbook1.xlsx";
            string workbookPath2 = "Workbook2.xlsx";

            // Load the workbooks (lifecycle rule: use Workbook constructor)
            Workbook wb1 = new Workbook(workbookPath1);
            Workbook wb2 = new Workbook(workbookPath2);

            // Extract threaded comment author information from each workbook
            var commentsInfo1 = ExtractThreadedComments(wb1);
            var commentsInfo2 = ExtractThreadedComments(wb2);

            // Compare the extracted information and display differences
            CompareCommentAuthors(commentsInfo1, commentsInfo2);
        }

        /// <summary>
        /// Traverses all worksheets, cells and extracts the list of author names for each cell that contains threaded comments.
        /// The key is a fully qualified cell address in the form "SheetName!A1".
        /// </summary>
        private static Dictionary<string, List<string>> ExtractThreadedComments(Workbook workbook)
        {
            var result = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);

            // Iterate through each worksheet
            for (int wsIndex = 0; wsIndex < workbook.Worksheets.Count; wsIndex++)
            {
                Worksheet sheet = workbook.Worksheets[wsIndex];
                Cells cells = sheet.Cells;

                // Determine the used range to limit iteration
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                // Scan each cell within the used range
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        // Retrieve threaded comments for the current cell
                        ThreadedCommentCollection threadedComments = sheet.Comments.GetThreadedComments(row, col);
                        if (threadedComments != null && threadedComments.Count > 0)
                        {
                            // Build a key like "Sheet1!B2"
                            string cellAddress = cells[row, col].Name; // e.g., "B2"
                            string key = $"{sheet.Name}!{cellAddress}";

                            // Collect author names for all threaded comments in this cell
                            var authors = new List<string>();
                            foreach (ThreadedComment tc in threadedComments)
                            {
                                // Guard against null author (should not happen, but be safe)
                                if (tc.Author != null)
                                    authors.Add(tc.Author.Name);
                            }

                            // Store the list (order is not important for comparison)
                            result[key] = authors;
                        }
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Compares two dictionaries containing threaded comment author lists and prints differences.
        /// </summary>
        private static void CompareCommentAuthors(
            Dictionary<string, List<string>> dict1,
            Dictionary<string, List<string>> dict2)
        {
            // Union of all keys from both workbooks
            var allKeys = new HashSet<string>(dict1.Keys, StringComparer.OrdinalIgnoreCase);
            allKeys.UnionWith(dict2.Keys);

            foreach (var key in allKeys.OrderBy(k => k))
            {
                bool inFirst = dict1.TryGetValue(key, out var authors1);
                bool inSecond = dict2.TryGetValue(key, out var authors2);

                if (!inFirst)
                {
                    Console.WriteLine($"Cell {key} exists only in Workbook2 with authors: {string.Join(", ", authors2)}");
                }
                else if (!inSecond)
                {
                    Console.WriteLine($"Cell {key} exists only in Workbook1 with authors: {string.Join(", ", authors1)}");
                }
                else
                {
                    // Both workbooks have threaded comments in this cell; compare author sets
                    var set1 = new HashSet<string>(authors1, StringComparer.OrdinalIgnoreCase);
                    var set2 = new HashSet<string>(authors2, StringComparer.OrdinalIgnoreCase);

                    if (!set1.SetEquals(set2))
                    {
                        var onlyInFirst = set1.Except(set2);
                        var onlyInSecond = set2.Except(set1);

                        Console.WriteLine($"Difference in authors for cell {key}:");
                        if (onlyInFirst.Any())
                            Console.WriteLine($"  Only in Workbook1: {string.Join(", ", onlyInFirst)}");
                        if (onlyInSecond.Any())
                            Console.WriteLine($"  Only in Workbook2: {string.Join(", ", onlyInSecond)}");
                    }
                }
            }
        }
    }
}