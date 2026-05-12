using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace ThreadedCommentComparison
{
    class Program
    {
        static void Main(string[] args)
        {
            // Paths to the two workbooks to compare
            string path1 = "Workbook1.xlsx";
            string path2 = "Workbook2.xlsx";

            // Load the workbooks (lifecycle rule: use Workbook constructor)
            Workbook wb1 = new Workbook(path1);
            Workbook wb2 = new Workbook(path2);

            // Extract threaded comment author information from each workbook
            var commentsInfo1 = ExtractThreadedComments(wb1);
            var commentsInfo2 = ExtractThreadedComments(wb2);

            // Compare the two dictionaries and report differences in author attribution
            Console.WriteLine("Differences in threaded comment authors between the two workbooks:");
            foreach (var key in commentsInfo1.Keys)
            {
                // If the same cell exists in the second workbook
                if (commentsInfo2.TryGetValue(key, out var authors2))
                {
                    var authors1 = commentsInfo1[key];
                    // Compare author sets (order independent)
                    var set1 = new HashSet<string>(authors1);
                    var set2 = new HashSet<string>(authors2);
                    if (!set1.SetEquals(set2))
                    {
                        Console.WriteLine($"Worksheet: {key.WorksheetName}, Cell: {key.CellName}");
                        Console.WriteLine($"  Workbook1 authors: {string.Join(", ", authors1)}");
                        Console.WriteLine($"  Workbook2 authors: {string.Join(", ", authors2)}");
                    }
                }
                else
                {
                    // Cell with threaded comments exists only in workbook1
                    Console.WriteLine($"Worksheet: {key.WorksheetName}, Cell: {key.CellName} exists only in Workbook1.");
                }
            }

            // Cells that exist only in workbook2
            foreach (var key in commentsInfo2.Keys)
            {
                if (!commentsInfo1.ContainsKey(key))
                {
                    Console.WriteLine($"Worksheet: {key.WorksheetName}, Cell: {key.CellName} exists only in Workbook2.");
                }
            }
        }

        // Helper method to extract threaded comment authors per worksheet/cell
        private static Dictionary<CellKey, List<string>> ExtractThreadedComments(Workbook workbook)
        {
            var result = new Dictionary<CellKey, List<string>>();

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Determine the used range to limit iteration
                int maxRow = sheet.Cells.MaxDataRow;
                int maxCol = sheet.Cells.MaxDataColumn;

                // Scan each cell within the used range
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        // Retrieve threaded comments for the current cell
                        ThreadedCommentCollection threadedComments = sheet.Comments.GetThreadedComments(row, col);
                        if (threadedComments != null && threadedComments.Count > 0)
                        {
                            var authors = new List<string>();
                            foreach (ThreadedComment tc in threadedComments)
                            {
                                // Author may be null if not set; guard against it
                                if (tc.Author != null)
                                    authors.Add(tc.Author.Name);
                            }

                            // Use cell address like "A1"
                            string cellName = CellsHelper.CellIndexToName(row, col);
                            var key = new CellKey(sheet.Name, cellName);
                            result[key] = authors;
                        }
                    }
                }
            }

            return result;
        }

        // Simple struct to act as a composite key for worksheet name + cell address
        private struct CellKey : IEquatable<CellKey>
        {
            public string WorksheetName { get; }
            public string CellName { get; }

            public CellKey(string worksheetName, string cellName)
            {
                WorksheetName = worksheetName;
                CellName = cellName;
            }

            public bool Equals(CellKey other)
            {
                return string.Equals(WorksheetName, other.WorksheetName, StringComparison.OrdinalIgnoreCase) &&
                       string.Equals(CellName, other.CellName, StringComparison.OrdinalIgnoreCase);
            }

            public override bool Equals(object obj) => obj is CellKey other && Equals(other);

            public override int GetHashCode()
            {
                unchecked
                {
                    return ((WorksheetName != null ? StringComparer.OrdinalIgnoreCase.GetHashCode(WorksheetName) : 0) * 397) ^
                           (CellName != null ? StringComparer.OrdinalIgnoreCase.GetHashCode(CellName) : 0);
                }
            }
        }
    }
}