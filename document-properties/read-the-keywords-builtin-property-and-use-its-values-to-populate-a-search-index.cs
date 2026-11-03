// Title: Create a keyword‑to‑cell address index by reading the Excel "Keywords" built‑in property with Aspose.Cells for .NET
// Common Searches: how to read the Keywords built‑in property from an Excel file using Aspose.Cells in C# | build a keyword search index of cell locations in a .xlsx workbook with Aspose.Cells | extract document property keywords and map them to cell addresses in .NET | scan all worksheets for keyword matches and list cell references using Aspose.Cells | C# code to split Excel Keywords property and create a dictionary of matches
// Tags: read built‑in Keywords property Aspose.Cells | create keyword dictionary from Excel cells C# | parse Keywords property delimiters Aspose.Cells | associate keywords with worksheet cell addresses .NET | enumerate worksheets and cells Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsKeywordIndexer
{
    // The example loads an Excel workbook with Aspose.Cells, reads the built‑in "Keywords" document property, splits it into individual terms, scans every cell across all worksheets, and builds a case‑insensitive dictionary that maps each keyword to the sheet name and cell address where it appears.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";

            // Verify that the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: File '{inputPath}' not found.");
                return;
            }

            Workbook workbook = null;
            try
            {
                // Load the workbook (no password). If a password is required, an exception will be thrown.
                workbook = new Workbook(inputPath);
            }
            catch (CellsException ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error while loading workbook: {ex.Message}");
                return;
            }

            try
            {
                // Access built‑in document properties
                var builtInProps = workbook.BuiltInDocumentProperties;

                // Retrieve the "Keywords" property value (may be null or empty)
                string keywordsValue = builtInProps["Keywords"]?.ToString() ?? string.Empty;

                // Split the keywords using common delimiters
                char[] delimiters = new char[] { ';', ',', '\n', '\r' };
                string[] keywords = keywordsValue.Split(delimiters, StringSplitOptions.RemoveEmptyEntries);

                // Trim whitespace and collect non‑empty keywords
                List<string> keywordList = new List<string>();
                foreach (var kw in keywords)
                {
                    string trimmed = kw.Trim();
                    if (!string.IsNullOrEmpty(trimmed))
                        keywordList.Add(trimmed);
                }

                // Build a simple search index: keyword -> list of cell addresses
                var searchIndex = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);

                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    var cells = sheet.Cells;
                    int maxRow = cells.MaxDataRow;
                    int maxCol = cells.MaxDataColumn;

                    for (int row = 0; row <= maxRow; row++)
                    {
                        for (int col = 0; col <= maxCol; col++)
                        {
                            var cell = cells[row, col];
                            if (cell.Value == null) continue;

                            string cellText = cell.Value.ToString();

                            foreach (string keyword in keywordList)
                            {
                                if (cellText.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    string address = cell.Name; // e.g., "A1"
                                    if (!searchIndex.ContainsKey(keyword))
                                        searchIndex[keyword] = new List<string>();

                                    searchIndex[keyword].Add($"{sheet.Name}!{address}");
                                }
                            }
                        }
                    }
                }

                // Output the constructed search index
                foreach (var entry in searchIndex)
                {
                    Console.WriteLine($"Keyword: {entry.Key}");
                    foreach (var location in entry.Value)
                    {
                        Console.WriteLine($"  Found in: {location}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during processing: {ex.Message}");
            }
        }
    }
}
