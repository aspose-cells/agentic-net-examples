// Title: Read the Keywords built‑in document property from an Excel workbook and create a keyword‑to‑cell address search index with Aspose.Cells for .NET
// AI Prompts: Extract the 'Keywords' built‑in document property from a workbook, split it into individual terms, and build a case‑insensitive dictionary that maps each keyword to the sheet name and cell address where it appears. | Traverse all worksheets and string cells using Aspose.Cells, compare each cell's text to the keyword list, and populate a search‑index dictionary of keyword → list of cell references.
// Common Searches: how to get the Keywords built‑in property from an Excel file using Aspose.Cells C# | create a keyword index of Excel cells with Aspose.Cells .NET | search Excel worksheet cells for keywords defined in document properties Aspose.Cells | C# dictionary mapping Excel keywords to cell addresses using Aspose.Cells | iterate over used range of worksheets in Aspose.Cells to find matching text
// Tags: Aspose.Cells built‑in document properties extraction | keyword‑to‑cell address dictionary Aspose.Cells | search index for Excel keywords C# | case‑insensitive cell text matching Aspose.Cells | iterate used range worksheets Aspose.Cells

using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Cells;

// The example loads an Excel workbook, reads the 'Keywords' built‑in document property, splits it into separate terms, then scans every worksheet's used range for string cells. Each occurrence of a keyword is added to a case‑insensitive dictionary that maps the keyword to the sheet name and cell address, producing a searchable index that is printed to the console.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the built‑in document properties collection
        var properties = workbook.BuiltInDocumentProperties;

        // Retrieve the "Keywords" property value (it may be null)
        string keywordsRaw = properties["Keywords"]?.Value?.ToString() ?? string.Empty;

        // Split the keywords string into individual terms.
        // Common delimiters are semicolon, comma, or newline.
        char[] delimiters = new char[] { ';', ',', '\n', '\r' };
        List<string> keywordList = keywordsRaw
            .Split(delimiters, StringSplitOptions.RemoveEmptyEntries)
            .Select(k => k.Trim())
            .Where(k => !string.IsNullOrEmpty(k))
            .ToList();

        // Populate a simple search index.
        // Here we use a Dictionary where each keyword maps to a list of cell addresses that contain it.
        Dictionary<string, List<string>> searchIndex = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);

        // Iterate through all worksheets and cells to index occurrences of each keyword.
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Get the used range of the sheet to limit the search.
            var cells = sheet.Cells;
            var maxRow = cells.MaxDataRow;
            var maxColumn = cells.MaxDataColumn;

            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxColumn; col++)
                {
                    var cell = cells[row, col];
                    if (cell.Type != CellValueType.IsString) continue;

                    string cellText = cell.StringValue;
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

        // Example: output the search index to console
        foreach (var entry in searchIndex)
        {
            Console.WriteLine($"Keyword: {entry.Key}");
            foreach (var location in entry.Value)
            {
                Console.WriteLine($"  Found in: {location}");
            }
        }

        // (Optional) Save the workbook if any modifications were made.
        // workbook.Save("output.xlsx");
    }
}
