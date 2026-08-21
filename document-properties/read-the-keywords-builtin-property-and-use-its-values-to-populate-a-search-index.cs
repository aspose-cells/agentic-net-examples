// Title: C# – Read Built‑In Keywords Property and Build a Worksheet Search Index with Aspose.Cells
// Description: Demonstrates how to create an Excel workbook, set its BuiltInDocumentProperties.Keywords, reload the file, parse the comma‑separated keywords, and generate a case‑insensitive dictionary that maps each keyword to worksheets whose names contain that keyword. The index is printed to the console for quick verification.
// Keywords: Aspose.Cells C# read Keywords property | Excel built‑in document properties | keyword‑based worksheet search | metadata driven index Aspose.Cells | C# dictionary from Excel keywords | case insensitive worksheet lookup | search index from Excel document properties | Aspose.Cells .NET example | Excel metadata search tutorial | global finance reporting Excel keywords
// Common Searches: How to read the Keywords built‑in property with Aspose.Cells in C# | Create a keyword to worksheet map from an Excel file using .NET | Build a simple search index from Excel document properties | Iterate worksheets and match names with keywords Aspose.Cells | Aspose.Cells example for extracting document metadata
// Developer Intent: Extract the Keywords built‑in property from an Excel workbook and use the values to generate a keyword‑to‑worksheet lookup dictionary.
// Use Cases: Provide fast navigation in reporting dashboards by linking finance‑related keywords (e.g., Finance, Q1) to the corresponding worksheets. | Implement a metadata‑driven search feature that filters available sheets without scanning the entire workbook each time. | Create a reusable index that can be serialized (JSON, XML) for external services such as search APIs or document management systems.
// AI Prompts: Generate C# code with Aspose.Cells that reads the Keywords built‑in property from an existing workbook and builds a case‑insensitive dictionary mapping each keyword to matching worksheet names. | Explain how to extend the keyword index to support partial matches, synonyms, and export the dictionary to JSON for use in a web service. | Show how to update the Keywords property programmatically and refresh the search index without reloading the workbook from disk.

using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Cells;

namespace AsposeCellsKeywordSearchIndex
{
    // Demonstrates how to create an Excel workbook, set its BuiltInDocumentProperties.Keywords, reload the file, parse the comma‑separated keywords, and generate a case‑insensitive dictionary that maps each keyword to worksheets whose names contain that keyword. The index is printed to the console for quick verification.
    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a workbook and set the Keywords property (demo purpose)
            // -----------------------------------------------------------------
            Workbook workbook = new Workbook();
            workbook.BuiltInDocumentProperties.Keywords = "Finance,Report,2024,Q1";

            // Save the workbook to a temporary file
            string filePath = "KeywordDemo.xlsx";
            workbook.Save(filePath);

            // -----------------------------------------------------------------
            // 2. Load the workbook and read the Keywords property
            // -----------------------------------------------------------------
            Workbook loadedWorkbook = new Workbook(filePath);
            string keywords = loadedWorkbook.BuiltInDocumentProperties.Keywords;

            // -----------------------------------------------------------------
            // 3. Populate a simple search index using the keywords
            //    For demonstration, the index maps each keyword to a list of
            //    worksheet names that contain the keyword in their name.
            // -----------------------------------------------------------------
            // Split the comma‑separated keywords and trim whitespace
            List<string> keywordList = keywords
                .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(k => k.Trim())
                .Where(k => !string.IsNullOrEmpty(k))
                .ToList();

            // Build the index: keyword -> list of worksheet names
            Dictionary<string, List<string>> searchIndex = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);

            foreach (string kw in keywordList)
            {
                // Initialize the list for each keyword
                searchIndex[kw] = new List<string>();
            }

            // Iterate through worksheets and associate them with matching keywords
            foreach (Worksheet sheet in loadedWorkbook.Worksheets)
            {
                string sheetName = sheet.Name;
                foreach (string kw in keywordList)
                {
                    if (sheetName.IndexOf(kw, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        searchIndex[kw].Add(sheetName);
                    }
                }
            }

            // -----------------------------------------------------------------
            // 4. Output the populated search index
            // -----------------------------------------------------------------
            Console.WriteLine("Search Index based on Keywords:");
            foreach (var entry in searchIndex)
            {
                string sheets = entry.Value.Count > 0 ? string.Join(", ", entry.Value) : "(no matching sheets)";
                Console.WriteLine($"Keyword: '{entry.Key}' -> Worksheets: {sheets}");
            }
        }
    }
}
