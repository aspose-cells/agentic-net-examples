using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsKeywordSearchIndex
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and set the Keywords built‑in property
            Workbook workbook = new Workbook();
            workbook.BuiltInDocumentProperties.Keywords = "Finance,Report,2024,Q1";

            // Save the workbook (lifecycle rule: save)
            string filePath = "KeywordWorkbook.xlsx";
            workbook.Save(filePath);

            // Load the saved workbook (lifecycle rule: load)
            Workbook loadedWorkbook = new Workbook(filePath);

            // Read the Keywords property
            string keywords = loadedWorkbook.BuiltInDocumentProperties.Keywords;

            // Prepare a simple in‑memory search index:
            // The index maps each keyword to a list of file paths that contain it.
            Dictionary<string, List<string>> searchIndex = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);

            // Split the keywords string (comma‑separated) and trim each entry
            if (!string.IsNullOrEmpty(keywords))
            {
                string[] keywordArray = keywords.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (string rawKeyword in keywordArray)
                {
                    string keyword = rawKeyword.Trim();
                    if (!searchIndex.ContainsKey(keyword))
                    {
                        searchIndex[keyword] = new List<string>();
                    }
                    searchIndex[keyword].Add(filePath);
                }
            }

            // Demonstrate searching the index
            Console.WriteLine("Search Index Contents:");
            foreach (var entry in searchIndex)
            {
                Console.WriteLine($"Keyword: '{entry.Key}' -> Files: {string.Join(", ", entry.Value)}");
            }

            // Example search: find all files containing the keyword "Report"
            string searchTerm = "Report";
            if (searchIndex.TryGetValue(searchTerm, out List<string> files))
            {
                Console.WriteLine($"\nFiles containing keyword '{searchTerm}': {string.Join(", ", files)}");
            }
            else
            {
                Console.WriteLine($"\nNo files found for keyword '{searchTerm}'.");
            }
        }
    }
}