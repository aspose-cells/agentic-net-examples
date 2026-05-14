using System;
using System.Collections.Generic;
using Aspose.Cells;

class KeywordIndexer
{
    static void Main()
    {
        // Create a new workbook (create rule)
        Workbook workbook = new Workbook();

        // Set the Keywords built‑in property
        workbook.BuiltInDocumentProperties.Keywords = "Aspose, Cells, Search, Index";

        // Save the workbook (save rule)
        string filePath = "KeywordDemo.xlsx";
        workbook.Save(filePath);

        // Load the saved workbook (load rule)
        Workbook loadedWorkbook = new Workbook(filePath);

        // Read the Keywords property
        string keywords = loadedWorkbook.BuiltInDocumentProperties.Keywords;

        // Simple in‑memory search index: keyword -> list of document paths
        var searchIndex = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);

        // Split the comma‑separated keywords and add them to the index
        foreach (string rawKeyword in keywords.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries))
        {
            string keyword = rawKeyword.Trim();
            if (!searchIndex.ContainsKey(keyword))
                searchIndex[keyword] = new List<string>();
            searchIndex[keyword].Add(filePath);
        }

        // Output the constructed index
        Console.WriteLine("Search index contents:");
        foreach (var entry in searchIndex)
        {
            Console.WriteLine($"Keyword: {entry.Key}");
            foreach (var doc in entry.Value)
                Console.WriteLine($"  Document: {doc}");
        }
    }
}