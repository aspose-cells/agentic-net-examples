using System;
using System.Collections.Generic;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Set the Keywords built‑in property
        workbook.BuiltInDocumentProperties.Keywords = "Aspose, Cells, Search, Example";

        // Save the workbook (create‑save lifecycle)
        string filePath = "keywords_demo.xlsx";
        workbook.Save(filePath);

        // Load the saved workbook (load‑read lifecycle)
        Workbook loadedWorkbook = new Workbook(filePath);

        // Read the Keywords property
        string keywords = loadedWorkbook.BuiltInDocumentProperties.Keywords;

        // Simple in‑memory search index: keyword -> list of document paths
        var searchIndex = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);

        // Split the keywords string by commas (or semicolons) and populate the index
        foreach (var token in keywords.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries))
        {
            string keyword = token.Trim();
            if (!searchIndex.ContainsKey(keyword))
                searchIndex[keyword] = new List<string>();

            searchIndex[keyword].Add(filePath);
        }

        // Output the constructed search index
        foreach (var entry in searchIndex)
        {
            Console.WriteLine($"Keyword: {entry.Key}");
            foreach (var doc in entry.Value)
            {
                Console.WriteLine($"  Document: {doc}");
            }
        }
    }
}