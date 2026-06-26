using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Cells;

namespace AsposeCellsKeywordIndexDemo
{
    class Program
    {
        static void Main()
        {
            // ---------- Create a new workbook and set Keywords ----------
            Workbook workbook = new Workbook(); // create
            workbook.BuiltInDocumentProperties.Keywords = "Aspose, Cells, Search, Index, Demo";
            workbook.Save("keywords_demo.xlsx"); // save

            // ---------- Load the workbook ----------
            Workbook loadedWorkbook = new Workbook("keywords_demo.xlsx"); // load

            // ---------- Read Keywords property ----------
            string keywords = loadedWorkbook.BuiltInDocumentProperties.Keywords;

            // ---------- Populate a simple in‑memory search index ----------
            // The index maps each keyword to a list of document identifiers (here just the file name)
            var searchIndex = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);

            // Split the keywords string by commas, trim whitespace, and add to the index
            foreach (string rawKeyword in keywords.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries))
            {
                string keyword = rawKeyword.Trim();
                if (!searchIndex.ContainsKey(keyword))
                {
                    searchIndex[keyword] = new List<string>();
                }
                // For demonstration we associate the keyword with the current workbook file name
                searchIndex[keyword].Add("keywords_demo.xlsx");
            }

            // ---------- Display the populated index ----------
            Console.WriteLine("Search Index Contents:");
            foreach (var entry in searchIndex)
            {
                string docs = string.Join(", ", entry.Value.Distinct());
                Console.WriteLine($"Keyword: '{entry.Key}' -> Documents: {docs}");
            }
        }
    }
}