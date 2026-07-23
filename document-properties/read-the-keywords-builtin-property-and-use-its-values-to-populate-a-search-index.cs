// Title: Create a Keyword‑Based Cell Search Index from Built‑in Keywords with Aspose.Cells for .NET
// Description: This example shows how to read the BuiltInDocumentProperties.Keywords of a workbook, split the comma‑separated list, scan all populated cells, and build a case‑insensitive dictionary that maps each keyword to the addresses of matching cells. The index is printed to the console and the workbook can be saved.
// Keywords: Aspose.Cells read built‑in keywords | keyword search index Aspose.Cells | document properties .NET | cell address lookup by keyword | C# Aspose.Cells example | search spreadsheet by metadata | keyword‑driven navigation Excel
// Common Searches: how to read Keywords property in Aspose.Cells C# | build keyword to cell address index Aspose.Cells | search cells using document properties .NET | create searchable spreadsheet from metadata | Aspose.Cells keyword based lookup
// Developer Intent: Extract the workbook's Keywords property and generate a dictionary that links each keyword to the cells containing that term.
// Use Cases: Generate a quick navigation map for financial reports where keywords are stored in document properties. | Implement a custom search feature that highlights cells matching predefined keywords in large spreadsheets. | Create an automated indexing routine for regulatory filings that need fast keyword‑based retrieval.
// AI Prompts: Write C# code with Aspose.Cells to read the BuiltInDocumentProperties.Keywords, split it, and build a case‑insensitive dictionary of keyword → cell addresses. | Show how to output the keyword index to the console and then save the workbook as an .xlsx file. | Suggest ways to improve performance of the keyword indexing process for worksheets with tens of thousands of rows.

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsKeywordSearchIndex
{
    // This example shows how to read the BuiltInDocumentProperties.Keywords of a workbook, split the comma‑separated list, scan all populated cells, and build a case‑insensitive dictionary that maps each keyword to the addresses of matching cells. The index is printed to the console and the workbook can be saved.
    class Program
    {
        static void Main()
        {
            // -------------------------------------------------
            // 1. Create a new workbook and set Keywords property
            // -------------------------------------------------
            Workbook workbook = new Workbook();

            // Example keywords (comma‑separated)
            workbook.BuiltInDocumentProperties.Keywords = "Finance,Report,2024,Q1";

            // Add some sample data to the worksheet (optional)
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Finance Report Q1 2024");
            sheet.Cells["A2"].PutValue("Other data");

            // -------------------------------------------------
            // 2. Load the Keywords value and build a simple index
            // -------------------------------------------------
            string keywordsRaw = workbook.BuiltInDocumentProperties.Keywords;

            // Split by commas and trim whitespace
            string[] keywords = keywordsRaw.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < keywords.Length; i++)
                keywords[i] = keywords[i].Trim();

            // Example search index: keyword -> list of cell addresses containing the keyword
            Dictionary<string, List<string>> searchIndex = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);

            // Initialize index entries
            foreach (string kw in keywords)
                searchIndex[kw] = new List<string>();

            // Scan worksheet cells and associate matching keywords
            int maxRow = sheet.Cells.MaxDataRow;
            int maxCol = sheet.Cells.MaxDataColumn;

            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    var cell = sheet.Cells[row, col];
                    if (cell.Value == null) continue;

                    string cellText = cell.StringValue;
                    foreach (string kw in keywords)
                    {
                        if (cellText.IndexOf(kw, StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            string address = cell.Name; // e.g., "A1"
                            searchIndex[kw].Add(address);
                        }
                    }
                }
            }

            // -------------------------------------------------
            // 3. Display the constructed search index
            // -------------------------------------------------
            Console.WriteLine("Search Index based on Keywords:");
            foreach (var entry in searchIndex)
            {
                Console.WriteLine($"Keyword: {entry.Key}");
                if (entry.Value.Count > 0)
                {
                    Console.WriteLine("  Cells: " + string.Join(", ", entry.Value));
                }
                else
                {
                    Console.WriteLine("  Cells: (none found)");
                }
            }

            // -------------------------------------------------
            // 4. Save the workbook (optional)
            // -------------------------------------------------
            workbook.Save("KeywordSearchIndexDemo.xlsx");
        }
    }
}
