// Title: Extract OData service URLs from an Excel workbook using Workbook.DataMashup and save them to a JSON file (C# Aspose.Cells)
// Common Searches: C# Aspose.Cells extract OData URLs from Power Query formulas in an .xlsx file | How to read DataMashup information and get OData service links with Aspose.Cells | Save extracted OData service URLs to a JSON configuration file using Aspose.Cells | Retrieve OData endpoint URLs from Excel workbook DataMashup in .NET | Parse Power Query formula definitions for OData URLs with Aspose.Cells
// Tags: Workbook.DataMashup OData extraction | Power Query formula URL parsing | C# JSON serialization of URL collection | HashSet deduplication of service URLs | Aspose.Cells query tables connection discovery

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.QueryTables;

namespace ODataUrlExtractor
{
    // The example loads an Excel workbook with Aspose.Cells, accesses its DataMashup, iterates through each PowerQueryFormula and its items, extracts HTTP URLs containing "odata" using a regex, stores distinct URLs in a case‑insensitive HashSet, serializes the collection into an indented JSON object, and writes the result to "odata_urls.json".
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source workbook – you can pass it as a command‑line argument
            string workbookPath = args.Length > 0 ? args[0] : "input.xlsx";

            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Error: Workbook file not found at '{workbookPath}'.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(workbookPath);

                // Access mashup data
                DataMashup mashup = workbook.DataMashup;
                if (mashup == null)
                {
                    Console.WriteLine("No DataMashup information found in the workbook.");
                    return;
                }

                // Collect OData URLs
                HashSet<string> odataUrls = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                // Iterate through each Power Query formula
                foreach (PowerQueryFormula formula in mashup.PowerQueryFormulas)
                {
                    // 1. Check the full formula definition text
                    ExtractUrls(formula.FormulaDefinition, odataUrls);

                    // 2. Check each item value (some URLs are stored in items)
                    foreach (PowerQueryFormulaItem item in formula.PowerQueryFormulaItems)
                    {
                        if (!string.IsNullOrEmpty(item.Value))
                        {
                            ExtractUrls(item.Value, odataUrls);
                        }
                    }
                }

                if (odataUrls.Count == 0)
                {
                    Console.WriteLine("No OData service URLs were found in the workbook.");
                    return;
                }

                // Prepare JSON structure
                var jsonObject = new
                {
                    ODataUrls = odataUrls
                };

                // Serialize to JSON with indentation
                JsonSerializerOptions options = new JsonSerializerOptions
                {
                    WriteIndented = true
                };
                string json = JsonSerializer.Serialize(jsonObject, options);

                // Write to configuration file
                string outputPath = "odata_urls.json";
                File.WriteAllText(outputPath, json);
                Console.WriteLine($"Successfully extracted {odataUrls.Count} OData URL(s) to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        private static void ExtractUrls(string text, HashSet<string> collection)
        {
            if (string.IsNullOrEmpty(text))
                return;

            // Simple URL regex – captures until whitespace or a terminating quote
            const string pattern = @"https?://[^\s'""]+";
            foreach (Match match in Regex.Matches(text, pattern, RegexOptions.IgnoreCase))
            {
                string url = match.Value.TrimEnd(',', ';', ')', ']');
                if (url.IndexOf("odata", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    collection.Add(url);
                }
            }
        }
    }
}
