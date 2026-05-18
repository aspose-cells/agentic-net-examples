using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.QueryTables;
using Aspose.Cells.ExternalConnections;

class ExtractODataUrls
{
    static void Main()
    {
        // Path to the workbook that contains Power Query / external connections
        string sourcePath = "source.xlsx";

        // Load the workbook (create/load rule)
        Workbook workbook = new Workbook(sourcePath);

        // Collection to store discovered OData service URLs
        List<string> odataUrls = new List<string>();

        // ----- Extract URLs from Power Query formulas -----
        DataMashup mashup = workbook.DataMashup;
        if (mashup != null && mashup.PowerQueryFormulas != null)
        {
            foreach (PowerQueryFormula formula in mashup.PowerQueryFormulas)
            {
                // Example formula: let Source = OData.Feed("https://service/odata") ...
                string definition = formula.FormulaDefinition;
                int start = definition.IndexOf("OData.Feed(\"", StringComparison.OrdinalIgnoreCase);
                if (start >= 0)
                {
                    start += "OData.Feed(\"".Length;
                    int end = definition.IndexOf("\")", start);
                    if (end > start)
                    {
                        string url = definition.Substring(start, end - start);
                        odataUrls.Add(url);
                    }
                }
            }
        }

        // ----- Extract URLs from external WebQuery connections -----
        foreach (ExternalConnection conn in workbook.DataConnections)
        {
            if (conn is WebQueryConnection webConn && !string.IsNullOrEmpty(webConn.Url))
            {
                // Treat any URL containing "odata" (case‑insensitive) as an OData service URL
                if (webConn.Url.IndexOf("odata", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    odataUrls.Add(webConn.Url);
                }
            }
        }

        // Remove duplicate entries
        var distinctUrls = new HashSet<string>(odataUrls);

        // Serialize the URL list to a formatted JSON string
        string json = JsonSerializer.Serialize(distinctUrls, new JsonSerializerOptions { WriteIndented = true });

        // Write the JSON configuration file (save rule)
        string jsonPath = "ODataConfig.json";
        File.WriteAllText(jsonPath, json);

        // Save the workbook unchanged (demonstrates save rule)
        workbook.Save("output.xlsx");

        Console.WriteLine($"Extracted {distinctUrls.Count} OData URLs to '{jsonPath}'.");
    }
}