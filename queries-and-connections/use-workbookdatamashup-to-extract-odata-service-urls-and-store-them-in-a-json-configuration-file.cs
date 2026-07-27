using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;
using Aspose.Cells.QueryTables;

namespace AsposeCellsODataExtractor
{
    // Simple model to hold extracted OData URLs for JSON serialization
    public class ODataConfig
    {
        public List<string> Urls { get; set; } = new List<string>();
    }

    public class Program
    {
        public static void Main()
        {
            // Path to the source workbook that may contain Power Query (OData) connections
            string sourcePath = "source.xlsx";

            // Load the workbook (using the standard Aspose.Cells load rule)
            Workbook workbook = new Workbook(sourcePath);

            // Prepare a collection to hold discovered URLs
            ODataConfig config = new ODataConfig();

            // -----------------------------------------------------------------
            // 1. Extract URLs from Power Query formulas (DataMashup.PowerQueryFormulas)
            // -----------------------------------------------------------------
            DataMashup mashup = workbook.DataMashup;
            if (mashup != null && mashup.PowerQueryFormulas != null)
            {
                foreach (PowerQueryFormula formula in mashup.PowerQueryFormulas)
                {
                    // The formula definition often contains the OData service URL.
                    // Example: let Source = OData.Feed("https://services.odata.org/V4/Northwind/Northwind.svc/Products")
                    string definition = formula.FormulaDefinition;
                    if (!string.IsNullOrEmpty(definition))
                    {
                        // Simple extraction: look for text between double quotes after OData.Feed(
                        int startIdx = definition.IndexOf("OData.Feed(", StringComparison.OrdinalIgnoreCase);
                        if (startIdx >= 0)
                        {
                            startIdx = definition.IndexOf('\"', startIdx);
                            if (startIdx >= 0)
                            {
                                int endIdx = definition.IndexOf('\"', startIdx + 1);
                                if (endIdx > startIdx)
                                {
                                    string url = definition.Substring(startIdx + 1, endIdx - startIdx - 1);
                                    if (!config.Urls.Contains(url))
                                    {
                                        config.Urls.Add(url);
                                    }
                                }
                            }
                        }
                    }
                }
            }

            // -----------------------------------------------------------------
            // 2. Extract URLs from external web query connections (DataConnections)
            // -----------------------------------------------------------------
            ExternalConnectionCollection connections = workbook.DataConnections;
            if (connections != null)
            {
                foreach (ExternalConnection conn in connections)
                {
                    // Only WebQueryConnection objects have a Url property that can point to OData services
                    if (conn is WebQueryConnection webConn && !string.IsNullOrEmpty(webConn.Url))
                    {
                        if (!config.Urls.Contains(webConn.Url))
                        {
                            config.Urls.Add(webConn.Url);
                        }
                    }
                }
            }

            // -----------------------------------------------------------------
            // 3. Serialize the collected URLs to a JSON configuration file
            // -----------------------------------------------------------------
            string jsonOutputPath = "ODataConfig.json";
            var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
            string json = JsonSerializer.Serialize(config, jsonOptions);
            File.WriteAllText(jsonOutputPath, json);

            // -----------------------------------------------------------------
            // 4. (Optional) Save the workbook if any modifications were made
            // -----------------------------------------------------------------
            // In this scenario we only read data, so saving is not required.
            // However, to demonstrate the save rule, we could save a copy.
            string destPath = "source_copy.xlsx";
            workbook.Save(destPath);

            Console.WriteLine($"Extracted {config.Urls.Count} OData URL(s) and saved to '{jsonOutputPath}'.");
        }
    }
}