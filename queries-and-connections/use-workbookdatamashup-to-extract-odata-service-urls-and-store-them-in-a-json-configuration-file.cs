using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;
using Aspose.Cells.QueryTables;

namespace ODataUrlExtractor
{
    class Program
    {
        static void Main()
        {
            // Path to the source workbook that may contain Power Query formulas or web query connections
            string sourcePath = "source.xlsx";

            // Load the workbook (using the standard Aspose.Cells load rule)
            Workbook workbook = new Workbook(sourcePath);

            // Collection to hold discovered OData service URLs
            List<string> odataUrls = new List<string>();

            // -----------------------------------------------------------------
            // 1. Extract URLs from Power Query formulas (DataMashup.PowerQueryFormulas)
            // -----------------------------------------------------------------
            DataMashup mashup = workbook.DataMashup;
            if (mashup != null && mashup.PowerQueryFormulas != null)
            {
                foreach (PowerQueryFormula formula in mashup.PowerQueryFormulas)
                {
                    // The formula definition may contain an OData URL; use a regex to find http/https URLs
                    string definition = formula.FormulaDefinition;
                    if (!string.IsNullOrEmpty(definition))
                    {
                        foreach (Match match in Regex.Matches(definition, @"https?://[^\s""]+"))
                        {
                            odataUrls.Add(match.Value);
                        }
                    }
                }
            }

            // -----------------------------------------------------------------
            // 2. Extract URLs from external web query connections (Workbook.DataConnections)
            // -----------------------------------------------------------------
            foreach (ExternalConnection connection in workbook.DataConnections)
            {
                if (connection is WebQueryConnection webConn && !string.IsNullOrEmpty(webConn.Url))
                {
                    odataUrls.Add(webConn.Url);
                }
            }

            // Remove duplicate entries
            odataUrls = odataUrls.Distinct().ToList();

            // -----------------------------------------------------------------
            // 3. Serialize the URL list to a JSON configuration file
            // -----------------------------------------------------------------
            string jsonOutput = JsonSerializer.Serialize(odataUrls, new JsonSerializerOptions { WriteIndented = true });
            string jsonPath = "odataUrls.json";
            File.WriteAllText(jsonPath, jsonOutput);

            Console.WriteLine($"Extracted {odataUrls.Count} OData URL(s) and saved to '{jsonPath}'.");
        }
    }
}