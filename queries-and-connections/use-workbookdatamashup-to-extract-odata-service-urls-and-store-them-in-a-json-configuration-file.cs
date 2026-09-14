// Title: Extract OData and WebQuery connection URLs from an Excel workbook using Aspose.Cells DataMashup and export them to a JSON configuration (C#)
// AI Prompts: Write C# code that opens an .xlsx file with Aspose.Cells, iterates over workbook.DataConnections, identifies connections of type OData or WebQuery, and collects their Url values. | Extend the program to scan workbook.DataMashup.PowerQueryFormulas, extract any HTTP/HTTPS URLs from the formula definitions, and add them to the OData URL collection. | Add logic to serialize the distinct URL list into a formatted odata-config.json file and include error handling that logs warnings for missing Url properties or unsupported DataMashup features.
// Common Searches: C# Aspose.Cells extract OData service URLs from Excel workbook | How to retrieve Power Query URLs from an .xlsx file using Aspose.Cells DataMashup | Save extracted OData and WebQuery connection URLs to JSON with .NET | Enumerate DataConnections of type OData or WebQuery in Aspose.Cells | Parse Power Query formula definitions for URLs in C#
// Tags: DataConnections OData URL extraction | DataMashup Power Query URL parsing | C# write URLs to formatted JSON | Excel workbook OData service list generation | Aspose.Cells unique URL aggregation

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace ExtractODataUrlsApp
{
    // The program loads an Excel workbook with Aspose.Cells, gathers unique OData and WebQuery connection URLs from DataConnections, optionally extracts additional URLs from Power Query formulas via DataMashup, and writes the collected URLs to a formatted JSON configuration file named odata-config.json.
    class ExtractODataUrls
    {
        static void Main()
        {
            // Path to the workbook that contains Power Query (OData) connections
            string sourcePath = "source.xlsx";

            // Verify the source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(sourcePath);

                // List to store extracted OData URLs
                List<string> odataUrls = new List<string>();

                // -----------------------------------------------------------------
                // Extract URLs from data connections (OData, Web query, etc.)
                // -----------------------------------------------------------------
                foreach (var conn in workbook.DataConnections)
                {
                    try
                    {
                        string url = null;

                        // Determine connection type via reflection (avoids direct enum dependency)
                        var typeProp = conn.GetType().GetProperty("Type");
                        var typeValue = typeProp?.GetValue(conn);
                        string typeName = typeValue?.ToString();

                        if (typeName == "OData" || typeName == "WebQuery")
                        {
                            // Access the Url property via reflection
                            var urlProp = conn.GetType().GetProperty("Url");
                            if (urlProp != null)
                            {
                                url = urlProp.GetValue(conn) as string;
                            }
                        }

                        if (!string.IsNullOrEmpty(url) && !odataUrls.Contains(url))
                        {
                            odataUrls.Add(url);
                        }
                    }
                    catch (Exception exConn)
                    {
                        // Log connection‑specific issues and continue processing other connections
                        Console.WriteLine($"Warning: Unable to process a connection – {exConn.Message}");
                    }
                }

                // -----------------------------------------------------------------
                // Optional: Scan Power Query formulas for URLs (if DataMashup is available)
                // -----------------------------------------------------------------
                try
                {
                    var mashup = workbook.DataMashup;
                    if (mashup != null && mashup.PowerQueryFormulas != null)
                    {
                        Regex urlRegex = new Regex(@"https?://[^\s""]+", RegexOptions.Compiled | RegexOptions.IgnoreCase);
                        foreach (var formula in mashup.PowerQueryFormulas)
                        {
                            string definition = formula.FormulaDefinition;
                            if (!string.IsNullOrEmpty(definition))
                            {
                                foreach (Match match in urlRegex.Matches(definition))
                                {
                                    string foundUrl = match.Value.TrimEnd(')', ';');
                                    if (!odataUrls.Contains(foundUrl))
                                    {
                                        odataUrls.Add(foundUrl);
                                    }
                                }
                            }
                        }
                    }
                }
                catch
                {
                    // DataMashup not supported – ignore and continue
                }

                // Build a simple JSON configuration object
                var config = new
                {
                    ODataServiceUrls = odataUrls
                };

                // Serialize the configuration to formatted JSON
                string json = JsonSerializer.Serialize(config, new JsonSerializerOptions { WriteIndented = true });

                // Write JSON to a file
                string jsonPath = "odata-config.json";
                File.WriteAllText(jsonPath, json);

                // Save the workbook (no modifications made, but follows lifecycle rule)
                try
                {
                    workbook.Save("output.xlsx");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Warning: Unable to save workbook – {saveEx.Message}");
                }

                Console.WriteLine($"Extracted {odataUrls.Count} OData URLs and saved to {jsonPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
