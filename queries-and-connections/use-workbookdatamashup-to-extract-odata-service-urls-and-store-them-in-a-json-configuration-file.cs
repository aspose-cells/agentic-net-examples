// Title: Extract OData service URLs from an Excel workbook using Aspose.Cells Workbook.DataMashup and write them to a JSON configuration file (C#)
// AI Prompts: Create a C# console application that loads an .xlsx file with Aspose.Cells, uses Workbook.DataMashup (via reflection) to gather all ODataConnectionInfos ServiceUrl values, and outputs the list to a formatted JSON file. | Write a method that iterates over the ODataConnectionInfos collection of a workbook's DataMashup and returns a List<string> containing each OData service URL. | Generate code that serializes a collection of OData URLs into an indented JSON document using System.Text.Json and saves it to a specified path.
// Common Searches: how to read ODataConnectionInfos ServiceUrl from an Excel file using Aspose.Cells C# | Aspose.Cells Workbook.DataMashup reflection example for extracting OData URLs | C# code to export OData service URLs from an Excel workbook to JSON | save OData connection URLs to a configuration file with System.Text.Json | enumerate OData connections in an .xlsx using Aspose.Cells DataMashup
// Tags: Aspose.Cells Workbook.DataMashup ODataConnectionInfos extraction | C# reflection OData URLs from Excel workbook | serialize OData service URLs to JSON with System.Text.Json | export OData connection URLs to config file | extract OData URLs from .xlsx using Aspose.Cells

using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace ODataExtractorApp
{
    // Loads an Excel workbook, uses reflection to read ODataConnectionInfos from Workbook.DataMashup, collects each ServiceUrl, and writes the URLs to an indented JSON configuration file.
    class ODataExtractor
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "config.json";

                // Verify that the input workbook exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: Input file '{inputPath}' was not found.");
                    return;
                }

                // Load the workbook from the specified file
                Workbook workbook = new Workbook(inputPath);

                // Collect all OData service URLs from the workbook's DataMashup using reflection
                List<string> odataUrls = new List<string>();
                try
                {
                    var mashup = workbook.DataMashup;
                    if (mashup != null)
                    {
                        var collectionProp = mashup.GetType().GetProperty("ODataConnectionInfos");
                        if (collectionProp != null)
                        {
                            var collection = collectionProp.GetValue(mashup) as IEnumerable;
                            if (collection != null)
                            {
                                foreach (var connection in collection)
                                {
                                    var urlProp = connection.GetType().GetProperty("ServiceUrl");
                                    if (urlProp != null)
                                    {
                                        var url = urlProp.GetValue(connection) as string;
                                        if (!string.IsNullOrEmpty(url))
                                        {
                                            odataUrls.Add(url);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Warning: Unable to extract OData URLs via reflection. {ex.Message}");
                }

                // Prepare a simple configuration object to hold the URLs
                var config = new
                {
                    ODataServiceUrls = odataUrls
                };

                // Serialize the configuration to formatted JSON using System.Text.Json
                var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
                string json = JsonSerializer.Serialize(config, jsonOptions);

                // Write the JSON to a configuration file
                File.WriteAllText(outputPath, json);

                Console.WriteLine($"OData service URLs have been extracted to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
