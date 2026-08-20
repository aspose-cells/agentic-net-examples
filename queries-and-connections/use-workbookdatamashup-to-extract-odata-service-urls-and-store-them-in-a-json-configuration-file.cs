// Title: Extract OData URLs from Excel Power Query using Aspose.Cells Workbook.DataMashup (C#)
// Description: Loads an Excel workbook with Aspose.Cells, reads the DataMashup to enumerate PowerQueryFormulas, applies a regex to each formula definition to capture HTTP/HTTPS OData service URLs, removes duplicates, and writes the list to a formatted JSON configuration file. The workbook can be saved afterwards if needed.
// Keywords: Aspose.Cells | Workbook.DataMashup | C# | Power Query | OData URL extraction | Excel external connections | regex URL parsing | JSON export | DataMashup API | Excel automation
// Common Searches: Aspose.Cells extract OData URLs from Power Query | C# read Workbook.DataMashup formulas | How to list external OData connections in an Excel file | Save extracted URLs to JSON with Aspose.Cells | Regex to find URLs in PowerQueryFormula definitions
// Developer Intent: Retrieve all OData service endpoints referenced in an Excel workbook's Power Query formulas and store them in a JSON configuration file using Aspose.Cells.
// Use Cases: Create an inventory of external OData sources for compliance auditing. | Generate a JSON manifest for downstream data‑integration pipelines. | Validate endpoint URLs before refreshing queries to enforce security policies.
// AI Prompts: Write C# code that uses Aspose.Cells Workbook.DataMashup to collect unique OData URLs from Power Query formulas and output an indented JSON file. | Explain step‑by‑step how to apply a regular expression to PowerQueryFormula.FormulaDefinition to isolate HTTP/HTTPS URLs. | Suggest enhancements for handling duplicate URLs, trimming trailing characters, and logging extraction results.

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Text.RegularExpressions;
using Aspose.Cells;
using Aspose.Cells.QueryTables;

namespace ODataUrlExtractor
{
    // Loads an Excel workbook with Aspose.Cells, reads the DataMashup to enumerate PowerQueryFormulas, applies a regex to each formula definition to capture HTTP/HTTPS OData service URLs, removes duplicates, and writes the list to a formatted JSON configuration file. The workbook can be saved afterwards if needed.
    class Program
    {
        static void Main()
        {
            // Path to the workbook that contains Power Query (OData) formulas
            string sourcePath = "input.xlsx";

            // Load the workbook (create/load rule)
            Workbook workbook = new Workbook(sourcePath);

            // Access mashup data which holds Power Query formulas
            DataMashup mashup = workbook.DataMashup;

            // Collection to store discovered OData URLs
            List<string> odataUrls = new List<string>();

            if (mashup != null && mashup.PowerQueryFormulas != null)
            {
                // Iterate through each Power Query formula
                foreach (PowerQueryFormula formula in mashup.PowerQueryFormulas)
                {
                    // The formula definition may contain the OData service URL
                    string definition = formula.FormulaDefinition;

                    // Use a regular expression to extract URLs starting with http or https
                    foreach (Match match in Regex.Matches(definition, @"https?://[^\s'\""]+"))
                    {
                        string url = match.Value.TrimEnd(';', ')'); // clean trailing characters
                        if (!odataUrls.Contains(url))
                        {
                            odataUrls.Add(url);
                        }
                    }
                }
            }

            // Serialize the list of URLs to a JSON configuration file
            string jsonOutput = JsonSerializer.Serialize(odataUrls, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText("odataUrls.json", jsonOutput);

            // Optionally, save the workbook (save rule) if any modifications were made
            workbook.Save("output.xlsx");

            Console.WriteLine("Extraction complete. URLs saved to odataUrls.json");
        }
    }
}
