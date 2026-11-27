// Title: Extract OData.Feed service endpoint URLs from Power Query formulas in an Excel workbook using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells read Power Query formulas and retrieve OData.Feed URLs | How to get OData service endpoints from an Excel DataMashup with Aspose.Cells | Log OData.Feed connection strings from a workbook using Aspose.Cells .NET | Programmatically parse Power Query formulas to find OData URLs in C# | Extract OData.Feed URLs from Excel files via Aspose.Cells API
// Tags: Aspose.Cells DataMashup OData extraction | C# extract OData.Feed URLs from Excel | Power Query formula parsing with Aspose.Cells | log OData service endpoints in .NET | regex detection of OData.Feed URLs in workbook

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;
using Aspose.Cells.QueryTables;

namespace ODataConnectionExtractor
{
    // The example loads an Excel workbook with Aspose.Cells, accesses its DataMashup object, iterates through all Power Query formulas, applies a regular expression to locate OData.Feed calls, extracts the service endpoint URLs, and writes each discovered URL to the console along with a summary list.
    class Program
    {
        /// <param name="args">Command‑line arguments.</param>
        static void Main(string[] args)
        {
            try
            {
                // Validate input arguments
                if (args.Length != 1)
                {
                    Console.WriteLine("Usage: ODataConnectionExtractor <path-to-workbook>");
                    return;
                }

                string workbookPath = args[0];
                if (!File.Exists(workbookPath))
                {
                    Console.WriteLine($"Error: File not found – \"{workbookPath}\"");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(workbookPath);

                // Access the DataMashup object
                DataMashup mashup = workbook.DataMashup;
                if (mashup == null)
                {
                    Console.WriteLine("The workbook does not contain any DataMashup (Power Query) data.");
                    return;
                }

                // Retrieve all Power Query formulas
                PowerQueryFormulaCollection formulas = mashup.PowerQueryFormulas;
                if (formulas == null || formulas.Count == 0)
                {
                    Console.WriteLine("No Power Query formulas were found in the workbook.");
                    return;
                }

                // Prepare a list to hold discovered OData URLs
                List<string> odataUrls = new List<string>();

                // Regular expression to capture the URL argument of OData.Feed(...)
                // Handles both double‑quoted and single‑quoted strings.
                Regex odataRegex = new Regex(@"OData\.Feed\(\s*[""'](?<url>[^""']+)[""']", RegexOptions.Compiled | RegexOptions.IgnoreCase);

                // Iterate through each formula and look for OData.Feed usage
                foreach (PowerQueryFormula formula in formulas)
                {
                    if (string.IsNullOrEmpty(formula.FormulaDefinition))
                        continue;

                    Match match = odataRegex.Match(formula.FormulaDefinition);
                    if (match.Success)
                    {
                        string url = match.Groups["url"].Value;
                        odataUrls.Add(url);
                        Console.WriteLine($"Found OData connection in formula \"{formula.Name}\":");
                        Console.WriteLine($"  Service Endpoint URL: {url}");
                    }
                }

                if (odataUrls.Count == 0)
                {
                    Console.WriteLine("No OData.Feed connections were detected in any Power Query formula.");
                }
                else
                {
                    Console.WriteLine();
                    Console.WriteLine("Summary of extracted OData service endpoints:");
                    foreach (string url in odataUrls)
                    {
                        Console.WriteLine($"- {url}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An unexpected error occurred:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
