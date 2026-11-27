// Title: Use LINQ to locate Excel XML maps that contain a given namespace with Aspose.Cells for .NET
// Common Searches: c# linq to get xml maps containing a specific namespace in an Excel workbook using Aspose.Cells | how to list xml map names that match a namespace with Aspose.Cells for .NET | search xml maps by namespace in a workbook with Aspose.Cells and LINQ
// Tags: LINQ filter on XmlMapCollection | search XML map namespace Aspose.Cells | retrieve XML map names by namespace C# | case-insensitive XML map lookup | enumerate Excel workbook XML maps

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Cells;

namespace AsposeCellsXmlMapLinqDemo
{
    // The example loads an Excel file with Aspose.Cells, accesses the workbook's XmlMapCollection, and applies a LINQ query that checks each map's DataBinding.Url for a case‑insensitive match to a specified namespace, then outputs the matching map names and root elements.
    class Program
    {
        /// <param name="args">
        /// args[0] - Path to the Excel file to load.
        /// args[1] - Namespace string to search for (case‑insensitive).
        /// </param>
        static void Main(string[] args)
        {
            // Validate arguments
            if (args == null || args.Length < 2)
            {
                Console.WriteLine("Usage: AsposeCellsXmlMapLinqDemo <excelFilePath> <namespaceToFind>");
                return;
            }

            string excelPath = args[0];
            string targetNamespace = args[1];

            // Ensure the file exists before attempting to load it
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"Error: The file \"{excelPath}\" does not exist.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(excelPath);

                // Get the collection of XML maps
                XmlMapCollection xmlMaps = workbook.Worksheets.XmlMaps;

                // Use LINQ to find maps whose data binding URL contains the target namespace
                // (Aspose.Cells does not expose the XSD string directly; we use the URL as a proxy)
                List<XmlMap> matchingMaps = xmlMaps
                    .Cast<XmlMap>()
                    .Where(map =>
                    {
                        // DataBinding may be null for some maps
                        if (map.DataBinding == null) return false;

                        // Use the Url property as a searchable string
                        string url = map.DataBinding.Url;
                        return !string.IsNullOrEmpty(url) &&
                               url.IndexOf(targetNamespace, StringComparison.OrdinalIgnoreCase) >= 0;
                    })
                    .ToList();

                // Output results
                if (matchingMaps.Count == 0)
                {
                    Console.WriteLine($"No XML maps containing the namespace \"{targetNamespace}\" were found.");
                }
                else
                {
                    Console.WriteLine($"Found {matchingMaps.Count} XML map(s) containing the namespace \"{targetNamespace}\":");
                    foreach (XmlMap map in matchingMaps)
                    {
                        Console.WriteLine($"- Map Name: \"{map.Name}\", Root Element: \"{map.RootElementName}\"");
                    }
                }
            }
            catch (Exception ex)
            {
                // Report any unexpected errors
                Console.WriteLine("An error occurred while processing the workbook:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
