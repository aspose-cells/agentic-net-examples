// Title: Filter XmlMap Objects by Namespace Using LINQ in Aspose.Cells for .NET (C#)
// Description: Creates a Workbook, adds XmlMap entries with distinct DataBinding URLs, defines a target namespace substring, and runs a LINQ query on workbook.Worksheets.XmlMaps to return maps whose DataBinding.Url contains the substring (case‑insensitive). The matching map names are displayed and the workbook is saved.
// Keywords: Aspose.Cells | C# | .NET | LINQ | XmlMap | namespace filter | DataBinding URL | XML schema lookup | Workbook XmlMaps query | search XML maps
// Common Searches: Aspose.Cells LINQ query on XmlMaps collection | filter XmlMap by namespace C# | find XML maps with specific schema URL Aspose.Cells | search XmlMap DataBinding.Url for substring | C# example to list XmlMaps containing a namespace
// Developer Intent: Retrieve every XmlMap in a workbook whose DataBinding URL includes a specified namespace string.
// Use Cases: Select only the XML maps that belong to a particular schema before importing data. | Validate that a workbook contains expected XML maps by checking namespace patterns. | Generate an audit report of XML maps that match a given namespace.
// AI Prompts: Write a C# method that accepts a Workbook and a namespace fragment, then returns the names of XmlMap objects whose DataBinding.Url contains that fragment using LINQ. | Show how to modify the LINQ expression for an exact, case‑sensitive match on the namespace URL. | Provide code that handles the situation where no XmlMap matches the namespace and logs a warning.

using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a Workbook, adds XmlMap entries with distinct DataBinding URLs, defines a target namespace substring, and runs a LINQ query on workbook.Worksheets.XmlMaps to return maps whose DataBinding.Url contains the substring (case‑insensitive). The matching map names are displayed and the workbook is saved.
    public class XmlMapNamespaceQueryDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Add XML maps with different namespaces (using URLs as placeholders)
                int mapIndex1 = workbook.Worksheets.XmlMaps.Add("https://example.com/schema1.xsd");
                XmlMap map1 = workbook.Worksheets.XmlMaps[mapIndex1];
                map1.Name = "MapWithExampleNamespace";

                int mapIndex2 = workbook.Worksheets.XmlMaps.Add("https://otherdomain.com/schema2.xsd");
                XmlMap map2 = workbook.Worksheets.XmlMaps[mapIndex2];
                map2.Name = "MapWithoutExampleNamespace";

                // Define the namespace (or part of it) to search for
                string targetNamespace = "example.com";

                // LINQ query on the XmlMaps collection to find maps whose DataBinding URL contains the target namespace
                List<XmlMap> matchingMaps = workbook.Worksheets.XmlMaps
                    .Cast<XmlMap>()
                    .Where(m => m.DataBinding != null &&
                                m.DataBinding.Url != null &&
                                m.DataBinding.Url.Contains(targetNamespace, StringComparison.OrdinalIgnoreCase))
                    .ToList();

                // Output the names of the matching maps
                Console.WriteLine($"XML maps containing namespace \"{targetNamespace}\":");
                foreach (XmlMap xmlMap in matchingMaps)
                {
                    Console.WriteLine($"- {xmlMap.Name}");
                }

                // Save the workbook (optional, just to demonstrate lifecycle usage)
                workbook.Save("XmlMapNamespaceQueryDemo.xlsx");
                Console.WriteLine("Workbook saved as XmlMapNamespaceQueryDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
